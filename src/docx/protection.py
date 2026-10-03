"""Read-only analysis of document protection and editable permission ranges.

This module interprets the protection rules recorded in a ``.docx`` file. It
combines the document-protection settings in ``settings.xml``
(``w:documentProtection``) with the permission-range markers
(``w:permStart``/``w:permEnd``) in the document body to report which content
is read-only and which content falls inside an editable exception range for a
given user or group identity.

The analysis is strictly read-only: it neither verifies the protection
password nor modifies the document in any way.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterator, Sequence

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.oxml.xmlchemy import BaseOxmlElement

# Editor-group values defined by ST_EdGrp in the OOXML schema.
KNOWN_EDITOR_GROUPS = (
    "none",
    "everyone",
    "administrators",
    "contributors",
    "editors",
    "owners",
    "current",
)


class ProtectionMode(enum.Enum):
    """Editing-restriction mode declared by w:documentProtection/@w:edit."""

    NONE = "none"
    READ_ONLY = "readOnly"
    COMMENTS = "comments"
    TRACKED_CHANGES = "trackedChanges"
    FORMS = "forms"


class SegmentStatus(enum.Enum):
    """Editability of a content segment for the analyzed identity."""

    NOT_PROTECTED = "not_protected"
    """Document does not enforce read-only protection; no exception needed."""

    EDITABLE_EXCEPTION = "editable_exception"
    """Read-only is enforced but a permission range grants this identity edit."""

    RESTRICTED = "restricted"
    """Read-only is enforced and no permission range grants this identity edit."""


class IssueKind(enum.Enum):
    """Kinds of malformed permission markup reported by the analysis."""

    UNPAIRED_PERM_START = "unpaired_perm_start"
    UNPAIRED_PERM_END = "unpaired_perm_end"
    UNKNOWN_EDITOR_GROUP = "unknown_editor_group"
    MISSING_EDITOR = "missing_editor"
    INVALID_COLUMN_RANGE = "invalid_column_range"


@dataclass(frozen=True)
class DocumentProtection:
    """Parsed ``w:documentProtection`` settings."""

    mode: ProtectionMode
    enforced: bool
    raw_edit: "str | None"
    raw_enforcement: "str | None"

    @property
    def read_only_enforced(self) -> bool:
        """True when read-only protection is actively enforced.

        Only in this case do permStart/permEnd exception ranges determine
        editability; otherwise the whole document is editable.
        """
        return self.mode is ProtectionMode.READ_ONLY and self.enforced


@dataclass(frozen=True)
class MarkerRef:
    """Traceable reference to a ``w:permStart`` or ``w:permEnd`` marker."""

    perm_id: "str | None"
    location: str
    """Document-order path such as body/p[2] or body/tbl[0]/tr[1]/tc[2]/p[0]."""


@dataclass(frozen=True)
class PermissionRange:
    """A ``w:permStart``/``w:permEnd`` pair matched by ``w:id``."""

    perm_id: "str | None"
    ed: "str | None"
    ed_grp: "str | None"
    col_first: "int | None"
    col_last: "int | None"
    start: MarkerRef
    end: MarkerRef
    valid: bool = True
    """False when the range carries defects (unknown group, invalid columns,
    missing editor). Invalid ranges never grant editing permission."""

    def grants(self, user_name: "str | None", user_groups: Sequence[str]) -> bool:
        """True if this range grants editing to the given identity."""
        if not self.valid:
            return False
        granted = False
        if self.ed is not None:
            granted = user_name is not None and user_name == self.ed
        if self.ed_grp is not None:
            if self.ed_grp == "everyone":
                granted = True
            elif self.ed_grp == "none":
                pass
            elif self.ed_grp in KNOWN_EDITOR_GROUPS:
                granted = granted or self.ed_grp in user_groups
            # unknown groups were reported as issues and never grant
        return granted


@dataclass(frozen=True)
class Segment:
    """A maximal run of text sharing the same active permission ranges."""

    location: str
    """Path of the containing paragraph, e.g. body/tbl[0]/tr[0]/tc[1]/p[0]."""

    start_offset: int
    end_offset: int
    text: str
    status: SegmentStatus
    active_perm_ids: "tuple[str, ...]"
    """Ids of all permission ranges covering this segment (traceability)."""

    granting_perm_ids: "tuple[str, ...]"
    """Subset of active_perm_ids that grant editing to the analyzed identity."""


@dataclass(frozen=True)
class ProtectionIssue:
    """A malformed or suspicious permission construct found in the document."""

    kind: IssueKind
    perm_id: "str | None"
    location: str
    message: str


@dataclass
class ProtectionReport:
    """Result of :func:`analyze_protection`."""

    protection: DocumentProtection
    user_name: "str | None"
    user_groups: "tuple[str, ...]"
    segments: "list[Segment]" = field(default_factory=list)
    permission_ranges: "list[PermissionRange]" = field(default_factory=list)
    issues: "list[ProtectionIssue]" = field(default_factory=list)

    @property
    def read_only_enforced(self) -> bool:
        return self.protection.read_only_enforced

    @property
    def editable_segments(self) -> "list[Segment]":
        """Segments the analyzed identity may edit (unprotected or exception)."""
        return [
            s
            for s in self.segments
            if s.status in (SegmentStatus.NOT_PROTECTED, SegmentStatus.EDITABLE_EXCEPTION)
        ]

    @property
    def restricted_segments(self) -> "list[Segment]":
        """Segments that remain read-only for the analyzed identity."""
        return [s for s in self.segments if s.status is SegmentStatus.RESTRICTED]


# ---------------------------------------------------------------------------
# internals
# ---------------------------------------------------------------------------

_W_PERM_START = qn("w:permStart")
_W_PERM_END = qn("w:permEnd")
_W_P = qn("w:p")
_W_TBL = qn("w:tbl")
_W_TR = qn("w:tr")
_W_TC = qn("w:tc")
_W_T = qn("w:t")
_W_TAB = qn("w:tab")
_W_BR = qn("w:br")
_W_CR = qn("w:cr")


@dataclass(frozen=True)
class _TableCtx:
    """Position of the content currently being walked within a table."""

    table_idx: int
    row_idx: int
    col_idx: int
    grid_cols: int
    row_key: "tuple[int, int]"


@dataclass(frozen=True)
class _Event:
    kind: str  # "perm_start" | "perm_end" | "text" | "row_end"
    path: str
    element: "BaseOxmlElement | None" = None
    text: str = ""
    tbl_ctx: "_TableCtx | None" = None


def _iter_block_events(
    el: "BaseOxmlElement", path: str, tbl_ctx: "_TableCtx | None"
) -> Iterator[_Event]:
    """Yield permission-marker and text events under `el` in document order."""
    counters: "dict[str, int]" = {}
    for child in el:
        tag = child.tag
        if tag == _W_PERM_START:
            yield _Event("perm_start", path, element=child, tbl_ctx=tbl_ctx)
        elif tag == _W_PERM_END:
            yield _Event("perm_end", path, element=child, tbl_ctx=tbl_ctx)
        elif tag == _W_TBL:
            idx = counters.get("tbl", 0)
            counters["tbl"] = idx + 1
            yield from _iter_table_events(child, path + "/tbl[%d]" % idx, idx)
        elif tag in (_W_P, _W_TR, _W_TC):
            name = tag.rsplit("}", 1)[1]
            idx = counters.get(name, 0)
            counters[name] = idx + 1
            yield from _iter_block_events(child, path + "/%s[%d]" % (name, idx), tbl_ctx)
        elif tag == _W_T:
            yield _Event("text", path, element=child, text=child.text or "", tbl_ctx=tbl_ctx)
        elif tag == _W_TAB:
            yield _Event("text", path, element=child, text="\t", tbl_ctx=tbl_ctx)
        elif tag in (_W_BR, _W_CR):
            yield _Event("text", path, element=child, text="\n", tbl_ctx=tbl_ctx)
        else:
            # descend into runs, hyperlinks, sdt, etc. without changing the path
            yield from _iter_block_events(child, path, tbl_ctx)


def _iter_table_events(tbl: "BaseOxmlElement", path: str, table_idx: int) -> Iterator[_Event]:
    tbl_grid = tbl.find(qn("w:tblGrid"))
    grid_cols = len(tbl_grid.findall(qn("w:gridCol"))) if tbl_grid is not None else 0
    row_idx = 0
    for child in tbl:
        if child.tag == _W_TR:
            col_idx = 0
            for tc in child.findall(_W_TC):
                span = 1
                span_el = tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan"))
                if span_el is not None:
                    try:
                        span = max(1, int(span_el.get(qn("w:val"), "1")))
                    except ValueError:
                        span = 1
                ctx = _TableCtx(table_idx, row_idx, col_idx, grid_cols, (id(tbl), row_idx))
                cell_path = path + "/tr[%d]/tc[%d]" % (row_idx, col_idx)
                yield from _iter_block_events(tc, cell_path, ctx)
                col_idx += span
            end_ctx = _TableCtx(table_idx, row_idx, col_idx, grid_cols, (id(tbl), row_idx))
            yield _Event("row_end", path + "/tr[%d]" % row_idx, tbl_ctx=end_ctx)
            row_idx += 1
        elif child.tag == _W_PERM_START:
            yield _Event("perm_start", path, element=child, tbl_ctx=None)
        elif child.tag == _W_PERM_END:
            yield _Event("perm_end", path, element=child, tbl_ctx=None)


@dataclass
class _PermStartRec:
    perm_id: "str | None"
    ed: "str | None"
    ed_grp: "str | None"
    col_first: "int | None"
    col_last: "int | None"
    location: str
    element_key: int
    valid: bool = True


def _parse_on_off(value: "str | None") -> bool:
    return value in ("1", "true", "on")


def _parse_document_protection(document: "Document") -> DocumentProtection:
    try:
        settings_el = document.settings.element
    except Exception:
        settings_el = None
    prot_el = settings_el.find(qn("w:documentProtection")) if settings_el is not None else None
    if prot_el is None:
        return DocumentProtection(ProtectionMode.NONE, False, None, None)
    raw_edit = prot_el.get(qn("w:edit"))
    raw_enforcement = prot_el.get(qn("w:enforcement"))
    try:
        mode = ProtectionMode(raw_edit) if raw_edit is not None else ProtectionMode.NONE
    except ValueError:
        mode = ProtectionMode.NONE
    return DocumentProtection(mode, _parse_on_off(raw_enforcement), raw_edit, raw_enforcement)


def _validate_perm_start(
    el: "BaseOxmlElement",
    location: str,
    tbl_ctx: "_TableCtx | None",
    issues: "list[ProtectionIssue]",
) -> _PermStartRec:
    """Validate a ``w:permStart`` marker, recording defects as issues.

    Defective markers (unknown editor group, missing editor, invalid column
    range) are still paired normally but flagged ``valid=False`` so they
    never grant editing permission; the analysis never infers intent.
    """
    perm_id = el.get(qn("w:id"))
    ed = el.get(qn("w:ed"))
    ed_grp = el.get(qn("w:edGrp"))
    col_first_raw = el.get(qn("w:colFirst"))
    col_last_raw = el.get(qn("w:colLast"))
    valid = True

    if ed is None and ed_grp is None:
        valid = False
        issues.append(
            ProtectionIssue(
                IssueKind.MISSING_EDITOR,
                perm_id,
                location,
                "permStart specifies neither w:ed nor w:edGrp; it grants no one",
            )
        )
    elif ed_grp is not None and ed_grp not in KNOWN_EDITOR_GROUPS:
        valid = False
        issues.append(
            ProtectionIssue(
                IssueKind.UNKNOWN_EDITOR_GROUP,
                perm_id,
                location,
                "unknown w:edGrp value %r; range grants no one" % (ed_grp,),
            )
        )

    col_first = col_last = None
    if col_first_raw is not None or col_last_raw is not None:
        col_ok, reason = True, ""
        if col_first_raw is None or col_last_raw is None:
            col_ok, reason = False, "w:colFirst and w:colLast must appear together"
        else:
            try:
                col_first, col_last = int(col_first_raw), int(col_last_raw)
            except ValueError:
                col_ok, reason = False, "w:colFirst/w:colLast are not integers"
                col_first = col_last = None
            else:
                if col_first < 0 or col_last < 0:
                    col_ok, reason = False, "column indexes must be non-negative"
                elif col_first > col_last:
                    col_ok, reason = (
                        False,
                        "w:colFirst (%d) exceeds w:colLast (%d)" % (col_first, col_last),
                    )
                elif tbl_ctx is None:
                    col_ok, reason = False, "column-scoped permission outside any table"
                    col_first = col_last = None
                elif tbl_ctx.grid_cols and col_last >= tbl_ctx.grid_cols:
                    col_ok, reason = (
                        False,
                        "w:colLast (%d) exceeds table column count (%d)"
                        % (col_last, tbl_ctx.grid_cols),
                    )
        if not col_ok:
            valid = False
            issues.append(
                ProtectionIssue(
                    IssueKind.INVALID_COLUMN_RANGE,
                    perm_id,
                    location,
                    "invalid column range: %s; range grants no one" % reason,
                )
            )
            col_first = col_last = None

    return _PermStartRec(
        perm_id=perm_id,
        ed=ed,
        ed_grp=ed_grp,
        col_first=col_first,
        col_last=col_last,
        location=location,
        element_key=id(el),
        valid=valid,
    )


def analyze_protection(
    document: "Document",
    user_name: "str | None" = None,
    user_groups: Sequence[str] = (),
) -> ProtectionReport:
    """Analyze document protection and editable exception ranges for an identity.

    `document` is a python-docx |Document| object. `user_name` is matched
    against ``w:ed`` and `user_groups` against ``w:edGrp`` on ``w:permStart``
    markers. The analysis is read-only: the document is never modified and
    the protection password is neither needed nor verified.

    Returns a |ProtectionReport| whose ``segments`` partition the body text
    (paragraphs and table cells) into restricted and editable spans, each
    traceable to the ``w:permStart``/``w:permEnd`` markers that govern it.
    Malformed markup (unpaired markers, unknown groups, invalid column
    ranges) is reported in ``issues`` and never silently treated as granting
    access.
    """
    protection = _parse_document_protection(document)
    report = ProtectionReport(protection, user_name, tuple(user_groups))

    body = document.element.body
    events = list(_iter_block_events(body, "body", None))

    # -- pass 1: pair permStart/permEnd markers by w:id ---------------------
    open_starts: "dict[str, list[_PermStartRec]]" = {}
    range_by_start_key: "dict[int, PermissionRange]" = {}
    range_by_end_key: "dict[int, PermissionRange]" = {}
    ranges: "list[PermissionRange]" = []
    for event in events:
        if event.kind == "perm_start":
            rec = _validate_perm_start(event.element, event.path, event.tbl_ctx, report.issues)
            if rec.perm_id is None:
                report.issues.append(
                    ProtectionIssue(
                        IssueKind.UNPAIRED_PERM_START,
                        None,
                        event.path,
                        "permStart has no w:id and cannot be paired",
                    )
                )
                continue
            open_starts.setdefault(rec.perm_id, []).append(rec)
            range_by_start_key[rec.element_key] = rec  # placeholder until paired
        elif event.kind == "perm_end":
            perm_id = event.element.get(qn("w:id"))
            stack = open_starts.get(perm_id) if perm_id is not None else None
            if stack:
                rec = stack.pop()
                pr = PermissionRange(
                    perm_id=rec.perm_id,
                    ed=rec.ed,
                    ed_grp=rec.ed_grp,
                    col_first=rec.col_first,
                    col_last=rec.col_last,
                    start=MarkerRef(rec.perm_id, rec.location),
                    end=MarkerRef(perm_id, event.path),
                    valid=rec.valid,
                )
                ranges.append(pr)
                range_by_start_key[rec.element_key] = pr
                range_by_end_key[id(event.element)] = pr
            else:
                report.issues.append(
                    ProtectionIssue(
                        IssueKind.UNPAIRED_PERM_END,
                        perm_id,
                        event.path,
                        "permEnd has no matching permStart",
                    )
                )
    for perm_id, stack in open_starts.items():
        for rec in stack:
            range_by_start_key.pop(rec.element_key, None)
            report.issues.append(
                ProtectionIssue(
                    IssueKind.UNPAIRED_PERM_START,
                    perm_id,
                    rec.location,
                    "permStart has no matching permEnd; it grants no one",
                )
            )
    report.permission_ranges = ranges

    # -- pass 2: partition text into segments by active permission ranges ---
    active: "list[tuple[PermissionRange, tuple[int, int] | None]]" = []
    segments: "list[Segment]" = []
    seg_path = ""
    seg_start = 0
    para_offset = 0
    seg_parts: "list[str]" = []
    seg_ranges: "tuple[PermissionRange, ...]" = ()

    def flush() -> None:
        nonlocal seg_parts
        if not seg_parts:
            return
        text = "".join(seg_parts)
        seg_parts = []
        if not text:
            return
        active_ids = tuple(pr.perm_id or "" for pr in seg_ranges)
        if not protection.read_only_enforced:
            status = SegmentStatus.NOT_PROTECTED
            granting: "tuple[str, ...]" = ()
        else:
            granting = tuple(
                pr.perm_id or "" for pr in seg_ranges if pr.grants(user_name, user_groups)
            )
            status = SegmentStatus.EDITABLE_EXCEPTION if granting else SegmentStatus.RESTRICTED
        segments.append(
            Segment(seg_path, seg_start, seg_start + len(text), text, status, active_ids, granting)
        )

    for event in events:
        if event.kind == "perm_start":
            pr = range_by_start_key.get(id(event.element))
            if isinstance(pr, PermissionRange):
                row_key = (
                    event.tbl_ctx.row_key
                    if (pr.col_first is not None and event.tbl_ctx is not None)
                    else None
                )
                active.append((pr, row_key))
        elif event.kind == "perm_end":
            pr = range_by_end_key.get(id(event.element))
            if pr is not None:
                for idx in range(len(active) - 1, -1, -1):
                    if active[idx][0] is pr:
                        del active[idx]
                        break
        elif event.kind == "row_end":
            row_key = event.tbl_ctx.row_key if event.tbl_ctx is not None else None
            active = [(pr, key) for pr, key in active if key is None or key != row_key]
        else:  # text
            if event.path != seg_path:
                flush()
                seg_path = event.path
                para_offset = 0
                seg_start = 0
            covering = tuple(
                pr
                for pr, _key in active
                if pr.col_first is None
                or (
                    event.tbl_ctx is not None
                    and pr.col_first <= event.tbl_ctx.col_idx <= (pr.col_last or 0)
                )
            )
            if seg_parts and covering != seg_ranges:
                flush()
                seg_start = para_offset
            if not seg_parts:
                seg_ranges = covering
            seg_parts.append(event.text)
            para_offset += len(event.text)
    flush()
    report.segments = segments
    return report
