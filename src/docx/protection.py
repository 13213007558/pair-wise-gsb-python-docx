"""Read-only analysis of document-protection editability exceptions.

This module interprets the protection rules recorded in a .docx file: the
'w:documentProtection' element in the settings part and the
'w:permStart'/'w:permEnd' range-permission markers in the document body.
Given a user and/or group identity, it reports which content in the body
and tables is restricted and which falls inside an editable exception range.

The analysis is strictly read-only: it neither verifies the protection
password nor modifies (unprotects) the document in any way.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Callable, Iterable

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.oxml.xmlchemy import BaseOxmlElement

#: 'w:edit' values for which range-permission exceptions apply.
_EXCEPTION_EDIT_MODES = ("readOnly", "forms")

#: 'w:ed' value reserved by Word to grant a range to every user.
EVERYONE = "everyone"

#: Segment reasons.
REASON_PROTECTION_NOT_ENFORCED = "protection-not-enforced"
REASON_EXCEPTION = "exception"
REASON_NO_EXCEPTION = "no-exception"

#: Issue kinds.
ISSUE_UNKNOWN_GROUP = "unknown-group"
ISSUE_MISSING_START = "missing-start"
ISSUE_MISSING_END = "missing-end"
ISSUE_INVALID_COLUMN_RANGE = "invalid-column-range"
ISSUE_MISSING_PRINCIPAL = "missing-principal"
ISSUE_INVALID_MARKER = "invalid-marker"


@dataclass(frozen=True)
class DocumentProtectionInfo:
    """Parsed 'w:documentProtection' settings element.

    This is a factual record of what the file declares; the password hash is
    never inspected or verified.
    """

    present: bool
    edit: str | None
    enforcement: bool

    @property
    def enforced(self) -> bool:
        """True when range-permission exceptions are in effect.

        That is the case only when protection is present, enforcement is
        switched on, and the edit mode is one ('readOnly' or 'forms') to
        which 'w:permStart'/'w:permEnd' exceptions apply.
        """
        return self.present and self.enforcement and self.edit in _EXCEPTION_EDIT_MODES


@dataclass(frozen=True)
class MarkerInfo:
    """A single 'w:permStart' or 'w:permEnd' marker, for traceability."""

    perm_id: int | None
    kind: str  # "permStart" | "permEnd"
    location: str
    ed: str | None
    ed_grp: str | None
    col_first: int | None
    col_last: int | None


@dataclass(frozen=True)
class ProtectionIssue:
    """A protection-recording problem found during analysis.

    Issues are reported, never silently resolved in the user's favor.
    """

    kind: str
    message: str
    marker: MarkerInfo


@dataclass
class Segment:
    """A contiguous run of body or table content sharing one editability verdict."""

    editable: bool
    reason: str  # REASON_* constant
    location: str
    text: str
    perm_id: int | None = None
    granted_by: str | None = None  # "ed" | "edGrp" | "everyone"
    start_marker: MarkerInfo | None = None
    end_marker: MarkerInfo | None = None


@dataclass
class EditabilityReport:
    """Result of :func:'analyze_editability'."""

    protection: DocumentProtectionInfo
    segments: list[Segment] = field(default_factory=list)
    issues: list[ProtectionIssue] = field(default_factory=list)

    @property
    def enforced(self) -> bool:
        """True when document protection is enforced for range permissions."""
        return self.protection.enforced

    @property
    def status(self) -> str:
        """'enforced' when exceptions apply, otherwise 'not-enforced'."""
        return "enforced" if self.enforced else "not-enforced"

    @property
    def editable_segments(self) -> list[Segment]:
        return [segment for segment in self.segments if segment.editable]

    @property
    def restricted_segments(self) -> list[Segment]:
        return [segment for segment in self.segments if not segment.editable]


def protection_info(settings_el: BaseOxmlElement | None) -> DocumentProtectionInfo:
    """Parse 'w:documentProtection' from a 'w:settings' element."""
    doc_protect = settings_el.find(qn("w:documentProtection")) if settings_el is not None else None
    if doc_protect is None:
        return DocumentProtectionInfo(present=False, edit=None, enforcement=False)
    edit = doc_protect.get(qn("w:edit"))
    enforcement = doc_protect.get(qn("w:enforcement")) in ("on", "1", "true")
    return DocumentProtectionInfo(present=True, edit=edit, enforcement=enforcement)


def analyze_editability(
    document: Document, user: str | None = None, groups: Iterable[str] = ()
) -> EditabilityReport:
    """Report which body and table content is editable for 'user'/'groups'.

    'document' is a :class:'docx.document.Document' instance. 'user' is
    matched against 'w:ed' and each of 'groups' against 'w:edGrp' on
    'w:permStart' markers. The document is only read, never modified.
    """
    protection = protection_info(document.settings.element)
    analyzer = _Analyzer(protection, user, frozenset(groups))
    return analyzer.analyze(document.element.body)


class _Range:
    """An open 'w:permStart' range, closed by the matching 'w:permEnd'."""

    def __init__(
        self,
        start_marker: MarkerInfo,
        granted: bool,
        granted_by: str | None,
        col_first: int | None,
        col_last: int | None,
    ):
        self.start_marker = start_marker
        self.granted = granted
        self.granted_by = granted_by
        self.col_first = col_first
        self.col_last = col_last
        self.end_marker: MarkerInfo | None = None

    def applies_to_column(self, col_idx: int | None) -> bool:
        """True when this range covers table-column 'col_idx'.

        A range without column attributes covers everything. A column-scoped
        range only covers cells whose column index lies within
        'col_first..col_last'; it grants nothing outside a table cell.
        """
        if self.col_first is None or self.col_last is None:
            return True
        if col_idx is None:
            return False
        return self.col_first <= col_idx <= self.col_last


class _Analyzer:
    """Single-pass document-order walker producing an |EditabilityReport|."""

    def __init__(
        self, protection: DocumentProtectionInfo, user: str | None, groups: frozenset[str]
    ):
        self._protection = protection
        self._user = user
        self._groups = groups
        self._active: dict[int, _Range] = {}
        self._col_stack: list[int] = []
        self._segments: list[Segment] = []
        self._issues: list[ProtectionIssue] = []
        self._segment_ranges: list[tuple[Segment, _Range]] = []

    def analyze(self, body_el: BaseOxmlElement) -> EditabilityReport:
        self._walk_blocks(body_el, "body")
        for range_ in self._active.values():
            self._issues.append(
                ProtectionIssue(
                    ISSUE_MISSING_END,
                    "permStart id=%d at %s has no matching permEnd"
                    % (range_.start_marker.perm_id, range_.start_marker.location),
                    range_.start_marker,
                )
            )
        for segment, range_ in self._segment_ranges:
            segment.start_marker = range_.start_marker
            segment.end_marker = range_.end_marker
        return EditabilityReport(self._protection, self._segments, self._issues)

    # -- element walking ---------------------------------------------------

    def _walk_blocks(self, parent_el: BaseOxmlElement, loc: str) -> None:
        p_idx, tbl_idx = 0, 0
        for child in parent_el:
            tag = child.tag
            if tag == qn("w:p"):
                self._paragraph(child, "%s/p[%d]" % (loc, p_idx))
                p_idx += 1
            elif tag == qn("w:tbl"):
                self._table(child, "%s/tbl[%d]" % (loc, tbl_idx))
                tbl_idx += 1
            elif tag == qn("w:permStart"):
                self._perm_start(child, loc)
            elif tag == qn("w:permEnd"):
                self._perm_end(child, loc)
            elif tag == qn("w:sectPr"):
                continue
            else:
                self._walk_blocks(child, loc)

    def _table(self, tbl_el: BaseOxmlElement, loc: str) -> None:
        row_idx = 0
        for child in tbl_el:
            if child.tag == qn("w:tr"):
                self._row(child, "%s/tr[%d]" % (loc, row_idx))
                row_idx += 1

    def _row(self, tr_el: BaseOxmlElement, loc: str) -> None:
        col_idx = 0
        for child in tr_el:
            if child.tag == qn("w:permStart"):
                self._perm_start(child, loc)
            elif child.tag == qn("w:permEnd"):
                self._perm_end(child, loc)
            elif child.tag == qn("w:tc"):
                self._col_stack.append(col_idx)
                self._walk_blocks(child, "%s/tc[%d]" % (loc, col_idx))
                self._col_stack.pop()
                col_idx += 1

    def _paragraph(self, p_el: BaseOxmlElement, loc: str) -> None:
        fragments: list[str] = []
        emitted = False

        def flush() -> None:
            nonlocal emitted
            text = "".join(fragments)
            if text or not emitted:
                self._emit(loc, text)
                emitted = True
            fragments.clear()

        self._walk_inline(p_el, loc, fragments, flush)
        flush()

    def _walk_inline(
        self,
        parent_el: BaseOxmlElement,
        loc: str,
        fragments: list[str],
        flush: Callable[[], None],
    ) -> None:
        for child in parent_el:
            tag = child.tag
            if tag == qn("w:permStart"):
                flush()
                self._perm_start(child, loc)
            elif tag == qn("w:permEnd"):
                flush()
                self._perm_end(child, loc)
            elif tag == qn("w:r"):
                fragments.extend(t.text or "" for t in child.iter(qn("w:t")))
            else:
                self._walk_inline(child, loc, fragments, flush)

    # -- permission markers --------------------------------------------------

    def _perm_start(self, el: BaseOxmlElement, loc: str) -> None:
        marker = self._marker_info(el, "permStart", loc)
        if marker.perm_id is None:
            self._issues.append(
                ProtectionIssue(
                    ISSUE_INVALID_MARKER, "permStart at %s has no usable w:id" % loc, marker
                )
            )
            return
        if marker.perm_id in self._active:
            self._issues.append(
                ProtectionIssue(
                    ISSUE_INVALID_MARKER,
                    "duplicate permStart id=%d at %s" % (marker.perm_id, loc),
                    marker,
                )
            )
            return
        granted, granted_by = self._match_identity(marker)
        col_first, col_last = marker.col_first, marker.col_last
        if (col_first is not None or col_last is not None) and not self._valid_column_range(
            el, marker
        ):
            # An illegal column range grants nothing; the range is still
            # tracked so its permEnd pairs cleanly.
            granted, granted_by = False, None
            col_first = col_last = None
        self._active[marker.perm_id] = _Range(marker, granted, granted_by, col_first, col_last)

    def _perm_end(self, el: BaseOxmlElement, loc: str) -> None:
        marker = self._marker_info(el, "permEnd", loc)
        if marker.perm_id is None:
            self._issues.append(
                ProtectionIssue(
                    ISSUE_INVALID_MARKER, "permEnd at %s has no usable w:id" % loc, marker
                )
            )
            return
        range_ = self._active.pop(marker.perm_id, None)
        if range_ is None:
            self._issues.append(
                ProtectionIssue(
                    ISSUE_MISSING_START,
                    "permEnd id=%d at %s has no matching permStart" % (marker.perm_id, loc),
                    marker,
                )
            )
            return
        range_.end_marker = marker

    def _match_identity(self, marker: MarkerInfo) -> tuple[bool, str | None]:
        ed, ed_grp = marker.ed, marker.ed_grp
        if ed is not None:
            if ed == EVERYONE:
                return True, "everyone"
            if self._user is not None and ed == self._user:
                return True, "ed"
        if ed_grp is not None:
            if ed_grp in self._groups:
                return True, "edGrp"
            self._issues.append(
                ProtectionIssue(
                    ISSUE_UNKNOWN_GROUP,
                    "permStart id=%d at %s references unknown group %r; membership "
                    "cannot be verified, so the range is not granted"
                    % (marker.perm_id, marker.location, ed_grp),
                    marker,
                )
            )
        if ed is None and ed_grp is None:
            self._issues.append(
                ProtectionIssue(
                    ISSUE_MISSING_PRINCIPAL,
                    "permStart id=%d at %s names neither w:ed nor w:edGrp"
                    % (marker.perm_id, marker.location),
                    marker,
                )
            )
        return False, None

    def _valid_column_range(self, el: BaseOxmlElement, marker: MarkerInfo) -> bool:
        parent = el.getparent()
        problem = None
        if marker.col_first is None or marker.col_last is None:
            problem = "only one of w:colFirst/w:colLast is present"
        elif parent is None or parent.tag != qn("w:tr"):
            problem = "marker is not a direct child of a table row"
        elif marker.col_first < 0 or marker.col_first > marker.col_last:
            problem = "w:colFirst=%d is not <= w:colLast=%d" % (
                marker.col_first,
                marker.col_last,
            )
        else:
            cell_count = sum(1 for c in parent if c.tag == qn("w:tc"))
            if marker.col_last >= cell_count:
                problem = "w:colLast=%d exceeds row cell count %d" % (
                    marker.col_last,
                    cell_count,
                )
        if problem is None:
            return True
        self._issues.append(
            ProtectionIssue(
                ISSUE_INVALID_COLUMN_RANGE,
                "permStart id=%d at %s has an invalid column range (%s); "
                "the range is not granted" % (marker.perm_id, marker.location, problem),
                marker,
            )
        )
        return False

    # -- segment emission ----------------------------------------------------

    def _emit(self, loc: str, text: str) -> None:
        if not self._protection.enforced:
            self._segments.append(Segment(True, REASON_PROTECTION_NOT_ENFORCED, loc, text))
            return
        col_idx = self._col_stack[-1] if self._col_stack else None
        for range_ in self._active.values():
            if range_.granted and range_.applies_to_column(col_idx):
                segment = Segment(
                    True,
                    REASON_EXCEPTION,
                    loc,
                    text,
                    perm_id=range_.start_marker.perm_id,
                    granted_by=range_.granted_by,
                )
                self._segments.append(segment)
                self._segment_ranges.append((segment, range_))
                return
        self._segments.append(Segment(False, REASON_NO_EXCEPTION, loc, text))

    # -- helpers -------------------------------------------------------------

    @staticmethod
    def _marker_info(el: BaseOxmlElement, kind: str, loc: str) -> MarkerInfo:
        def int_attr(name: str) -> int | None:
            try:
                return int(el.get(qn(name)))
            except (TypeError, ValueError):
                return None

        return MarkerInfo(
            perm_id=int_attr("w:id"),
            kind=kind,
            location=loc,
            ed=el.get(qn("w:ed")),
            ed_grp=el.get(qn("w:edGrp")),
            col_first=int_attr("w:colFirst"),
            col_last=int_attr("w:colLast"),
        )
