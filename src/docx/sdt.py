"""Content controls (w:sdt) bound to custom XML data-stores.

A content control can be bound to an item in a custom XML data-store with a
w:dataBinding element naming a storeItemID, an XPath, and prefix mappings. The
objects in this module locate those controls, update the underlying data item, and
refresh the displayed content of every control bound to the same item, so the
document view and the stored data stay consistent.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import TYPE_CHECKING, Any, Dict, Iterator, List, Optional, Tuple

from lxml import etree

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.document import Document
    from docx.opc.part import Part


# --- content-control types supported for bound updates ---
TEXT = "text"
DATE = "date"
CHECKBOX = "checkbox"
UNKNOWN = "unknown"

# --- binding-conflict reasons ---
STORE_NOT_FOUND = "store-not-found"
XPATH_NO_MATCH = "xpath-no-match"
XPATH_MULTIPLE_MATCHES = "xpath-multiple-matches"
XPATH_INVALID = "xpath-invalid"
UNSUPPORTED_CONTROL = "unsupported-control"
INCOMPATIBLE_CONTROL = "incompatible-control"


class BindingConflict:
    """A single, locatable reason a bound-control update cannot proceed."""

    def __init__(
        self,
        reason: str,
        message: str,
        part_name: Optional[str] = None,
        control: Optional[str] = None,
    ):
        self.reason = reason
        self.message = message
        self.part_name = part_name
        self.control = control

    @property
    def location(self) -> Optional[str]:
        """Best available locator for this conflict, e.g. '/word/header1.xml#sdt[123]'."""
        return self.control or self.part_name

    def __str__(self) -> str:
        if self.location is None:
            return "%s: %s" % (self.reason, self.message)
        return "%s at %s: %s" % (self.reason, self.location, self.message)


class BindingConflictError(ValueError):
    """Raised when a bound-control update cannot be applied consistently.

    The update is atomic: when this exception is raised, neither the custom XML
    data-store nor any control display has been modified. The .conflicts member
    holds a |BindingConflict| for each distinct problem found.
    """

    def __init__(self, conflicts: List[BindingConflict]):
        self.conflicts = list(conflicts)
        super().__init__("; ".join(str(c) for c in self.conflicts))


class BoundUpdateResult:
    """Summary of a successfully applied bound-control update."""

    def __init__(
        self,
        store_item_id: str,
        xpath: str,
        stored_value: str,
        controls: List[BoundContentControl],
    ):
        self.store_item_id = store_item_id
        self.xpath = xpath
        self.stored_value = stored_value
        self.updated_controls = list(controls)

    def __iter__(self) -> Iterator[BoundContentControl]:
        return iter(self.updated_controls)

    def __len__(self) -> int:
        return len(self.updated_controls)

    @property
    def locations(self) -> List[str]:
        """Locators of the controls whose display was refreshed."""
        return [c.location for c in self.updated_controls]


class BoundContentControl(ElementProxy):
    """A w:sdt content control bound to a custom XML data item."""

    def __init__(self, sdt: Any, part: Part):
        super().__init__(sdt, part)
        self._sdt = sdt
        self._part = part

    @property
    def control_id(self) -> Optional[str]:
        """The w:id value of this control, used as a stable per-part identifier."""
        sdtPr = self._sdtPr
        if sdtPr is None:
            return None
        id_elm = sdtPr.find(qn("w:id"))
        return None if id_elm is None else id_elm.get(qn("w:val"))

    @property
    def control_type(self) -> str:
        """One of 'text', 'date', 'checkbox', or 'unknown' for unsupported kinds."""
        sdtPr = self._sdtPr
        if sdtPr is None:
            return UNKNOWN
        if sdtPr.find(qn("w14:checkbox")) is not None:
            return CHECKBOX
        if sdtPr.find(qn("w:date")) is not None:
            return DATE
        if sdtPr.find(qn("w:text")) is not None:
            return TEXT
        return UNKNOWN

    @property
    def displayed_text(self) -> str:
        """The text currently displayed by this control."""
        content = self._sdt.find(qn("w:sdtContent"))
        if content is None:
            return ""
        return "".join(t.text or "" for t in content.iter(qn("w:t")))

    @property
    def location(self) -> str:
        """Locator string identifying this control, e.g. '/word/document.xml#sdt[42]'."""
        return "%s#sdt[%s]" % (self.part_name, self.control_id or "?")

    @property
    def part_name(self) -> str:
        """Partname of the story part containing this control."""
        return str(self._part.partname)

    @property
    def prefix_mappings(self) -> Dict[str, str]:
        """Prefix-to-namespace-URI map parsed from w:prefixMappings."""
        return parse_prefix_mappings(self._binding_attr("prefixMappings") or "")

    @property
    def store_item_id(self) -> Optional[str]:
        """The storeItemID GUID this control is bound to."""
        return self._binding_attr("storeItemID")

    @property
    def xpath(self) -> Optional[str]:
        """The XPath this control is bound with."""
        return self._binding_attr("xpath")

    def refresh(self, value: Any) -> None:
        """Update the displayed content of this control to show `value`.

        Only the runs inside w:sdtContent are touched; the control properties,
        including styles, locking, and any unknown extension nodes, are preserved.
        """
        control_type = self.control_type
        if control_type == CHECKBOX:
            self._refresh_checkbox(_as_bool(value))
        elif control_type in (TEXT, DATE):
            self._replace_content_text(self._display_text_for(control_type, value))
        else:
            raise ValueError("cannot refresh a control of unknown type")

    @property
    def _sdtPr(self):
        return self._sdt.find(qn("w:sdtPr"))

    def _binding_attr(self, name: str) -> Optional[str]:
        sdtPr = self._sdtPr
        if sdtPr is None:
            return None
        binding = sdtPr.find(qn("w:dataBinding"))
        return None if binding is None else binding.get(qn("w:%s" % name))

    def _date_format(self) -> str:
        """The w:dateFormat value of a date control, defaulting to ISO format."""
        sdtPr = self._sdtPr
        if sdtPr is not None:
            date_elm = sdtPr.find(qn("w:date"))
            if date_elm is not None:
                fmt_elm = date_elm.find(qn("w:dateFormat"))
                if fmt_elm is not None:
                    fmt = fmt_elm.get(qn("w:val"))
                    if fmt:
                        return fmt
        return "yyyy-MM-dd"

    def _display_text_for(self, control_type: str, value: Any) -> str:
        if control_type == DATE:
            moment, _ = _as_datetime(value)
            return moment.strftime(word_date_format_to_strftime(self._date_format()))
        return str(value)

    def _refresh_checkbox(self, checked: bool) -> None:
        sdtPr = self._sdtPr
        checkbox = sdtPr.find(qn("w14:checkbox")) if sdtPr is not None else None
        if checkbox is None:
            return
        checked_elm = checkbox.find(qn("w14:checked"))
        if checked_elm is None:
            checked_elm = OxmlElement("w14:checked")
            checkbox.insert(0, checked_elm)
        checked_elm.set(qn("w14:val"), "1" if checked else "0")
        state = checkbox.find(qn("w14:checkedState" if checked else "w14:uncheckedState"))
        glyph = "\u2612" if checked else "\u2610"
        if state is not None:
            state_val = state.get(qn("w14:val"))
            if state_val:
                glyph = chr(int(state_val, 16))
        self._replace_content_text(glyph)

    def _replace_content_text(self, text: str) -> None:
        content = self._sdt.find(qn("w:sdtContent"))
        if content is None:
            return
        runs = list(content.iter(qn("w:r")))
        if runs:
            run = runs[0]
            for extra in runs[1:]:
                extra.getparent().remove(extra)
        else:
            run = OxmlElement("w:r")
            paragraph = content.find(qn("w:p"))
            (paragraph if paragraph is not None else content).append(run)
        for child in list(run):
            if child.tag != qn("w:rPr"):
                run.remove(child)
        t = OxmlElement("w:t")
        t.text = text
        t.set(qn("xml:space"), "preserve")
        run.append(t)


_PREFIX_MAPPING_RE = re.compile(r"xmlns:([A-Za-z_][\w.-]*)\s*=\s*[\x22\x27]([^\x22\x27]*)[\x22\x27]")


def parse_prefix_mappings(prefix_mappings: str) -> Dict[str, str]:
    """Return prefix-to-URI dict parsed from a w:prefixMappings attribute value.

    The attribute looks like: xmlns:ns0='http://example.com/a' xmlns:ns1="http://b".
    """
    return {
        match.group(1): match.group(2)
        for match in _PREFIX_MAPPING_RE.finditer(prefix_mappings or "")
    }


# --- Word date-format tokens, longest first so greedy matching works ---
_WORD_DATE_TOKENS = [
    ("yyyy", "%Y"),
    ("yyy", "%Y"),
    ("yy", "%y"),
    ("MMMM", "%B"),
    ("MMM", "%b"),
    ("MM", "%m"),
    ("M", "%m"),
    ("dd", "%d"),
    ("d", "%d"),
    ("HH", "%H"),
    ("H", "%H"),
    ("hh", "%I"),
    ("h", "%I"),
    ("mm", "%M"),
    ("m", "%M"),
    ("ss", "%S"),
    ("s", "%S"),
    ("tt", "%p"),
]

_DATE_TOKEN_RE = re.compile("|".join(token for token, _ in _WORD_DATE_TOKENS))
_STRFTIME_FOR_TOKEN = dict(_WORD_DATE_TOKENS)


def word_date_format_to_strftime(date_format: str) -> str:
    """Return strftime format string approximating a Word w:dateFormat value.

    Word tokens like yyyy, MM, and dd are translated; other characters, including
    quoted literal text, pass through literally.
    """
    out: List[str] = []
    pos = 0
    for match in _DATE_TOKEN_RE.finditer(date_format):
        out.append(date_format[pos : match.start()].replace("%", "%%"))
        out.append(_STRFTIME_FOR_TOKEN[match.group(0)])
        pos = match.end()
    out.append(date_format[pos:].replace("%", "%%"))
    return "".join(out).replace('"', "")


_TRUE_VALUES = {"true", "1"}
_FALSE_VALUES = {"false", "0"}


def _as_bool(value: Any) -> bool:
    """Return `value` as a bool, raising ValueError if it is not boolean-like."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in _TRUE_VALUES:
            return True
        if normalized in _FALSE_VALUES:
            return False
    raise ValueError("not a boolean value: %r" % (value,))


def _as_datetime(value: Any) -> Tuple[dt.datetime, bool]:
    """Return (datetime, had-time) pair for `value`, raising ValueError if invalid."""
    if isinstance(value, dt.datetime):
        return value, True
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day), False
    if isinstance(value, str):
        text = value.strip()
        try:
            moment = dt.datetime.fromisoformat(text)
        except ValueError:
            raise ValueError("not an ISO date/time value: %r" % (value,))
        return moment, ("T" in text or " " in text)
    raise ValueError("not a date/time value: %r" % (value,))


def _accepts(control_type: str, value: Any) -> bool:
    """True if a control of `control_type` can display `value`."""
    if control_type == TEXT:
        return isinstance(value, str)
    if control_type == DATE:
        try:
            _as_datetime(value)
            return True
        except ValueError:
            return False
    if control_type == CHECKBOX:
        try:
            _as_bool(value)
            return True
        except ValueError:
            return False
    return False


def _canonical_stored_value(control_type: str, value: Any) -> str:
    """Return the canonical string form of `value` for storage in custom XML."""
    if control_type == CHECKBOX:
        return "true" if _as_bool(value) else "false"
    if control_type == DATE:
        moment, had_time = _as_datetime(value)
        if had_time:
            return moment.isoformat(timespec="seconds")
        return moment.date().isoformat()
    return str(value)


def _infer_stored_value(value: Any) -> str:
    """Return canonical stored form of `value` when no control constrains it."""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (dt.date, dt.datetime)):
        return _canonical_stored_value(DATE, value)
    return str(value)


def _set_target_value(target: Any, stored_value: str) -> None:
    """Write `stored_value` to the custom XML node `target`.

    `target` is an element when the bound XPath selects an element, or an
    attribute smart-string when it selects an attribute.
    """
    if isinstance(target, etree._Element):
        target.text = stored_value
    else:
        parent = target.getparent()
        if parent is None:
            raise ValueError("cannot update an attribute of a detached node")
        parent.set(target.attrname, stored_value)


def update_bound_controls(
    document: Document,
    store_item_id: str,
    xpath: str,
    value: Any,
    namespaces: Optional[Dict[str, str]] = None,
) -> BoundUpdateResult:
    """Update a custom XML data item and refresh every control bound to it.

    Raises |BindingConflictError| before modifying anything when the update
    cannot be applied consistently.
    """
    store = document.custom_xml_stores.get_by_item_id(store_item_id)
    if store is None:
        raise BindingConflictError(
            [
                BindingConflict(
                    STORE_NOT_FOUND,
                    "no custom XML store with storeItemID %r" % store_item_id,
                )
            ]
        )

    controls = document.find_bound_controls(store_item_id, xpath)

    if namespaces is None:
        namespaces = {}
        for control in controls:
            for prefix, uri in control.prefix_mappings.items():
                namespaces.setdefault(prefix, uri)

    conflicts: List[BindingConflict] = []

    try:
        targets = store.find(xpath, namespaces)
    except etree.XPathError as exc:
        conflicts.append(
            BindingConflict(
                XPATH_INVALID,
                "xpath %r cannot be evaluated: %s" % (xpath, exc),
            )
        )
        raise BindingConflictError(conflicts)

    if len(targets) == 0:
        conflicts.append(
            BindingConflict(
                XPATH_NO_MATCH,
                "xpath %r matches no node in store %s" % (xpath, store.store_item_id),
            )
        )
    elif len(targets) > 1:
        conflicts.append(
            BindingConflict(
                XPATH_MULTIPLE_MATCHES,
                "xpath %r matches %d nodes in store %s"
                % (xpath, len(targets), store.store_item_id),
            )
        )

    for control in controls:
        control_type = control.control_type
        if control_type == UNKNOWN:
            conflicts.append(
                BindingConflict(
                    UNSUPPORTED_CONTROL,
                    "control type is not supported for bound updates",
                    part_name=control.part_name,
                    control=control.location,
                )
            )
        elif not _accepts(control_type, value):
            conflicts.append(
                BindingConflict(
                    INCOMPATIBLE_CONTROL,
                    "%s control cannot display value %r" % (control_type, value),
                    part_name=control.part_name,
                    control=control.location,
                )
            )

    if conflicts:
        raise BindingConflictError(conflicts)

    stored_value = (
        _canonical_stored_value(controls[0].control_type, value)
        if controls
        else _infer_stored_value(value)
    )
    _set_target_value(targets[0], stored_value)
    for control in controls:
        control.refresh(value)

    return BoundUpdateResult(store.store_item_id or store_item_id, xpath, stored_value, controls)

