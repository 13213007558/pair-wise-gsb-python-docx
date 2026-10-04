"""Content-control (SDT) data-binding refresh.

Word content controls (w:sdt elements) can be bound to nodes in a custom XML part via
a w:dataBinding element. Word refreshes the displayed content of such controls from
the stored XML data when a document is opened; this module provides the same refresh
explicitly, so a caller can update the custom XML data and then synchronize the
displayed control content without reopening the document in Word.

Refresh is a two-step affair. DocumentBindings.plan_refresh() inspects all bound
controls and the custom XML data and produces a |RefreshPlan| describing the proposed
changes along with any problems (unsupported XPath, unexpected match count, value
type mismatch). The plan can be inspected before RefreshPlan.apply() commits any
change; apply() refuses to modify the document while problems are present.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, Any, List, NamedTuple, Tuple

from docx.customxml import (
    custom_xml_stores,
    element_text,
    evaluate_xpath_subset,
    parse_prefix_mappings,
)
from docx.exceptions import BindingRefreshError, UnsupportedXPathError
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.customxml import CustomXmlStore
    from docx.oxml.document import CT_Document
    from docx.parts.document import DocumentPart

# -- control kinds handled by refresh --
PLAIN_TEXT = "plain-text"
CHECKBOX = "checkbox"
REPEATING_SECTION = "repeating-section"

_TRUE_VALUES = frozenset({"true", "1", "yes", "on"})
_FALSE_VALUES = frozenset({"false", "0", "no", "off"})


class BindingProblem(NamedTuple):
    """Describes why a single bound content control cannot be refreshed.

    'item_index' is the 0-based position within the matched nodes of a repeating
    section when the problem belongs to a control inside a repeating section item,
    otherwise |None|.
    """

    control_id: "int | None"
    tag: str
    kind: "str | None"
    xpath: str
    reason: str
    item_index: "int | None" = None

    def __str__(self) -> str:
        location = f"control id={self.control_id} tag={self.tag!r} kind={self.kind}"
        if self.item_index is not None:
            location += f" item={self.item_index}"
        return f"{location} xpath={self.xpath!r}: {self.reason}"


class PlanEntry(NamedTuple):
    """A single proposed refresh action, inspectable before the plan is applied."""

    control_id: "int | None"
    tag: str
    kind: str
    store_item_id: str
    xpath: str
    action: str
    value: "str | None" = None
    item_count: "int | None" = None


class RefreshPlan:
    """The proposed result of refreshing all bound content controls in a document.

    Produced by DocumentBindings.plan_refresh(). Reading the plan has no effect on
    the document; call apply() to commit the proposed changes. When any control
    cannot be refreshed, '.problems' identifies each offending control and apply()
    raises |BindingRefreshError| before anything is modified.
    """

    def __init__(
        self,
        document_element: "CT_Document",
        entries: List[Tuple[PlanEntry, Any]],
        problems: List[BindingProblem],
    ):
        self._document_element = document_element
        self._entries = entries
        self._problems = problems

    @property
    def entries(self) -> Tuple[PlanEntry, ...]:
        """The proposed refresh actions, in document order."""
        return tuple(entry for entry, _ in self._entries)

    @property
    def problems(self) -> Tuple[BindingProblem, ...]:
        """Problems found while planning, each identifying its content control."""
        return tuple(self._problems)

    @property
    def ok(self) -> bool:
        """True when the plan has no problems and can be applied."""
        return not self._problems

    def apply(self) -> None:
        """Commit all proposed changes to the document.

        Raises |BindingRefreshError| listing every problem when any bound control
        could not be refreshed; the document is left unmodified in that case.
        """
        if self._problems:
            raise BindingRefreshError(self._problems)
        allocator = _SdtIdAllocator(self._document_element)
        for entry, payload in self._entries:
            _apply_entry(entry, payload, allocator)


class DocumentBindings:
    """Interface for refreshing content controls from custom XML data.

    Accessed as Document.bindings. Typical usage is to update the custom XML data
    through a |CustomXmlStore|, then call refresh() (or plan_refresh() followed by
    RefreshPlan.apply()) to synchronize the displayed content of all bound controls.
    """

    def __init__(self, document_part: "DocumentPart", document_element: "CT_Document"):
        self._part = document_part
        self._document_element = document_element

    @property
    def stores(self) -> "Tuple[CustomXmlStore, ...]":
        """All custom XML data stores in the package, in partname order."""
        return tuple(custom_xml_stores(self._part.package))

    def store(self, item_id: str) -> "CustomXmlStore":
        """The |CustomXmlStore| whose ds:itemID matches 'item_id'.

        Matching is case-insensitive on the GUID. Raises |KeyError| when no store in
        the package declares that item ID.
        """
        for store in self.stores:
            if store.item_id is not None and store.item_id.lower() == item_id.lower():
                return store
        raise KeyError(f"no custom XML store with item ID {item_id!r}")

    def plan_refresh(self) -> RefreshPlan:
        """Inspect bound controls and custom XML data and return a |RefreshPlan|.

        This performs all data reading and validation without modifying the
        document, so the proposed changes and any problems can be reviewed before
        RefreshPlan.apply() commits them.
        """
        stores = list(self.stores)
        entries: List[Tuple[PlanEntry, Any]] = []
        problems: List[BindingProblem] = []
        for sdt in _iter_top_level_sdts(self._document_element):
            control = _Control(sdt)
            if not control.bound:
                continue
            entry, payload = _plan_control(control, stores, problems)
            if entry is not None:
                entries.append((entry, payload))
        return RefreshPlan(self._document_element, entries, problems)

    def refresh(self) -> RefreshPlan:
        """Plan and immediately apply a refresh of all bound content controls.

        Returns the applied |RefreshPlan|. Raises |BindingRefreshError| when any
        control could not be refreshed, leaving the document unmodified.
        """
        plan = self.plan_refresh()
        plan.apply()
        return plan


# ============================================================================
# planning internals
# ============================================================================


class _Control:
    """Parsed view of a single w:sdt element and its w:sdtPr properties."""

    def __init__(self, sdt: "_Element"):
        self.sdt = sdt
        sdtPr = sdt.find(qn("w:sdtPr"))
        self.control_id = _sdt_id(sdtPr)
        self.tag = _str_attr(sdtPr, "w:tag")
        self.alias = _str_attr(sdtPr, "w:alias")
        self.kind = _control_kind(sdtPr)
        binding = sdtPr.find(qn("w:dataBinding")) if sdtPr is not None else None
        self.bound = binding is not None
        self.store_item_id = binding.get(qn("w:storeItemID"), "") if self.bound else ""
        self.xpath = binding.get(qn("w:xpath"), "") if self.bound else ""
        self.prefix_mappings = binding.get(qn("w:prefixMappings"), "") if self.bound else ""

    def problem(self, reason: str, item_index: "int | None" = None) -> BindingProblem:
        """A |BindingProblem| identifying this control."""
        return BindingProblem(
            control_id=self.control_id,
            tag=self.tag,
            kind=self.kind,
            xpath=self.xpath,
            reason=reason,
            item_index=item_index,
        )


def _sdt_id(sdtPr: "_Element | None") -> "int | None":
    """The integer w:id of the control, or |None| when absent or not an integer."""
    if sdtPr is None:
        return None
    id_el = sdtPr.find(qn("w:id"))
    if id_el is None:
        return None
    try:
        return int(id_el.get(qn("w:val"), ""))
    except ValueError:
        return None


def _str_attr(sdtPr: "_Element | None", tag: str) -> str:
    """The w:val of the 'tag' child of 'sdtPr', or the empty string."""
    if sdtPr is None:
        return ""
    el = sdtPr.find(qn(tag))
    if el is None:
        return ""
    return el.get(qn("w:val"), "")


def _control_kind(sdtPr: "_Element | None") -> "str | None":
    """The refresh-supported kind of the control, or |None| when unsupported."""
    if sdtPr is None:
        return None
    if sdtPr.find(qn("w15:repeatingSection")) is not None:
        return REPEATING_SECTION
    if sdtPr.find(qn("w14:checkbox")) is not None:
        return CHECKBOX
    if sdtPr.find(qn("w:text")) is not None:
        return PLAIN_TEXT
    return None


def _is_repeating_section_sdt(element: "_Element") -> bool:
    """True if 'element' is a w:sdt whose sdtPr declares w15:repeatingSection."""
    if element.tag != qn("w:sdt"):
        return False
    sdtPr = element.find(qn("w:sdtPr"))
    return sdtPr is not None and sdtPr.find(qn("w15:repeatingSection")) is not None


def _inside_repeating_section(sdt: "_Element") -> bool:
    """True if 'sdt' is contained within a repeating section control."""
    ancestor = sdt.getparent()
    while ancestor is not None:
        if _is_repeating_section_sdt(ancestor):
            return True
        ancestor = ancestor.getparent()
    return False


def _iter_top_level_sdts(root: "_Element"):
    """All w:sdt elements in 'root' except those inside a repeating section.

    Controls inside a repeating section are planned as part of the section, not
    independently.
    """
    for sdt in root.iter(qn("w:sdt")):
        if _inside_repeating_section(sdt):
            continue
        yield sdt


def _find_store(stores: "List[CustomXmlStore]", item_id: str) -> "CustomXmlStore | None":
    """The store in 'stores' whose ds:itemID matches 'item_id', if any."""
    for store in stores:
        if store.item_id is not None and store.item_id.lower() == item_id.lower():
            return store
    return None


def _plan_control(
    control: _Control,
    stores: "List[CustomXmlStore]",
    problems: List[BindingProblem],
) -> "Tuple[PlanEntry | None, Any]":
    """Plan the refresh of a single bound control.

    Returns an (entry, payload) pair on success. On failure, appends to 'problems'
    and returns (None, None); the entry is still produced for repeating sections so
    the plan describes what was attempted.
    """
    if control.kind is None:
        problems.append(control.problem("control type is not supported for refresh"))
        return None, None
    store = _find_store(stores, control.store_item_id)
    if store is None:
        problems.append(
            control.problem(f"custom XML store {control.store_item_id!r} not found in package")
        )
        return None, None
    if control.kind == REPEATING_SECTION:
        return _plan_section(control, store, problems)
    return _plan_simple(control, store, problems)


def _evaluate(control: _Control, context: "_Element", expand_final: bool = False):
    """Nodes matching the control XPath, or a reason string when not evaluable."""
    nsmap = parse_prefix_mappings(control.prefix_mappings)
    try:
        return evaluate_xpath_subset(
            context, control.xpath, nsmap, ignore_final_predicate=expand_final
        )
    except UnsupportedXPathError as exc:
        return str(exc)


def _plan_simple(
    control: _Control, store: "CustomXmlStore", problems: List[BindingProblem]
) -> "Tuple[PlanEntry | None, Any]":
    """Plan the refresh of a plain-text or checkbox control."""
    nodes = _evaluate(control, store.root)
    if isinstance(nodes, str):
        problems.append(control.problem(nodes))
        return None, None
    if len(nodes) != 1:
        problems.append(
            control.problem(
                f"XPath matched {len(nodes)} nodes; exactly 1 is required for a"
                f" {control.kind} control"
            )
        )
        return None, None
    text = element_text(nodes[0])
    if control.kind == CHECKBOX:
        checked = _parse_checkbox_value(text)
        if checked is None:
            problems.append(
                control.problem(
                    f"value {text!r} cannot be interpreted as a checkbox state;"
                    " expected one of true/false, 1/0, yes/no, on/off"
                )
            )
            return None, None
        entry = PlanEntry(
            control_id=control.control_id,
            tag=control.tag,
            kind=CHECKBOX,
            store_item_id=control.store_item_id,
            xpath=control.xpath,
            action="set checked" if checked else "set unchecked",
            value="true" if checked else "false",
        )
        return entry, (CHECKBOX, control.sdt, checked)
    entry = PlanEntry(
        control_id=control.control_id,
        tag=control.tag,
        kind=PLAIN_TEXT,
        store_item_id=control.store_item_id,
        xpath=control.xpath,
        action=f"set text to {text!r}",
        value=text,
    )
    return entry, (PLAIN_TEXT, control.sdt, text)


def _parse_checkbox_value(text: str) -> "bool | None":
    """Boolean state for checkbox bound 'text', or |None| when not recognized."""
    normalized = text.strip().lower()
    if normalized in _TRUE_VALUES:
        return True
    if normalized in _FALSE_VALUES:
        return False
    return None


def _section_items(content: "_Element") -> "List[_Element]":
    """The repeating section item w:sdt elements directly contained in 'content'."""
    items = []
    for child in content:
        if child.tag != qn("w:sdt"):
            continue
        sdtPr = child.find(qn("w:sdtPr"))
        if sdtPr is not None and sdtPr.find(qn("w15:repeatingSectionItem")) is not None:
            items.append(child)
    return items


def _bound_child_controls(item_sdt: "_Element") -> List[_Control]:
    """Bound controls within a repeating section item, in document order.

    Controls inside a nested repeating section are excluded; the nested section is
    reported as its own bound child so planning can flag it as unsupported.
    """
    controls = []
    for sdt in item_sdt.iter(qn("w:sdt")):
        ancestor = sdt.getparent()
        skipped = False
        while ancestor is not None and ancestor is not item_sdt:
            if _is_repeating_section_sdt(ancestor):
                skipped = True
                break
            ancestor = ancestor.getparent()
        if skipped:
            continue
        control = _Control(sdt)
        if control.bound:
            controls.append(control)
    return controls


def _plan_section(
    control: _Control, store: "CustomXmlStore", problems: List[BindingProblem]
) -> "Tuple[PlanEntry | None, Any]":
    """Plan the expansion of a repeating section and the fill of each item."""
    nodes = _evaluate(control, store.root, expand_final=True)
    if isinstance(nodes, str):
        problems.append(control.problem(nodes))
        return None, None
    if not nodes:
        problems.append(
            control.problem("XPath matched 0 nodes; a repeating section needs at least 1")
        )
        return None, None
    content = control.sdt.find(qn("w:sdtContent"))
    items = _section_items(content) if content is not None else []
    if not items:
        problems.append(
            control.problem("repeating section has no item to use as an expansion template")
        )
        return None, None

    child_kinds: List[str] = []
    child_values: List[list] = []
    ok = True
    for child in _bound_child_controls(items[0]):
        if child.kind in (None, REPEATING_SECTION):
            problems.append(
                child.problem("control type is not supported inside a repeating section item")
            )
            ok = False
            continue
        values = []
        for item_index, node in enumerate(nodes):
            child_nodes = _evaluate(child, node)
            if isinstance(child_nodes, str):
                problems.append(child.problem(child_nodes, item_index=item_index))
                ok = False
                break
            if len(child_nodes) != 1:
                problems.append(
                    child.problem(
                        f"XPath matched {len(child_nodes)} nodes; exactly 1 is required"
                        f" for a {child.kind} control",
                        item_index=item_index,
                    )
                )
                ok = False
                break
            text = element_text(child_nodes[0])
            if child.kind == CHECKBOX:
                checked = _parse_checkbox_value(text)
                if checked is None:
                    problems.append(
                        child.problem(
                            f"value {text!r} cannot be interpreted as a checkbox state",
                            item_index=item_index,
                        )
                    )
                    ok = False
                    break
                values.append(checked)
            else:
                values.append(text)
        child_kinds.append(child.kind)
        child_values.append(values)

    entry = PlanEntry(
        control_id=control.control_id,
        tag=control.tag,
        kind=REPEATING_SECTION,
        store_item_id=control.store_item_id,
        xpath=control.xpath,
        action=f"expand to {len(nodes)} item(s)",
        item_count=len(nodes),
    )
    payload = (REPEATING_SECTION, control.sdt, len(nodes), child_kinds, child_values)
    return (entry, payload) if ok else (entry, None)


# ============================================================================
# apply internals
# ============================================================================


class _SdtIdAllocator:
    """Allocates w:id values for content controls, unique within the document."""

    def __init__(self, root: "_Element"):
        self._used = set()
        for sdtPr in root.iter(qn("w:sdtPr")):
            id_el = sdtPr.find(qn("w:id"))
            if id_el is None:
                continue
            try:
                self._used.add(int(id_el.get(qn("w:val"), "")))
            except ValueError:
                continue
        self._next = max(self._used) + 1 if self._used else 1

    def alloc(self) -> int:
        """The next unused control ID."""
        while self._next in self._used:
            self._next += 1
        value = self._next
        self._used.add(value)
        self._next += 1
        return value


def _apply_entry(entry: PlanEntry, payload: Any, allocator: _SdtIdAllocator) -> None:
    """Commit a single planned refresh action to the document."""
    if payload is None:
        return
    kind = payload[0]
    if kind == PLAIN_TEXT:
        _, sdt, text = payload
        _set_sdt_text(sdt, text)
    elif kind == CHECKBOX:
        _, sdt, checked = payload
        _set_sdt_checkbox(sdt, checked)
    elif kind == REPEATING_SECTION:
        _, sdt, item_count, child_kinds, child_values = payload
        _expand_section(sdt, item_count, child_kinds, child_values, allocator)


def _make_run(text: str, rPr: "_Element | None") -> "_Element":
    """A new w:r element containing 'text', optionally with run properties 'rPr'."""
    run = OxmlElement("w:r")
    if rPr is not None:
        run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    run.append(t)
    return run


def _prototype_rPr(content: "_Element") -> "_Element | None":
    """A deep copy of the first run properties in 'content', to preserve formatting."""
    first_run = content.find(f".//{qn('w:r')}")
    if first_run is None:
        return None
    rPr = first_run.find(qn("w:rPr"))
    return copy.deepcopy(rPr) if rPr is not None else None


def _set_sdt_text(sdt: "_Element", text: str) -> None:
    """Replace the displayed content of a plain-text control with 'text'."""
    content = sdt.find(qn("w:sdtContent"))
    if content is None:
        return
    rPr = _prototype_rPr(content)
    paragraphs = content.findall(qn("w:p"))
    if paragraphs:
        target, extras = paragraphs[0], paragraphs[1:]
        for extra in extras:
            content.remove(extra)
        for child in list(target):
            if child.tag != qn("w:pPr"):
                target.remove(child)
        target.append(_make_run(text, rPr))
    else:
        for child in list(content):
            content.remove(child)
        content.append(_make_run(text, rPr))


def _state_glyph(checkbox: "_Element", checked: bool) -> str:
    """The display character for the checked or unchecked state of 'checkbox'."""
    state_tag = "w14:checkedState" if checked else "w14:uncheckedState"
    state_el = checkbox.find(qn(state_tag))
    if state_el is not None:
        val = state_el.get(qn("w14:val"))
        if val:
            try:
                return chr(int(val, 16))
            except ValueError:
                pass
    return "☒" if checked else "☐"


def _set_sdt_checkbox(sdt: "_Element", checked: bool) -> None:
    """Update the state element and displayed glyph of a checkbox control."""
    sdtPr = sdt.find(qn("w:sdtPr"))
    checkbox = sdtPr.find(qn("w14:checkbox")) if sdtPr is not None else None
    if checkbox is None:
        return
    checked_el = checkbox.find(qn("w14:checked"))
    if checked_el is None:
        checked_el = OxmlElement("w14:checked")
        checkbox.insert(0, checked_el)
    checked_el.set(qn("w14:val"), "1" if checked else "0")
    glyph = _state_glyph(checkbox, checked)
    content = sdt.find(qn("w:sdtContent"))
    if content is None:
        return
    texts = content.findall(f".//{qn('w:t')}")
    if texts:
        texts[0].text = glyph
        for extra in texts[1:]:
            extra.text = ""
    else:
        content.append(_make_run(glyph, None))


def _reassign_ids(sdt: "_Element", allocator: _SdtIdAllocator) -> None:
    """Give every control within 'sdt' (including itself) a fresh unique w:id."""
    for sdtPr in sdt.iter(qn("w:sdtPr")):
        id_el = sdtPr.find(qn("w:id"))
        if id_el is None:
            id_el = OxmlElement("w:id")
            sdtPr.append(id_el)
        id_el.set(qn("w:val"), str(allocator.alloc()))


def _expand_section(
    sdt: "_Element",
    item_count: int,
    child_kinds: List[str],
    child_values: List[list],
    allocator: _SdtIdAllocator,
) -> None:
    """Expand a repeating section to 'item_count' items and fill each from the plan.

    All existing items are replaced by fresh clones of the first item, so applying
    the same plan repeatedly does not accumulate duplicate items.
    """
    content = sdt.find(qn("w:sdtContent"))
    items = _section_items(content)
    template = copy.deepcopy(items[0])
    anchor = list(content).index(items[0])
    for item in items:
        content.remove(item)
    for index in range(item_count):
        clone = copy.deepcopy(template)
        _reassign_ids(clone, allocator)
        content.insert(anchor + index, clone)
        children = _bound_child_controls(clone)
        for child, kind, values in zip(children, child_kinds, child_values):
            if kind == PLAIN_TEXT:
                _set_sdt_text(child.sdt, values[index])
            elif kind == CHECKBOX:
                _set_sdt_checkbox(child.sdt, values[index])
