"""Content controls ('w:sdt') and custom-XML data-binding refresh.

Some templates bind content controls to custom XML parts. Editing the displayed
text in 'document.xml' directly is pointless there because Word overwrites it
from the XML data store on open. This module provides the counterpart workflow:
the caller updates the custom XML data, then explicitly refreshes the displayed
content of the bound controls.

Reading and rewriting are separated: 'ContentControls.plan_refresh()' computes a
|RefreshPlan| that can be inspected ('plan.actions', 'plan.errors') before any
change is committed with 'plan.apply()'.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum
from typing import TYPE_CHECKING, Any, Dict, List, Mapping

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.oxml.sdt import CT_Sdt

if TYPE_CHECKING:
    from docx.oxml.document import CT_Document
    from docx.parts.customxml import CustomXmlPart
    from docx.parts.document import DocumentPart


class BindingErrorKind(Enum):
    """The category of a |BindingError|."""

    STORE_NOT_FOUND = "store-not-found"
    UNSUPPORTED_XPATH = "unsupported-xpath"
    RESULT_COUNT = "result-count"
    TYPE_MISMATCH = "type-mismatch"
    UNSUPPORTED_CONTROL = "unsupported-control"


@dataclass
class BindingError:
    """A problem found while planning a refresh, tied to a specific control."""

    kind: BindingErrorKind
    control: str
    message: str
    xpath: str | None = None

    def __str__(self) -> str:
        return "[%s] %s: %s" % (self.kind.value, self.control, self.message)


class ContentControlRefreshError(Exception):
    """Raised when 'RefreshPlan.apply()' is called on a plan that has errors."""

    def __init__(self, errors: List[BindingError]):
        self.errors = list(errors)
        summary = "; ".join(str(e) for e in self.errors)
        super().__init__(
            "content-control refresh plan has %d error(s): %s" % (len(self.errors), summary)
        )


class XPathNotSupportedError(ValueError):
    """Raised when a binding XPath falls outside the supported subset."""


# ---------------------------------------------------------------------------
# restricted XPath subset

_STEP_RE = re.compile(r"^(?:([A-Za-z_][\w.\-]*):)?([A-Za-z_][\w.\-]*)(?:\[(\d+)\])?$")
_ATTR_RE = re.compile(r"^@(?:([A-Za-z_][\w.\-]*):)?([A-Za-z_][\w.\-]*)$")
_PREFIX_MAPPING_RE = re.compile(r"xmlns:([A-Za-z_][\w.\-]*)\s*=\s*(['\"])(.*?)\2")


def parse_prefix_mappings(raw: str | None) -> Dict[str, str]:
    """Parse a 'w:prefixMappings' value into a prefix-to-namespace-uri dict.

    The attribute has the form 'xmlns:p1=\'uri1\' xmlns:p2=\'uri2\''.
    """
    if not raw:
        return {}
    return {m.group(1): m.group(3) for m in _PREFIX_MAPPING_RE.finditer(raw)}


@dataclass
class _Step:
    """A single child-axis step, e.g. 'ns:item' or 'ns:item[2]'."""

    qname: str
    index: int | None


class SimpleXPath:
    """A restricted XPath expression evaluated without a full XPath engine.

    Supported forms are absolute or relative sequences of child-axis element
    steps with an optional 1-based positional predicate ('/ns:root/ns:item[2]'),
    optionally ending in a single attribute step ('/ns:root/ns:item/@ns:code').
    Anything else (descendant axis, '..', wildcards, functions, other
    predicates) raises |XPathNotSupportedError|.
    """

    def __init__(
        self,
        steps: List[_Step],
        attribute: str | None,
        absolute: bool,
        source: str,
    ):
        self._steps = steps
        self._attribute = attribute
        self._absolute = absolute
        self._source = source

    def __str__(self) -> str:
        return self._source

    @property
    def is_absolute(self) -> bool:
        return self._absolute

    @classmethod
    def parse(cls, source: str, nsmap: Mapping[str, str]) -> "SimpleXPath":
        """Parse 'source' using 'nsmap' to resolve namespace prefixes."""

        def qname_of(prefix: str | None, name: str) -> str:
            if prefix is None:
                return name
            if prefix not in nsmap:
                raise XPathNotSupportedError(
                    "unknown namespace prefix '%s' in XPath '%s'" % (prefix, source)
                )
            return "{%s}%s" % (nsmap[prefix], name)

        text = (source or "").strip()
        if not text:
            raise XPathNotSupportedError("empty XPath expression")
        absolute = text.startswith("/")
        body = text[1:] if absolute else text
        raw_steps = body.split("/")
        if any(not step for step in raw_steps):
            raise XPathNotSupportedError("unsupported XPath expression '%s'" % source)

        attribute = None
        attr_match = _ATTR_RE.match(raw_steps[-1])
        if attr_match:
            attribute = qname_of(attr_match.group(1), attr_match.group(2))
            raw_steps.pop()
            if not raw_steps:
                raise XPathNotSupportedError("unsupported XPath expression '%s'" % source)

        steps = []
        for raw in raw_steps:
            match = _STEP_RE.match(raw)
            if match is None:
                raise XPathNotSupportedError(
                    "unsupported step '%s' in XPath '%s'" % (raw, source)
                )
            index = int(match.group(3)) if match.group(3) else None
            if index is not None and index < 1:
                raise XPathNotSupportedError("positional predicate in '%s' is 1-based" % source)
            steps.append(_Step(qname_of(match.group(1), match.group(2)), index))
        return cls(steps, attribute, absolute, source)

    def select(self, context: Any) -> list:
        """Evaluate against 'context', an element of the custom XML store.

        Returns a list of matching elements, or a list of strings when the
        expression ends in an attribute step.
        """
        steps = self._steps
        if self._absolute:
            first = steps[0]
            if context.tag != first.qname or (first.index or 1) != 1:
                return []
            nodes = [context]
            steps = steps[1:]
        else:
            nodes = [context]
        for step in steps:
            matched = []
            for node in nodes:
                children = [c for c in node if c.tag == step.qname]
                if step.index is None:
                    matched.extend(children)
                elif step.index <= len(children):
                    matched.append(children[step.index - 1])
            nodes = matched
        if self._attribute is not None:
            return [n.get(self._attribute) for n in nodes if n.get(self._attribute) is not None]
        return nodes

    def select_elements(self, context: Any) -> list:
        """Like 'select()', but an attribute step yields the parent elements."""
        if self._attribute is None:
            return self.select(context)
        element_path = SimpleXPath(self._steps, None, self._absolute, self._source)
        return element_path.select(context)

    def relative_to(self, base: "SimpleXPath") -> "SimpleXPath | None":
        """Remainder of this path when it extends 'base', otherwise |None|.

        Step names are compared without positional predicates, so
        '/r/item[3]/name' is an extension of '/r/item'.
        """
        if not (self._absolute and base._absolute):
            return None
        if len(self._steps) < len(base._steps):
            return None
        if len(self._steps) == len(base._steps) and self._attribute is None:
            return None
        for step, base_step in zip(self._steps, base._steps):
            if step.qname != base_step.qname:
                return None
        rest = self._steps[len(base._steps):]
        return SimpleXPath(rest, self._attribute, False, self._source)


# ---------------------------------------------------------------------------
# data binding and custom XML stores


class DataBinding:
    """A 'w:dataBinding' element, linking a content control to custom XML data."""

    def __init__(self, element: Any):
        self._element = element

    @property
    def store_item_id(self) -> str | None:
        """Value of 'w:storeItemID', identifying the custom XML part."""
        return self._element.get(qn("w:storeItemID"))

    @property
    def xpath(self) -> str | None:
        """Value of 'w:xpath', selecting the bound node within the store."""
        return self._element.get(qn("w:xpath"))

    @property
    def prefix_mappings(self) -> str:
        """Raw value of 'w:prefixMappings', the empty string when absent."""
        return self._element.get(qn("w:prefixMappings")) or ""

    @property
    def nsmap(self) -> Dict[str, str]:
        """Prefix-to-namespace-uri mapping parsed from 'w:prefixMappings'."""
        return parse_prefix_mappings(self.prefix_mappings)

    def compile(self) -> SimpleXPath:
        """The binding XPath parsed into the supported subset."""
        return SimpleXPath.parse(self.xpath or "", self.nsmap)


class CustomXmlStore:
    """A custom XML data store (one 'customXml/itemN.xml' part) in the package."""

    def __init__(self, item_id: str, part: "CustomXmlPart"):
        self.item_id = item_id
        self._part = part

    def __repr__(self) -> str:
        return "CustomXmlStore(%r, partname=%s)" % (self.item_id, self._part.partname)

    @property
    def element(self) -> Any:
        """Root element of the custom XML part; edits are saved with the package."""
        return self._part.element

    def select(self, xpath: str, prefix_mappings: str = "") -> list:
        """Evaluate restricted-XPath 'xpath' against this store, return matches."""
        path = SimpleXPath.parse(xpath, parse_prefix_mappings(prefix_mappings))
        return path.select(self.element)

    def get_value(self, xpath: str, prefix_mappings: str = "") -> str:
        """Text of the single node selected by 'xpath'.

        Raises |ValueError| when the path does not select exactly one node.
        """
        results = self.select(xpath, prefix_mappings)
        if len(results) != 1:
            raise ValueError(
                "XPath '%s' matched %d nodes, expected exactly 1" % (xpath, len(results))
            )
        node = results[0]
        return node if isinstance(node, str) else (node.text or "")

    def set_value(self, xpath: str, value: str, prefix_mappings: str = "") -> None:
        """Set the text of the single node selected by 'xpath' to 'value'.

        Raises |ValueError| when the path does not select exactly one node.
        """
        path = SimpleXPath.parse(xpath, parse_prefix_mappings(prefix_mappings))
        results = path.select_elements(self.element)
        if len(results) != 1:
            raise ValueError(
                "XPath '%s' matched %d nodes, expected exactly 1" % (xpath, len(results))
            )
        node = results[0]
        if path._attribute is not None:
            node.set(path._attribute, value)
        else:
            node.text = value


class CustomXmlStores(Mapping):
    """Mapping of custom XML stores in a package, keyed by 'ds:itemID'."""

    def __init__(self, stores: Dict[str, CustomXmlStore]):
        self._stores = dict(stores)

    def __getitem__(self, item_id: str) -> CustomXmlStore:
        return self._stores[item_id]

    def __iter__(self) -> Any:
        return iter(self._stores)

    def __len__(self) -> int:
        return len(self._stores)

    @classmethod
    def discover(cls, document_part: "DocumentPart") -> "CustomXmlStores":
        """All custom XML stores related from 'document_part', keyed by item id."""
        stores = {}
        for rel in document_part.rels.values():
            if rel.is_external or rel.reltype != RT.CUSTOM_XML:
                continue
            part = rel.target_part
            item_id = cls._item_id_for(part)
            if item_id is not None:
                stores[item_id] = CustomXmlStore(item_id, part)
        return cls(stores)

    @staticmethod
    def _item_id_for(part: "CustomXmlPart") -> str | None:
        """The 'ds:itemID' of 'part', read from its item-properties part."""
        for rel in part.rels.values():
            if rel.is_external or rel.reltype != RT.CUSTOM_XML_PROPS:
                continue
            props = parse_xml(rel.target_part.blob)
            item_id = props.get(qn("ds:itemID"))
            if item_id:
                return item_id
        return None


# ---------------------------------------------------------------------------
# content control proxies


class ControlType(Enum):
    """The kind of a content control, as far as refresh is concerned."""

    PLAIN_TEXT = "plain-text"
    CHECKBOX = "checkbox"
    REPEATING_SECTION = "repeating-section"
    REPEATING_SECTION_ITEM = "repeating-section-item"
    OTHER = "other"


class ContentControl:
    """A single content control ('w:sdt' element) in the document."""

    def __init__(self, sdt: CT_Sdt):
        self._sdt = sdt

    @property
    def id(self) -> int | None:
        """The 'w:id' of this control, or |None| if not present."""
        return self._sdt.sdt_id

    @property
    def alias(self) -> str | None:
        """The title of this control, or |None| if it has none."""
        return self._sdt.alias

    @property
    def tag(self) -> str | None:
        """The tag of this control, or |None| if it has none."""
        return self._sdt.sdt_tag

    @property
    def control_type(self) -> ControlType:
        """Member of |ControlType| describing this control."""
        sdt = self._sdt
        if sdt.is_checkbox:
            return ControlType.CHECKBOX
        if sdt.is_repeating_section:
            return ControlType.REPEATING_SECTION
        if sdt.is_repeating_section_item:
            return ControlType.REPEATING_SECTION_ITEM
        if sdt.is_plain_text:
            return ControlType.PLAIN_TEXT
        return ControlType.OTHER

    @property
    def binding(self) -> DataBinding | None:
        """|DataBinding| for this control, or |None| when it is unbound."""
        element = self._sdt.binding
        return DataBinding(element) if element is not None else None

    @property
    def text(self) -> str:
        """The currently displayed text of this control."""
        return self._sdt.display_text()

    @property
    def label(self) -> str:
        """Human-readable identification of this control for messages."""
        parts = []
        if self.alias:
            parts.append("alias '%s'" % self.alias)
        if self.tag:
            parts.append("tag '%s'" % self.tag)
        parts.append("id %s" % (self.id if self.id is not None else "?"))
        return "content control (%s)" % ", ".join(parts)


# ---------------------------------------------------------------------------
# refresh plan


@dataclass
class PlannedAction:
    """A single document rewrite planned by |RefreshPlan|.

    'kind' is one of 'set-text', 'set-checkbox', or 'expand-section'. 'detail'
    describes the change in human-readable form, e.g. the old and new value.
    """

    control: str
    kind: str
    detail: str
    apply_fn: Any = None

    def __str__(self) -> str:
        return "%s %s: %s" % (self.kind, self.control, self.detail)


class RefreshPlan:
    """The result of 'ContentControls.plan_refresh()'.

    Reading ('plan_refresh()') and rewriting ('apply()') are separated so the
    caller can inspect 'actions' and 'errors' before committing any change.
    'apply()' raises |ContentControlRefreshError| when the plan has errors, so
    a problematic control is always reported before any change is committed.
    """

    def __init__(self, actions: List[PlannedAction], errors: List[BindingError]):
        self.actions = list(actions)
        self.errors = list(errors)

    def __len__(self) -> int:
        return len(self.actions)

    def __iter__(self) -> Any:
        return iter(self.actions)

    @property
    def ok(self) -> bool:
        """True when no errors were found while planning."""
        return not self.errors

    def apply(self) -> None:
        """Commit all planned actions to the document.

        Raises |ContentControlRefreshError| when 'errors' is not empty; no
        change is made in that case.
        """
        if self.errors:
            raise ContentControlRefreshError(self.errors)
        for action in self.actions:
            action.apply_fn()


def _parse_checkbox_value(text: str) -> bool | None:
    """Boolean value of checkbox-bound 'text', or |None| when not boolean-like."""
    value = text.strip().lower()
    if value in ("1", "true"):
        return True
    if value in ("0", "false"):
        return False
    return None


def _bound_child_sdts(section_sdt: CT_Sdt) -> List[CT_Sdt]:
    """Bound 'w:sdt' descendants of a repeating-section template, in document order.

    Controls nested inside a nested repeating section are excluded; nested
    repeating sections are reported as unsupported by the planner.
    """
    result = []
    for child in section_sdt.iter_sdt():
        ancestor = child.getparent()
        nested = False
        while ancestor is not None and ancestor is not section_sdt:
            if isinstance(ancestor, CT_Sdt) and (
                ancestor.is_repeating_section or ancestor.is_repeating_section_item
            ):
                nested = True
                break
            ancestor = ancestor.getparent()
        if not nested and child.binding is not None:
            result.append(child)
    return result


class _RefreshPlanner:
    """Computes a |RefreshPlan| for the controls of one document."""

    def __init__(self, stores: CustomXmlStores):
        self._stores = stores
        self.actions: List[PlannedAction] = []
        self.errors: List[BindingError] = []
        self._used_ids_cache: "set | None" = None
        self._document_element: "CT_Document | None" = None

    def plan(self, document_element: "CT_Document") -> RefreshPlan:
        """Plan a refresh of all bound controls in 'document_element'."""
        self._document_element = document_element
        for sdt in document_element.iter(qn("w:sdt")):
            if self._is_inside_repeating_section(sdt):
                continue
            if sdt.is_repeating_section_item:
                continue
            self._plan_control(sdt, item_node=None, section_path=None, label_prefix="")
        return RefreshPlan(self.actions, self.errors)

    @staticmethod
    def _is_inside_repeating_section(sdt: CT_Sdt) -> bool:
        ancestor = sdt.getparent()
        while ancestor is not None:
            if isinstance(ancestor, CT_Sdt) and (
                ancestor.is_repeating_section or ancestor.is_repeating_section_item
            ):
                return True
            ancestor = ancestor.getparent()
        return False

    def _error(
        self, kind: BindingErrorKind, label: str, message: str, xpath: str | None = None
    ) -> None:
        self.errors.append(BindingError(kind, label, message, xpath))

    def _plan_control(
        self,
        sdt: CT_Sdt,
        item_node: Any,
        section_path: "SimpleXPath | None",
        label_prefix: str,
    ) -> None:
        control = ContentControl(sdt)
        label = label_prefix + control.label
        binding = control.binding
        if binding is None:
            return
        if sdt.is_repeating_section:
            self._plan_section(sdt, binding, label)
            return
        update = self._single_value_update(sdt, item_node, section_path, label)
        if update is None:
            return
        kind, value = update
        if kind == "checkbox":
            self.actions.append(
                PlannedAction(
                    label,
                    "set-checkbox",
                    "set checked to %s" % value,
                    apply_fn=lambda sdt=sdt, value=value: sdt.set_checkbox_state(value),
                )
            )
        else:
            self.actions.append(
                PlannedAction(
                    label,
                    "set-text",
                    "set text to %r" % value,
                    apply_fn=lambda sdt=sdt, value=value: sdt.set_display_text(value),
                )
            )

    def _single_value_update(
        self,
        sdt: CT_Sdt,
        item_node: Any,
        section_path: "SimpleXPath | None",
        label: str,
    ) -> "tuple[str, Any] | None":
        """('text', str) or ('checkbox', bool) for a bound single-value control.

        Returns |None| for unbound or unsupported controls, and after recording
        an error for bindings that cannot be resolved.
        """
        binding = ContentControl(sdt).binding
        if binding is None:
            return None
        if not (sdt.is_plain_text or sdt.is_checkbox):
            return None
        value = self._resolve_value(binding, item_node, section_path, label)
        if value is None:
            return None
        if sdt.is_checkbox:
            checked = _parse_checkbox_value(value)
            if checked is None:
                self._error(
                    BindingErrorKind.TYPE_MISMATCH,
                    label,
                    "checkbox binding value %r is not one of 1/0/true/false" % value,
                    binding.xpath,
                )
                return None
            return ("checkbox", checked)
        return ("text", value)

    def _resolve_value(
        self,
        binding: DataBinding,
        item_node: Any,
        section_path: "SimpleXPath | None",
        label: str,
    ) -> "str | None":
        """Bound text value for a single-value control, or |None| after an error."""
        resolved = self._resolve_path(binding, item_node, section_path, label)
        if resolved is None:
            return None
        path, context = resolved
        results = path.select(context)
        if len(results) != 1:
            self._error(
                BindingErrorKind.RESULT_COUNT,
                label,
                "XPath matched %d nodes, expected exactly 1" % len(results),
                binding.xpath,
            )
            return None
        node = results[0]
        if isinstance(node, str):
            return node
        if len(node):
            self._error(
                BindingErrorKind.TYPE_MISMATCH,
                label,
                "bound element has child elements, expected text content",
                binding.xpath,
            )
            return None
        return node.text or ""

    def _resolve_path(
        self,
        binding: DataBinding,
        item_node: Any,
        section_path: "SimpleXPath | None",
        label: str,
    ) -> "tuple[SimpleXPath, Any] | None":
        """(path, context-node) pair for 'binding', or |None| after an error.

        Inside a repeating-section item, a relative XPath is evaluated from the
        item node; an absolute XPath that extends the section path is rebased
        onto the item node; any other absolute XPath is evaluated from the
        store root, so data bound in several places is still updated together.
        """
        store = self._store_for(binding, label)
        if store is None:
            return None
        try:
            path = binding.compile()
        except XPathNotSupportedError as exc:
            self._error(BindingErrorKind.UNSUPPORTED_XPATH, label, str(exc), binding.xpath)
            return None
        if item_node is not None:
            if not path.is_absolute:
                return path, item_node
            if section_path is not None:
                rebased = path.relative_to(section_path)
                if rebased is not None:
                    return rebased, item_node
        return path, store.element

    def _store_for(self, binding: DataBinding, label: str) -> "CustomXmlStore | None":
        item_id = binding.store_item_id
        if not item_id:
            self._error(
                BindingErrorKind.STORE_NOT_FOUND,
                label,
                "binding has no storeItemID",
                binding.xpath,
            )
            return None
        store = self._stores.get(item_id)
        if store is None:
            self._error(
                BindingErrorKind.STORE_NOT_FOUND,
                label,
                "no custom XML part found with itemID '%s'" % item_id,
                binding.xpath,
            )
        return store

    def _plan_section(self, sdt: CT_Sdt, binding: DataBinding, label: str) -> None:
        store = self._store_for(binding, label)
        if store is None:
            return
        try:
            path = binding.compile()
        except XPathNotSupportedError as exc:
            self._error(BindingErrorKind.UNSUPPORTED_XPATH, label, str(exc), binding.xpath)
            return
        if not path.is_absolute:
            self._error(
                BindingErrorKind.UNSUPPORTED_XPATH,
                label,
                "repeating-section binding XPath must be absolute",
                binding.xpath,
            )
            return
        nodes = path.select(store.element)
        if len(nodes) == 0:
            self._error(
                BindingErrorKind.RESULT_COUNT,
                label,
                "section XPath matched 0 nodes, expected at least 1",
                binding.xpath,
            )
            return
        if any(isinstance(node, str) for node in nodes):
            self._error(
                BindingErrorKind.TYPE_MISMATCH,
                label,
                "section binding must select elements, not an attribute",
                binding.xpath,
            )
            return
        for child in sdt.iter_sdt():
            if child.is_repeating_section or child.is_repeating_section_item:
                self._error(
                    BindingErrorKind.UNSUPPORTED_CONTROL,
                    label,
                    "nested repeating sections are not supported",
                )
                return
        children = _bound_child_sdts(sdt)
        item_updates = []
        for index, node in enumerate(nodes):
            updates = []
            prefix = "%s [item %d] >> " % (label, index + 1)
            for position, child in enumerate(children):
                child_label = prefix + ContentControl(child).label
                update = self._single_value_update(child, node, path, child_label)
                if update is not None:
                    kind, value = update
                    updates.append((position, kind, value))
            item_updates.append(updates)
        self.actions.append(
            PlannedAction(
                label,
                "expand-section",
                "expand to %d item(s)" % len(nodes),
                apply_fn=lambda sdt=sdt, item_updates=item_updates: self._apply_section(
                    sdt, item_updates
                ),
            )
        )

    def _apply_section(self, sdt: CT_Sdt, item_updates: list) -> None:
        """Expand a repeating-section template to one control per update list.

        Previously generated item copies following the template are removed
        first, so applying a refresh twice does not accumulate duplicates.
        """
        sibling = sdt.getnext()
        while (
            sibling is not None
            and isinstance(sibling, CT_Sdt)
            and sibling.is_repeating_section_item
        ):
            next_sibling = sibling.getnext()
            sdt.getparent().remove(sibling)
            sibling = next_sibling
        self._apply_item_updates(_bound_child_sdts(sdt), item_updates[0])
        used_ids = self._used_ids()
        anchor = sdt
        for updates in item_updates[1:]:
            clone = sdt.clone()
            clone.mark_as_repeating_item()
            clone.renew_ids(used_ids)
            anchor.addnext(clone)
            anchor = clone
            self._apply_item_updates(_bound_child_sdts(clone), updates)

    @staticmethod
    def _apply_item_updates(children: List[CT_Sdt], updates: list) -> None:
        for position, kind, value in updates:
            child = children[position]
            if kind == "checkbox":
                child.set_checkbox_state(value)
            else:
                child.set_display_text(value)

    def _used_ids(self) -> set:
        """All 'w:id' values currently in the document, as a mutable set."""
        if self._used_ids_cache is None:
            values = self._document_element.xpath(".//w:sdtPr/w:id/@w:val")
            self._used_ids_cache = {int(v) for v in values} or {0}
        return self._used_ids_cache


# ---------------------------------------------------------------------------
# content controls collection


class ContentControls:
    """The content controls ('w:sdt' elements) of a document, in document order.

    Accessed via 'Document.content_controls'. Call 'plan_refresh()' after
    updating custom XML data to compute a |RefreshPlan|, inspect it, then
    commit it with 'RefreshPlan.apply()'.
    """

    def __init__(self, document_element: "CT_Document", part: "DocumentPart"):
        self._element = document_element
        self._part = part

    def __iter__(self) -> Any:
        return (ContentControl(sdt) for sdt in self._element.iter(qn("w:sdt")))

    def __len__(self) -> int:
        return sum(1 for _ in self._element.iter(qn("w:sdt")))

    @property
    def custom_xml_stores(self) -> CustomXmlStores:
        """|CustomXmlStores| holding the custom XML data of this document."""
        return self._part.custom_xml_stores

    def plan_refresh(self) -> RefreshPlan:
        """Compute a |RefreshPlan| reflecting the current custom XML data.

        No document change is made until 'RefreshPlan.apply()' is called.
        """
        return _RefreshPlanner(self.custom_xml_stores).plan(self._element)

    def refresh(self) -> None:
        """Plan and immediately apply a refresh.

        Raises |ContentControlRefreshError| when any bound control has a
        problem; no change is committed in that case.
        """
        self.plan_refresh().apply()
