"""Custom XML data-store access for content-control data binding.

Word content controls (w:sdt elements) can be bound to nodes in a custom XML part
(/customXml/itemN.xml) via a w:dataBinding element. This module provides
|CustomXmlStore|, a read/write view of such a part, along with the restricted XPath
subset used to address nodes within it.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Dict, List, NamedTuple

from docx.exceptions import UnsupportedXPathError
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.opc.part import Part

# -- a step is an element name, optionally followed by a 1-based positional predicate --
_STEP_RE = re.compile(r"^(?P<name>(?:[A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*)(?:\[(?P<idx>\d+)\])?$")

# -- prefix mappings as serialized by Word, e.g. "xmlns:ns0='http://x' xmlns:a='http://y'" --
_PREFIX_MAPPING_RE = re.compile(r"xmlns:([A-Za-z_][\w.-]*)\s*=\s*(['\"])(.*?)\2")


class _Step(NamedTuple):
    """A single parsed XPath step like 'ns0:item' or 'item[2]'."""

    name: str
    index: int | None


def parse_prefix_mappings(prefix_mappings: str | None) -> Dict[str, str]:
    """Return prefix-to-namespace-URI mapping parsed from `prefix_mappings`.

    Uses the form Word writes into w:dataBinding/@w:prefixMappings, e.g.
    "xmlns:ns0='http://example.com' xmlns:ns1='http://other.com'". An empty or |None|
    value produces an empty mapping.
    """
    if not prefix_mappings:
        return {}
    return {m.group(1): m.group(3) for m in _PREFIX_MAPPING_RE.finditer(prefix_mappings)}


def parse_xpath_subset(xpath: str) -> "tuple[bool, List[_Step]]":
    """Parse `xpath` into an (is_absolute, steps) pair.

    Only a deliberately small XPath subset is supported, covering the expressions Word
    writes into w:dataBinding/@w:xpath: an optional leading '/', then one or more
    child-axis steps of the form 'name' or 'name[n]' where n is a 1-based position.
    Names may carry a namespace prefix.

    Raises |UnsupportedXPathError| for anything else, such as '//' descent, attribute
    steps, wildcard names, or non-positional predicates.
    """
    text = (xpath or "").strip()
    if not text:
        raise UnsupportedXPathError("XPath expression is empty")
    if text.startswith("//"):
        raise UnsupportedXPathError(f"descendant axis '//' is not supported in XPath {xpath!r}")
    is_absolute = text.startswith("/")
    body = text[1:] if is_absolute else text
    if not body:
        raise UnsupportedXPathError(f"XPath {xpath!r} selects no element steps")
    steps: List[_Step] = []
    for raw_step in body.split("/"):
        m = _STEP_RE.match(raw_step)
        if m is None:
            raise UnsupportedXPathError(
                f"step {raw_step!r} in XPath {xpath!r} is outside the supported subset;"
                " only 'name' and 'name[n]' child steps are supported"
            )
        index = int(m.group("idx")) if m.group("idx") is not None else None
        steps.append(_Step(m.group("name"), index))
    return is_absolute, steps


def _name_matches(element: _Element, name: str, nsmap: Dict[str, str]) -> bool:
    """True if `element` tag matches step `name` under prefix mapping `nsmap`."""
    if ":" in name:
        prefix, local = name.split(":", 1)
        if prefix not in nsmap:
            raise UnsupportedXPathError(
                f"namespace prefix {prefix!r} is not defined in prefix mappings"
            )
        return element.tag == "{%s}%s" % (nsmap[prefix], local)
    return element.tag == name


def evaluate_xpath_subset(
    context: _Element,
    xpath: str,
    nsmap: Dict[str, str],
    ignore_final_predicate: bool = False,
) -> "List[_Element]":
    """Return elements of the custom XML tree selected by `xpath`.

    `context` is the store root element for an absolute `xpath` (leading '/') or
    the node a relative `xpath` is evaluated against (used for controls inside a
    repeating section item). `nsmap` maps namespace prefixes to URIs, typically
    parsed from w:dataBinding/@w:prefixMappings with parse_prefix_mappings().

    When `ignore_final_predicate` is True, a positional predicate on the last step
    is disregarded so all sibling matches are returned. This mirrors how a repeating
    section binding like '/root/items/item[1]' expands to every 'item'.
    """
    is_absolute, steps = parse_xpath_subset(xpath)
    if ignore_final_predicate and steps:
        last = steps[-1]
        steps[-1] = _Step(last.name, None)

    if is_absolute:
        first, remaining = steps[0], steps[1:]
        current = [context] if _name_matches(context, first.name, nsmap) else []
        if first.index not in (None, 1):
            current = []
    else:
        remaining = steps
        current = [context]

    for step in remaining:
        matched: List[_Element] = []
        for node in current:
            children = [c for c in node if _name_matches(c, step.name, nsmap)]
            if step.index is not None:
                children = children[step.index - 1 : step.index]
            matched.extend(children)
        current = matched
    return current


def element_text(element: _Element) -> str:
    """Text content of `element`, the concatenation of all descendant text."""
    return "".join(element.itertext())


class CustomXmlStore:
    """Read/write access to one custom XML part (/customXml/itemN.xml).

    Wraps the package |Part| holding the custom XML payload. Changes made through
    set_text() are written through to the part immediately, so they are included when
    the document is saved.
    """

    def __init__(self, part: Part, item_id: str | None):
        self._part = part
        self._item_id = item_id

    @property
    def item_id(self) -> str | None:
        """The ds:itemID of this store, e.g. '{EF278816-...}'.

        This is the value a w:dataBinding/@w:storeItemID attribute refers to. |None|
        when the store has no item-properties part to declare one.
        """
        return self._item_id

    @property
    def partname(self) -> str:
        """Partname of the underlying custom XML part, e.g. /customXml/item1.xml."""
        return str(self._part.partname)

    @lazyproperty
    def root(self) -> "_Element":
        """Root element of the custom XML payload."""
        return parse_xml(self._part.blob)

    def find(self, xpath: str, prefix_mappings: str | None = None) -> "List[_Element]":
        """Elements selected by `xpath`, evaluated with the supported XPath subset."""
        return evaluate_xpath_subset(self.root, xpath, parse_prefix_mappings(prefix_mappings))

    def find_text(self, xpath: str, prefix_mappings: str | None = None) -> str:
        """Text of the single element selected by `xpath`.

        Raises |UnsupportedXPathError| if `xpath` does not select exactly one
        element.
        """
        matches = self.find(xpath, prefix_mappings)
        if len(matches) != 1:
            raise UnsupportedXPathError(
                f"XPath {xpath!r} matched {len(matches)} elements, expected exactly 1"
            )
        return element_text(matches[0])

    def set_text(self, xpath: str, value: str, prefix_mappings: str | None = None) -> None:
        """Replace the text of the single element selected by `xpath` with `value`.

        Child elements of the target are discarded. The change is written through to
        the underlying part immediately. Raises |UnsupportedXPathError| if `xpath`
        does not select exactly one element.
        """
        matches = self.find(xpath, prefix_mappings)
        if len(matches) != 1:
            raise UnsupportedXPathError(
                f"XPath {xpath!r} matched {len(matches)} elements, expected exactly 1"
            )
        target = matches[0]
        for child in list(target):
            target.remove(child)
        target.text = value
        self._commit()

    def _commit(self) -> None:
        """Write the current XML payload back to the underlying part."""
        from docx.opc.oxml import serialize_part_xml

        self._part._blob = serialize_part_xml(self.root)


def custom_xml_stores(package) -> "List[CustomXmlStore]":
    """All |CustomXmlStore| objects for custom XML item parts in `package`."""
    from docx.opc.constants import CONTENT_TYPE as CT
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    stores = []
    for part in package.iter_parts():
        if not str(part.partname).startswith("/customXml/item"):
            continue
        if part.content_type != CT.XML:
            continue
        stores.append(CustomXmlStore(part, _item_id_for(part, RT.CUSTOM_XML_PROPS)))
    return stores


def _item_id_for(part: Part, props_reltype: str) -> "str | None":
    """The ds:itemID declared by the item-properties part related to `part`."""
    for rel in part.rels.values():
        if rel.is_external or rel.reltype != props_reltype:
            continue
        props_root = parse_xml(rel.target_part.blob)
        return props_root.get(qn("ds:itemID"))
    return None
