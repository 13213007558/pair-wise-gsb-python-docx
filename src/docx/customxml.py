"""API objects for working with custom XML data-stores.

A custom XML data-store is an arbitrary XML document stored in the package under
/customXml and identified by a storeItemID GUID. Content controls (w:sdt elements)
in the document, headers, footers, and tables can be bound to items in a store via
a w:dataBinding element. Updating the store item and refreshing the bound controls
together keeps the two views of the data consistent.
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Any, Dict, Iterator, List, Optional

from lxml import etree

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.customxml import CustomXmlPart, CustomXmlPropertiesPart

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage


def normalize_store_item_id(item_id: str) -> str:
    """Return canonical form of a storeItemID GUID for comparison.

    GUIDs compare case-insensitively and surrounding whitespace is insignificant.
    """
    return item_id.strip().lower()


class CustomXmlStore:
    """A single custom XML data-store, proxy for a |CustomXmlPart|."""

    def __init__(self, part: CustomXmlPart):
        self._part = part

    @property
    def element(self):
        """Root element of the custom XML data-store document."""
        return self._part.element

    @property
    def part(self) -> CustomXmlPart:
        """The |CustomXmlPart| this store is read from and written to."""
        return self._part

    @property
    def store_item_id(self) -> Optional[str]:
        """The storeItemID GUID identifying this store, e.g. '{9C1B...}'."""
        return self._part.store_item_id

    def find(self, xpath: str, namespaces: Optional[Dict[str, str]] = None) -> List[Any]:
        """Return the list of nodes in this store matched by `xpath`.

        `namespaces` maps the prefixes used in `xpath` to namespace URIs, as parsed
        from the w:prefixMappings attribute of a data binding. Matched nodes are
        elements, or attribute "smart strings" when `xpath` selects an attribute.
        """
        return etree._Element.xpath(self.element, xpath, namespaces=namespaces or {})


class CustomXmlStores:
    """Collection of the custom XML data-stores in a package."""

    def __init__(self, package: OpcPackage):
        self._package = package

    def __iter__(self) -> Iterator[CustomXmlStore]:
        for rel in self._package.rels.values():
            if rel.is_external or rel.reltype != RT.CUSTOM_XML:
                continue
            yield CustomXmlStore(rel.target_part)  # type: ignore[arg-type]

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __getitem__(self, idx: int) -> CustomXmlStore:
        return list(self)[idx]

    def add(self, xml: str | bytes) -> CustomXmlStore:
        """Return a new store containing `xml`, added to the package.

        A fresh storeItemID GUID is generated and recorded in a new
        custom-XML-properties part related to the new item part.
        """
        package = self._package
        item_id = "{%s}" % str(uuid.uuid4()).upper()
        props_part = CustomXmlPropertiesPart.new(package, item_id)
        item_part = CustomXmlPart.new(package, xml)
        item_part.relate_to(props_part, RT.CUSTOM_XML_PROPS)
        package.relate_to(item_part, RT.CUSTOM_XML)
        return CustomXmlStore(item_part)

    def get_by_item_id(self, store_item_id: str) -> Optional[CustomXmlStore]:
        """Return the store identified by `store_item_id`, or |None|.

        Comparison is case-insensitive, as GUIDs are.
        """
        wanted = normalize_store_item_id(store_item_id)
        for store in self:
            if store.store_item_id is None:
                continue
            if normalize_store_item_id(store.store_item_id) == wanted:
                return store
        return None
