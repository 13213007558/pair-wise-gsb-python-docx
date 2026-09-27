"""Custom XML part objects.

A WordprocessingML package can contain one or more *custom XML* parts, each holding
an arbitrary XML data-store. Content controls (w:sdt) in the document can be bound to
items in these stores via a w:dataBinding element. Each custom XML item part
(/customXml/itemN.xml) is paired with a properties part (/customXml/itemPropsN.xml)
that carries the storeItemID used by bindings.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart
from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.package import Package


class CustomXmlPart(XmlPart):
    """A /customXml/itemN.xml part holding a custom XML data-store."""

    @classmethod
    def new(cls, package: Package, xml: str | bytes) -> CustomXmlPart:
        """Return a new custom XML item part containing `xml` as its data-store."""
        partname = package.next_partname("/customXml/item%d.xml")
        element = parse_xml(xml)
        return cls(partname, CT.XML, element, package)

    @property
    def store_item_id(self) -> str | None:
        """The storeItemID GUID identifying this data-store in bindings.

        The value is read from the related custom-XML-properties part and is |None|
        when no such related part is present.
        """
        props_part = self._props_part
        return None if props_part is None else props_part.item_id

    @property
    def _props_part(self) -> CustomXmlPropertiesPart | None:
        """The |CustomXmlPropertiesPart| related to this item part, if any."""
        for rel in self.rels.values():
            if rel.reltype == RT.CUSTOM_XML_PROPS and not rel.is_external:
                return rel.target_part  # type: ignore[return-value]
        return None


class CustomXmlPropertiesPart(XmlPart):
    """A /customXml/itemPropsN.xml part.

    Holds metadata for its companion custom XML item part, in particular the
    ds:itemID attribute that content-control data bindings refer to as
    w:storeItemID.
    """

    @classmethod
    def new(cls, package: Package, item_id: str) -> CustomXmlPropertiesPart:
        """Return a new custom-XML-properties part carrying `item_id`."""
        partname = package.next_partname("/customXml/itemProps%d.xml")
        element = parse_xml(
            '<ds:datastoreItem xmlns:ds="%s" ds:itemID="%s"/>' % (nsmap["ds"], item_id)
        )
        return cls(partname, CT.OFC_CUSTOM_XML_PROPERTIES, element, package)

    @property
    def item_id(self) -> str:
        """The ds:itemID value of the companion custom XML data-store."""
        return self._element.get(qn("ds:itemID"))  # type: ignore[return-value]
