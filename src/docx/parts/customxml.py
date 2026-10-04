"""|CustomXmlPart| and closely related objects."""

from __future__ import annotations

from docx.opc.part import XmlPart


class CustomXmlPart(XmlPart):
    """A custom XML part, e.g. '/customXml/item1.xml'.

    Holds the XML data-store content that content controls ('w:sdt' elements) bind
    to via a 'w:dataBinding' element. Changes made to '.element' are serialized
    back into the package when the document is saved.
    """
