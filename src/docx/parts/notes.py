"""Contains footnotes and endnotes of the document."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.notes import Endnotes, Footnotes
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.parser import parse_xml
from docx.package import Package
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.oxml.notes import CT_Endnotes, CT_Footnotes
    from docx.package import Package


class FootnotesPart(StoryPart):
    """Container part for the footnotes of the document."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_Footnotes, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._footnotes = element

    @property
    def footnotes(self) -> Footnotes:
        """A |Footnotes| proxy object for the `w:footnotes` root element of this part."""
        return Footnotes(self._footnotes, self)

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created footnotes part.

        Contains only the separator and continuation-separator footnotes, the same
        starting state Word uses.
        """
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = cast("CT_Footnotes", parse_xml(cls._default_footnotes_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_footnotes_xml(cls) -> bytes:
        """A byte-string containing XML for a default footnotes part."""
        path = os.path.join(
            os.path.split(__file__)[0], "..", "templates", "default-footnotes.xml"
        )
        with open(path, "rb") as f:
            xml_bytes = f.read()
        return xml_bytes


class EndnotesPart(StoryPart):
    """Container part for the endnotes of the document."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_Endnotes, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._endnotes = element

    @property
    def endnotes(self) -> Endnotes:
        """An |Endnotes| proxy object for the `w:endnotes` root element of this part."""
        return Endnotes(self._endnotes, self)

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created endnotes part.

        Contains only the separator and continuation-separator endnotes, the same
        starting state Word uses.
        """
        partname = PackURI("/word/endnotes.xml")
        content_type = CT.WML_ENDNOTES
        element = cast("CT_Endnotes", parse_xml(cls._default_endnotes_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_endnotes_xml(cls) -> bytes:
        """A byte-string containing XML for a default endnotes part."""
        path = os.path.join(
            os.path.split(__file__)[0], "..", "templates", "default-endnotes.xml"
        )
        with open(path, "rb") as f:
            xml_bytes = f.read()
        return xml_bytes
