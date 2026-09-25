"""Contains footnotes and endnotes parts."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.parser import parse_xml
from docx.package import Package
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.notes import Endnotes, Footnotes
    from docx.oxml.notes import CT_Endnotes, CT_Footnotes, CT_Note, CT_Notes
    from docx.package import Package


class _BaseNotesPart(StoryPart):
    """Base-class for footnotes and endnotes parts.

    A notes part is a story part; its notes can contain rich content including
    paragraphs, tables, hyperlinks, and images. Relationships for such content (like
    image and hyperlink relationships) are added to this part, not the document part.
    """

    _default_partname: str
    _default_content_type: str
    _default_template_filename: str
    _reference_tag: str

    def __init__(self, partname: PackURI, content_type: str, element: CT_Notes, package: Package):
        super().__init__(partname, content_type, element, package)
        self._notes = element

    def note_references(self, note_id: int) -> list[CT_Note]:
        """All reference elements in the document part that refer to `note_id`.

        A reference element is a `w:footnoteReference` or `w:endnoteReference`
        element, according to the note-type contained in this part.
        """
        document_element = self._document_part.element
        return document_element.xpath(f".//{self._reference_tag}[@w:id='{note_id}']")

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created notes part, containing the default separator notes."""
        element = cast("CT_Notes", parse_xml(cls._default_notes_xml()))
        return cls(PackURI(cls._default_partname), cls._default_content_type, element, package)

    @classmethod
    def _default_notes_xml(cls) -> bytes:
        """A byte-string containing XML for a default notes part."""
        path = os.path.join(
            os.path.split(__file__)[0], "..", "templates", cls._default_template_filename
        )
        with open(path, "rb") as f:
            xml_bytes = f.read()
        return xml_bytes


class FootnotesPart(_BaseNotesPart):
    """Container part for the footnotes of the document."""

    _default_partname = "/word/footnotes.xml"
    _default_content_type = CT.WML_FOOTNOTES
    _default_template_filename = "default-footnotes.xml"
    _reference_tag = "w:footnoteReference"

    @property
    def footnotes(self) -> Footnotes:
        """A |Footnotes| proxy object for the `w:footnotes` root element of this part."""
        from docx.notes import Footnotes

        return Footnotes(cast("CT_Footnotes", self._notes), self)


class EndnotesPart(_BaseNotesPart):
    """Container part for the endnotes of the document."""

    _default_partname = "/word/endnotes.xml"
    _default_content_type = CT.WML_ENDNOTES
    _default_template_filename = "default-endnotes.xml"
    _reference_tag = "w:endnoteReference"

    @property
    def endnotes(self) -> Endnotes:
        """An |Endnotes| proxy object for the `w:endnotes` root element of this part."""
        from docx.notes import Endnotes

        return Endnotes(cast("CT_Endnotes", self._notes), self)
