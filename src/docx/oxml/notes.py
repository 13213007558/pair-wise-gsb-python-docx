"""Custom element classes related to footnotes and endnotes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Notes(BaseOxmlElement):
    """Base-class for `w:footnotes` and `w:endnotes` root elements.

    A notes root element contains a collection of note elements (`w:footnote` or
    `w:endnote`), each identified by a unique `w:id` attribute. Notes having a
    `w:type` attribute (like "separator" and "continuationSeparator") are special
    notes that do not hold note-content; these conventionally have non-positive ids
    (0 and -1) and are preserved as-is.
    """

    # -- subclasses provide these, based on their declared note tag --
    _note_tag: str
    _note_reference_style: str

    def add_note(self) -> CT_Note:
        """Return newly added note child element of this notes root element.

        The returned element is the minimum valid note, having a `w:id` value unique
        within the existing notes and content limited to a single paragraph containing
        the note-reference mark but no text. Content is added by adding runs to this
        first paragraph and by adding additional paragraphs or tables as needed.
        """
        next_id = self._next_available_note_id()
        note = cast(
            CT_Note,
            parse_xml(
                f'<{self._note_tag} {nsdecls("w")} w:id="{next_id}">'
                f"  <w:p>"
                f"    <w:pPr>"
                f'      <w:pStyle w:val="{self._note_text_style}"/>'
                f"    </w:pPr>"
                f"    <w:r>"
                f"      <w:rPr>"
                f'        <w:rStyle w:val="{self._note_reference_style}"/>'
                f"      </w:rPr>"
                f"      <w:noteRef/>"
                f"    </w:r>"
                f"  </w:p>"
                f"</{self._note_tag}>"
            ),
        )
        self.append(note)
        return note

    def get_note_by_id(self, note_id: int) -> CT_Note | None:
        """Return the note element identified by `note_id`, or |None| if not found."""
        note_elms = self.xpath(f"(./{self._note_tag}[@w:id='{note_id}'])[1]")
        return note_elms[0] if note_elms else None

    @property
    def note_lst(self) -> List[CT_Note]:
        """All note child elements of this notes root element."""
        return self.xpath(f"./{self._note_tag}")

    @property
    def _note_text_style(self) -> str:
        """The paragraph-style id applied to the first paragraph of a new note."""
        return "FootnoteText" if self._note_tag == "w:footnote" else "EndnoteText"

    def _next_available_note_id(self) -> int:
        """The next available note id.

        Special notes like the separator and continuation-separator carry
        non-positive ids (conventionally -1 and 0) so content notes are numbered
        from 1. The default mechanism is `max() + 1`; if that value would overflow
        a 32-bit signed integer, the first unused positive integer is used instead.
        """
        used_ids = [int(x) for x in self.xpath(f"./{self._note_tag}/@w:id")]

        # -- special notes carry non-positive ids; content notes are numbered from 1 --
        next_id = max(max(used_ids, default=0), 0) + 1

        if next_id <= 2**31 - 1:
            return next_id

        # -- fall-back to enumerating all used ids to find the first unused one --
        positive_ids = sorted(id_ for id_ in used_ids if id_ > 0)
        for expected, actual in enumerate(positive_ids, start=1):
            if expected != actual:
                return expected

        return len(positive_ids) + 1


class CT_Note(BaseOxmlElement):
    """Base-class for `w:footnote` and `w:endnote` elements.

    A footnote or endnote is a so-called "story" and can contain paragraphs and tables
    much like a table-cell, including rich text, hyperlinks, and images.

    The `w:id` attribute is the internal identifier used to reference this note from
    the document text. It is distinct from the number displayed to the reader, which
    Word assigns dynamically based on the position of the reference and the note
    numbering settings.
    """

    # -- attributes, common to both `w:footnote` and `w:endnote` --
    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:type", ST_String
    )

    # -- children --

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    # -- type-declarations for methods added by metaclass --

    add_p: Callable[[], CT_P]
    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this note."""
        return self.xpath("./w:p | ./w:tbl")


class CT_Footnotes(CT_Notes):
    """`w:footnotes` element, the root element for the footnotes part."""

    _note_tag = "w:footnote"
    _note_reference_style = "FootnoteReference"

    footnote = ZeroOrMore("w:footnote")


class CT_Footnote(CT_Note):
    """`w:footnote` element, representing a single footnote."""


class CT_Endnotes(CT_Notes):
    """`w:endnotes` element, the root element for the endnotes part."""

    _note_tag = "w:endnote"
    _note_reference_style = "EndnoteReference"

    endnote = ZeroOrMore("w:endnote")


class CT_Endnote(CT_Note):
    """`w:endnote` element, representing a single endnote."""
