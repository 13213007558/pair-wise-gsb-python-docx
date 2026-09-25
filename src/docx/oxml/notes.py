"""Custom element classes related to footnotes and endnotes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P

# -- `w:type` values reserved for notes that do not appear in the note stream --
SPECIAL_NOTE_TYPES = ("separator", "continuationSeparator", "continuationNotice")


class _BaseNote(BaseOxmlElement):
    """Shared behavior for `w:footnote` and `w:endnote` elements.

    A footnote or endnote is a so-called "story" and can contain paragraphs and
    tables much like a table-cell, including rich text, hyperlinks, and images.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:type", ST_String
    )

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    add_p: Callable[[], CT_P]
    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """All `w:p` and `w:tbl` children of this note, in document order."""
        return self.xpath("./w:p | ./w:tbl")

    @property
    def is_special(self) -> bool:
        """True when this note is a separator, continuation-separator, or continuation-notice.

        Such notes are not part of the user-visible note stream; they define the
        layout furniture (like the separator line) and are never referenced from
        the document text.
        """
        return self.type in SPECIAL_NOTE_TYPES


class _BaseNotes(BaseOxmlElement):
    """Shared behavior for `w:footnotes` and `w:endnotes` root elements."""

    _note_tag: str

    def _next_available_note_id(self) -> int:
        """The next available note id.

        Ids `-1` and `0` are reserved for the separator and
        continuation-separator notes, so the smallest assignable id is 1. The
        default mechanism is `max() + 1`, which never reuses an id already
        present in a loaded file. If that would overflow a 32-bit signed integer,
        fall back to the first unused positive integer.
        """
        used_ids = [int(x) for x in self.xpath("./%s/@w:id" % self._note_tag)]

        next_id = max(used_ids, default=0) + 1
        if next_id < 1:
            next_id = 1

        if next_id <= 2**31 - 1:
            return next_id

        positive_ids = sorted(id_ for id_ in used_ids if id_ >= 1)
        for expected, actual in enumerate(positive_ids, start=1):
            if expected != actual:
                return expected

        return len(positive_ids) + 1


class CT_Footnote(_BaseNote):
    """
    `w:footnote` element, representing a single footnote.
    """


class CT_Footnotes(_BaseNotes):
    """
    `w:footnotes` element, the root element of the footnotes part.

    Contains a `w:footnote` element for each footnote, including the special
    separator and continuation-separator footnotes (ids -1 and 0 by convention).
    Each footnote is identified by a unique `w:id` attribute, used to reference the
    footnote from the document text with a `w:footnoteReference` element.
    """

    _note_tag = "w:footnote"

    # -- type-declarations to fill in the gaps for metaclass-added methods --
    footnote_lst: list[CT_Footnote]

    footnote = ZeroOrMore("w:footnote")

    def add_footnote(self) -> CT_Footnote:
        """Return newly added `w:footnote` child of this `w:footnotes`.

        The returned `w:footnote` element is the minimum valid value, having a
        `w:id` value unique within the existing footnotes. Its content is a single
        paragraph containing the footnote-reference mark but no text. Content is
        added by adding runs to this first paragraph and by adding additional
        paragraphs or tables as needed.
        """
        next_id = self._next_available_note_id()
        footnote = cast(
            CT_Footnote,
            parse_xml(
                '<w:footnote %s w:id="%d">'
                "  <w:p>"
                "    <w:pPr>"
                '      <w:pStyle w:val="FootnoteText"/>'
                "    </w:pPr>"
                "    <w:r>"
                "      <w:rPr>"
                '        <w:rStyle w:val="FootnoteReference"/>'
                "      </w:rPr>"
                "      <w:footnoteRef/>"
                "    </w:r>"
                "  </w:p>"
                "</w:footnote>" % (nsdecls("w"), next_id)
            ),
        )
        self.append(footnote)
        return footnote

    def get_footnote_by_id(self, footnote_id: int) -> CT_Footnote | None:
        """Return the `w:footnote` element identified by `footnote_id`.

        Returns |None| if no footnote has that id, including when `footnote_id`
        identifies a special footnote such as the separator.
        """
        footnote_elms = self.xpath("(./w:footnote[@w:id='%d'])[1]" % footnote_id)
        return footnote_elms[0] if footnote_elms else None


class CT_Endnote(_BaseNote):
    """
    `w:endnote` element, representing a single endnote.
    """


class CT_Endnotes(_BaseNotes):
    """
    `w:endnotes` element, the root element of the endnotes part.

    Mirrors `CT_Footnotes` for endnotes, each referenced from the document text
    with a `w:endnoteReference` element.
    """

    _note_tag = "w:endnote"

    # -- type-declarations to fill in the gaps for metaclass-added methods --
    endnote_lst: list[CT_Endnote]

    endnote = ZeroOrMore("w:endnote")

    def add_endnote(self) -> CT_Endnote:
        """Return newly added `w:endnote` child of this `w:endnotes`.

        Mirrors `CT_Footnotes.add_footnote()` for endnotes.
        """
        next_id = self._next_available_note_id()
        endnote = cast(
            CT_Endnote,
            parse_xml(
                '<w:endnote %s w:id="%d">'
                "  <w:p>"
                "    <w:pPr>"
                '      <w:pStyle w:val="EndnoteText"/>'
                "    </w:pPr>"
                "    <w:r>"
                "      <w:rPr>"
                '        <w:rStyle w:val="EndnoteReference"/>'
                "      </w:rPr>"
                "      <w:endnoteRef/>"
                "    </w:r>"
                "  </w:p>"
                "</w:endnote>" % (nsdecls("w"), next_id)
            ),
        )
        self.append(endnote)
        return endnote

    def get_endnote_by_id(self, endnote_id: int) -> CT_Endnote | None:
        """Return the `w:endnote` element identified by `endnote_id`.

        Returns |None| if no endnote has that id, including when `endnote_id`
        identifies a special endnote such as the separator.
        """
        endnote_elms = self.xpath("(./w:endnote[@w:id='%d'])[1]" % endnote_id)
        return endnote_elms[0] if endnote_elms else None
