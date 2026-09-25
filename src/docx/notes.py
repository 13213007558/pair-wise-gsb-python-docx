# pyright: reportPrivateUsage=false

"""Footnote and endnote related proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from docx.blkcntnr import BlockItemContainer
from docx.oxml.ns import qn
from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.notes import CT_Endnote, CT_Endnotes, CT_Footnote, CT_Footnotes
    from docx.oxml.text.run import CT_R
    from docx.parts.notes import EndnotesPart, FootnotesPart
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


class _Notes:
    """Shared behavior for |Footnotes| and |Endnotes| collections."""

    # -- local tag-names, e.g. "w:footnote" and "w:footnoteReference" --
    _note_tag: str
    _reference_tag: str

    def __init__(self, notes_elm: CT_Footnotes | CT_Endnotes, notes_part):
        self._notes_elm = notes_elm
        self._notes_part = notes_part

    def __len__(self) -> int:
        """The number of notes in this collection.

        Special notes like the separator and continuation-separator are not
        included in this count.
        """
        return len(self._normal_note_elms)

    def _populate_text(self, note: _Note, text: str) -> None:
        """Add paragraphs to `note` from `text`, split on newlines."""
        para_text_iter = iter(text.split("\n"))

        first_para_text = next(para_text_iter)
        first_para = note.paragraphs[0]
        first_para.add_run(first_para_text)

        for s in para_text_iter:
            note.add_paragraph(text=s)

    def _prune(self) -> list[int]:
        """Remove notes that are no longer referenced from the document.

        Special notes (separator, continuation-separator, continuation-notice)
        are always kept. Returns the ids of the notes that were removed.
        """
        referenced_ids = self._referenced_note_ids
        removed_ids: list[int] = []
        for note_elm in self._normal_note_elms:
            if note_elm.id not in referenced_ids:
                self._notes_elm.remove(note_elm)
                removed_ids.append(note_elm.id)
        return removed_ids

    @property
    def _normal_note_elms(self) -> list[CT_Footnote | CT_Endnote]:
        """The note elements that are part of the user-visible note stream.

        Excludes special notes like the separator and continuation-separator.
        """
        return self._notes_elm.xpath(
            "./%s[not(@w:type) or @w:type='normal']" % self._note_tag
        )

    @property
    def _referenced_note_ids(self) -> list[int]:
        """Ids of all notes referenced from the document text, in document order."""
        document_elm = self._notes_part._document_part.element
        return [
            int(id_str) for id_str in document_elm.xpath(".//%s/@w:id" % self._reference_tag)
        ]


class Footnotes(_Notes):
    """Collection of the footnotes in this document."""

    _note_tag = "w:footnote"
    _reference_tag = "w:footnoteReference"

    def __init__(self, footnotes_elm: CT_Footnotes, footnotes_part: FootnotesPart):
        super().__init__(footnotes_elm, footnotes_part)

    def __iter__(self) -> Iterator[Footnote]:
        """Iterator over the footnotes in this collection.

        Special footnotes like the separator and continuation-separator are not
        included in this iteration.
        """
        return (Footnote(note_elm, self._notes_part) for note_elm in self._normal_note_elms)

    def add_footnote(self, text: str = "") -> Footnote:
        """Return a new |Footnote| added to the end of this collection.

        The footnote is assigned a unique footnote-id that does not collide with
        any id already present, including ids in a loaded document.

        If `text` is provided, it is added to the footnote, with newlines
        (`"\\n"`) starting a new paragraph. The default is a single empty
        paragraph, to which runs can be added with
        `footnote.paragraphs[0].add_run()`.

        Note the footnote does not appear in the rendered document until a
        reference to it is inserted in the document text; use
        `Document.add_footnote()` to create the footnote and its reference in
        one step, or `Footnote.add_reference()` to reference an existing
        footnote.
        """
        footnote_elm = self._notes_elm.add_footnote()
        footnote = Footnote(footnote_elm, self._notes_part)

        if text != "":
            self._populate_text(footnote, text)

        return footnote

    def get(self, footnote_id: int) -> Footnote | None:
        """Return the footnote identified by `footnote_id`, or |None| if not found.

        Special footnotes like the separator are not returned, even though their
        ids (-1 and 0 by convention) are present in the XML.
        """
        footnote_elm = self._notes_elm.get_footnote_by_id(footnote_id)
        if footnote_elm is None or footnote_elm.is_special:
            return None
        return Footnote(footnote_elm, self._notes_part)

    def prune(self) -> list[int]:
        """Remove footnotes that are no longer referenced from the document.

        Special footnotes (separator, continuation-separator, and
        continuation-notice) are always preserved. Returns the ids of the
        footnotes that were removed.
        """
        return self._prune()


class Endnotes(_Notes):
    """Collection of the endnotes in this document."""

    _note_tag = "w:endnote"
    _reference_tag = "w:endnoteReference"

    def __init__(self, endnotes_elm: CT_Endnotes, endnotes_part: EndnotesPart):
        super().__init__(endnotes_elm, endnotes_part)

    def __iter__(self) -> Iterator[Endnote]:
        """Iterator over the endnotes in this collection.

        Special endnotes like the separator and continuation-separator are not
        included in this iteration.
        """
        return (Endnote(note_elm, self._notes_part) for note_elm in self._normal_note_elms)

    def add_endnote(self, text: str = "") -> Endnote:
        """Return a new |Endnote| added to the end of this collection.

        Mirrors |Footnotes.add_footnote| for endnotes.
        """
        endnote_elm = self._notes_elm.add_endnote()
        endnote = Endnote(endnote_elm, self._notes_part)

        if text != "":
            self._populate_text(endnote, text)

        return endnote

    def get(self, endnote_id: int) -> Endnote | None:
        """Return the endnote identified by `endnote_id`, or |None| if not found.

        Special endnotes like the separator are not returned, even though their
        ids (-1 and 0 by convention) are present in the XML.
        """
        endnote_elm = self._notes_elm.get_endnote_by_id(endnote_id)
        if endnote_elm is None or endnote_elm.is_special:
            return None
        return Endnote(endnote_elm, self._notes_part)

    def prune(self) -> list[int]:
        """Remove endnotes that are no longer referenced from the document.

        Special endnotes (separator, continuation-separator, and
        continuation-notice) are always preserved. Returns the ids of the
        endnotes that were removed.
        """
        return self._prune()


class _Note(BlockItemContainer):
    """Shared behavior for |Footnote| and |Endnote| proxy objects.

    A note is a block-item container, similar to a table cell, so it can contain
    both paragraphs and tables and its paragraphs can contain rich text,
    hyperlinks, and images. Relationships for images and hyperlinks added to a
    note belong to the part containing the note (the footnotes or endnotes
    part), not to the main document part.
    """

    # -- local tag-names and default style-ids, differ for footnote/endnote --
    _note_tag: str
    _reference_tag: str
    _text_style: str

    def __init__(self, note_elm: CT_Footnote | CT_Endnote, notes_part):
        super().__init__(note_elm, notes_part)
        self._note_elm = note_elm
        self._notes_part = notes_part

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this note.

        The paragraph has `text` in a single run if present, and is given
        paragraph style `style`. When `style` is |None| or omitted, the default
        note style ("FootnoteText" or "EndnoteText") is applied.
        """
        paragraph = super().add_paragraph(text, style)

        # -- have to assign style directly to element because `paragraph.style` raises
        # -- when a style is not present in the styles part
        if style is None:
            paragraph._p.style = self._text_style

        return paragraph

    @property
    def is_special(self) -> bool:
        """True when this note is a separator, continuation-separator, or continuation-notice.

        Special notes define layout furniture rather than note content and are
        never referenced from the document text.
        """
        return self._note_elm.is_special

    @property
    def number(self) -> int | None:
        """The number displayed for this note in the rendered document.

        This is distinct from the note id: the id is an arbitrary unique
        identifier stored in the XML, while the number is computed from the
        position of the note's first reference among all note references in the
        document, which is how Word numbers notes under the default continuous
        numbering. Returns |None| when the note is not referenced from the
        document text.
        """
        note_id = self._note_elm.id
        seen: list[int] = []
        normal_ids = {note_elm.id for note_elm in self._normal_note_elms}
        for ref_id in self._referenced_note_ids:
            if ref_id in normal_ids and ref_id not in seen:
                seen.append(ref_id)
        return seen.index(note_id) + 1 if note_id in seen else None

    @property
    def text(self) -> str:
        """The text content of this note as a string.

        Only content in paragraphs is included and of course all emphasis and
        styling is stripped. Paragraph boundaries are indicated with a newline
        (`"\\n"`).
        """
        return "\n".join(p.text for p in self.paragraphs)

    def delete(self) -> None:
        """Remove this note and all references to it from the document."""
        for reference in self._references:
            reference._remove_run()
        self._note_elm.getparent().remove(self._note_elm)

    @property
    def _references(self) -> list[_NoteReference]:
        """All references to this note in the document text, in document order."""
        document_elm = self._notes_part._document_part.element
        return [
            self._reference_cls(r_elm, self._notes_part)
            for r_elm in document_elm.xpath(
                ".//w:r[%s[@w:id='%d']]" % (self._reference_tag, self._note_elm.id)
            )
        ]

    @property
    def _reference_cls(self) -> type[_NoteReference]:
        raise NotImplementedError

    @property
    def _normal_note_elms(self) -> list[CT_Footnote | CT_Endnote]:
        notes_elm = self._note_elm.getparent()
        return notes_elm.xpath("./%s[not(@w:type) or @w:type='normal']" % self._note_tag)

    @property
    def _referenced_note_ids(self) -> list[int]:
        document_elm = self._notes_part._document_part.element
        return [
            int(id_str) for id_str in document_elm.xpath(".//%s/@w:id" % self._reference_tag)
        ]


class Footnote(_Note):
    """Proxy for a single footnote in the document."""

    _note_tag = "w:footnote"
    _reference_tag = "w:footnoteReference"
    _text_style = "FootnoteText"

    def __init__(self, footnote_elm: CT_Footnote, footnotes_part: FootnotesPart):
        super().__init__(footnote_elm, footnotes_part)

    @property
    def footnote_id(self) -> int:
        """The unique identifier of this footnote.

        This is the value of the `w:id` attribute in the XML. It is an
        arbitrary unique identifier, not the number displayed for the footnote
        in the rendered document; see `Footnote.number` for that.
        """
        return self._note_elm.id

    @property
    def references(self) -> list[FootnoteReference]:
        """All references to this footnote in the document text, in document order.

        A footnote can be referenced more than once, in which case each
        reference displays the same footnote number.
        """
        return [r for r in self._references if isinstance(r, FootnoteReference)]

    def add_reference(self, run: Run) -> FootnoteReference:
        """Insert a reference to this footnote immediately after `run`.

        Use this to reference the same footnote from more than one place in the
        document text.
        """
        reference_run = run._r.insert_footnote_reference_below(self.footnote_id)
        return FootnoteReference(reference_run, self._notes_part)

    @property
    def _reference_cls(self) -> type[FootnoteReference]:
        return FootnoteReference


class Endnote(_Note):
    """Proxy for a single endnote in the document."""

    _note_tag = "w:endnote"
    _reference_tag = "w:endnoteReference"
    _text_style = "EndnoteText"

    def __init__(self, endnote_elm: CT_Endnote, endnotes_part: EndnotesPart):
        super().__init__(endnote_elm, endnotes_part)

    @property
    def endnote_id(self) -> int:
        """The unique identifier of this endnote.

        This is the value of the `w:id` attribute in the XML. It is an
        arbitrary unique identifier, not the number displayed for the endnote
        in the rendered document; see `Endnote.number` for that.
        """
        return self._note_elm.id

    @property
    def references(self) -> list[EndnoteReference]:
        """All references to this endnote in the document text, in document order."""
        return [r for r in self._references if isinstance(r, EndnoteReference)]

    def add_reference(self, run: Run) -> EndnoteReference:
        """Insert a reference to this endnote immediately after `run`.

        Use this to reference the same endnote from more than one place in the
        document text.
        """
        reference_run = run._r.insert_endnote_reference_below(self.endnote_id)
        return EndnoteReference(reference_run, self._notes_part)

    @property
    def _reference_cls(self) -> type[EndnoteReference]:
        return EndnoteReference


class _NoteReference(ElementProxy):
    """Shared behavior for |FootnoteReference| and |EndnoteReference|.

    A note reference is the run in the document text that contains the
    reference mark (`w:footnoteReference` or `w:endnoteReference`) pointing
    to a note.
    """

    _reference_tag: str

    def __init__(self, r_elm: CT_R, notes_part):
        super().__init__(r_elm, notes_part)
        self._r = r_elm
        self._notes_part = notes_part

    def delete(self) -> None:
        """Remove this reference from the document.

        The referenced note is also removed when this was its last remaining
        reference, since a note that is no longer referenced is dead content.
        Special notes (separator and the like) are never removed.
        """
        note_id = self._note_id
        notes = self._notes
        self._remove_run()
        if note_id not in notes._referenced_note_ids:
            note_elm = notes._notes_elm.xpath(
                "./%s[@w:id='%d']" % (notes._note_tag, note_id)
            )
            if note_elm and not note_elm[0].is_special:
                notes._notes_elm.remove(note_elm[0])

    @property
    def _note_id(self) -> int:
        """The id of the note this reference points to."""
        reference_elm = self._r.find(qn(self._reference_tag))
        return int(reference_elm.get(qn("w:id")))

    def _remove_run(self) -> None:
        """Remove the run containing this reference from the document."""
        self._r.getparent().remove(self._r)

    @property
    def _notes(self) -> _Notes:
        raise NotImplementedError


class FootnoteReference(_NoteReference):
    """Proxy for a footnote-reference mark in the document text."""

    _reference_tag = "w:footnoteReference"

    @property
    def footnote_id(self) -> int:
        """The id of the footnote this reference points to."""
        return self._note_id

    @property
    def _notes(self) -> Footnotes:
        return self._notes_part.footnotes


class EndnoteReference(_NoteReference):
    """Proxy for an endnote-reference mark in the document text."""

    _reference_tag = "w:endnoteReference"

    @property
    def endnote_id(self) -> int:
        """The id of the endnote this reference points to."""
        return self._note_id

    @property
    def _notes(self) -> Endnotes:
        return self._notes_part.endnotes
