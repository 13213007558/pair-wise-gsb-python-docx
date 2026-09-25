# pyright: reportPrivateUsage=false

"""Footnote and endnote proxy objects.

A footnote or endnote (collectively "note") is a so-called "story" that can contain
paragraphs and tables, much like a table-cell. Each note is identified by a unique
note-id, which is distinct from the number displayed to the reader. The displayed
number is assigned dynamically by Word based on the position of the note reference in
the document and the note-numbering settings; it is not stored in the XML and is not
computed by this library.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from docx.blkcntnr import BlockItemContainer
from docx.oxml.ns import qn
from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.notes import CT_Note, CT_Notes
    from docx.parts.notes import _BaseNotesPart
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph


class Notes:
    """Base-class for |Footnotes| and |Endnotes| collections.

    Provides iteration, length, and id-based access to the content notes of the
    document. Special notes like the separator and continuation-separator are not
    included in iteration or the count; they are preserved unchanged in the XML.
    """

    _note_proxy_class: type[Note]

    def __init__(self, notes_elm: CT_Notes, notes_part: _BaseNotesPart):
        self._notes_elm = notes_elm
        self._notes_part = notes_part

    def __iter__(self) -> Iterator[Note]:
        """Iterator over the content notes in this collection.

        Special notes (those having a `w:type` attribute, like the separator and
        continuation-separator notes) are skipped.
        """
        return (
            self._note_proxy_class(note_elm, self._notes_part)
            for note_elm in self._notes_elm.note_lst
            if note_elm.type is None
        )

    def __len__(self) -> int:
        """The number of content notes in this collection.

        Special notes like the separator and continuation-separator are not counted.
        """
        return len([note_elm for note_elm in self._notes_elm.note_lst if note_elm.type is None])

    def get(self, note_id: int) -> Note | None:
        """Return the note identified by `note_id`, or |None| if not found.

        Note that `note_id` is the internal note identifier, not the number displayed
        to the reader, which Word assigns dynamically.
        """
        note_elm = self._notes_elm.get_note_by_id(note_id)
        return self._note_proxy_class(note_elm, self._notes_part) if note_elm is not None else None

    def remove(self, note: Note) -> None:
        """Remove `note` from this collection.

        Raises |ValueError| if `note` is still referenced from the document text. A
        note can only be removed once all its references have been removed, for
        example using `Note.references` and `NoteReference.remove()`, or
        `Note.delete()` which removes both. This ensures a note is only deleted when
        it is genuinely no longer referenced.
        """
        if note.references:
            raise ValueError(
                f"note with id {note.note_id} is still referenced from the document;"
                " remove its references first"
            )
        self._notes_elm.remove(note._note_elm)

    def _add_note(self, text: str = "") -> Note:
        """Return a new content note added to the end of this collection.

        The note is assigned a note-id that is unique within the existing notes,
        including any notes that were already present when the document was loaded.
        """
        note_elm = self._notes_elm.add_note()
        note = self._note_proxy_class(note_elm, self._notes_part)

        if text == "":
            return note

        para_text_iter = iter(text.split("\n"))

        first_para_text = next(para_text_iter)
        first_para = note.paragraphs[0]
        first_para.add_run(first_para_text)

        for s in para_text_iter:
            note.add_paragraph(text=s)

        return note


class Footnotes(Notes):
    """Collection providing access to the footnotes of this document."""

    def add_footnote(self, text: str = "") -> Footnote:
        """Return a new |Footnote| added to this collection.

        The footnote is assigned a unique footnote-id. If `text` is provided, it is
        added to the footnote; multiple paragraphs can be added by separating their
        text with newlines. A reference to the new footnote can be inserted at a run
        boundary using `Run.add_footnote_reference()`.
        """
        return self._add_note(text)  # type: ignore[return-value]


class Endnotes(Notes):
    """Collection providing access to the endnotes of this document."""

    def add_endnote(self, text: str = "") -> Endnote:
        """Return a new |Endnote| added to this collection.

        The endnote is assigned a unique endnote-id. If `text` is provided, it is
        added to the endnote; multiple paragraphs can be added by separating their
        text with newlines. A reference to the new endnote can be inserted at a run
        boundary using `Run.add_endnote_reference()`.
        """
        return self._add_note(text)  # type: ignore[return-value]


class Note(BlockItemContainer):
    """Base-class proxy for a single footnote or endnote.

    A note is a block-item container, similar to a table cell, so it can contain both
    paragraphs and tables and its paragraphs can contain rich text, hyperlinks, and
    images. Relationships for content added to a note (like images and hyperlinks)
    belong to the part containing the note, not the document part.
    """

    _default_paragraph_style: str

    def __init__(self, note_elm: CT_Note, notes_part: _BaseNotesPart):
        super().__init__(note_elm, notes_part)
        self._note_elm = note_elm
        self._notes_part = notes_part

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this note.

        The paragraph has `text` in a single run if present, and is given paragraph
        style `style`. When `style` is |None| or omitted, the default note paragraph
        style ("Footnote Text" or "Endnote Text") is applied.
        """
        paragraph = super().add_paragraph(text, style)

        # -- have to assign style directly to element because `paragraph.style` raises when
        # -- a style is not present in the styles part
        if style is None:
            paragraph._p.style = self._default_paragraph_style

        return paragraph

    def delete(self) -> None:
        """Remove this note along with all references to it from the document text."""
        for reference in self.references:
            reference.remove()
        self._note_elm.getparent().remove(self._note_elm)

    @property
    def note_id(self) -> int:
        """The internal identifier of this note.

        This value is distinct from the number displayed to the reader, which Word
        assigns dynamically based on the position of the note reference and the note
        numbering settings.
        """
        return self._note_elm.id

    @property
    def note_type(self) -> str | None:
        """The note-type attribute value of this note, like "separator".

        |None| for a regular content note. Special notes like the separator and
        continuation-separator notes have a note-type and do not appear in the
        `Footnotes` or `Endnotes` collection iteration.
        """
        return self._note_elm.type

    @property
    def references(self) -> list[NoteReference]:
        """A |NoteReference| object for each reference to this note in the document.

        A note is commonly referenced exactly once, from the position in the document
        text where the note mark appears. The same note can legitimately be referenced
        from more than one position, in which case the same displayed number appears
        at each referencing position.
        """
        return [
            NoteReference(reference_elm)
            for reference_elm in self._notes_part.note_references(self.note_id)
        ]

    @property
    def text(self) -> str:
        """The text content of this note as a string.

        Only content in paragraphs is included and of course all emphasis and styling
        is stripped. Paragraph boundaries are indicated with a newline.
        """
        return "\n".join(p.text for p in self.paragraphs)


class Footnote(Note):
    """Proxy for a single footnote in the document."""

    _default_paragraph_style = "FootnoteText"


class Endnote(Note):
    """Proxy for a single endnote in the document."""

    _default_paragraph_style = "EndnoteText"


class NoteReference(ElementProxy):
    """Proxy for a `w:footnoteReference` or `w:endnoteReference` element.

    A note reference occurs in a run in the document text and marks the position at
    which the note mark (the displayed note number) appears.
    """

    @property
    def note_id(self) -> int:
        """The id of the note this reference refers to."""
        return int(self._element.get(qn("w:id"), "0"))

    def remove(self) -> None:
        """Remove this reference from the document text.

        The containing run is also removed when it holds no other content, as is the
        case for a reference inserted by `Run.add_footnote_reference()` or
        `Run.add_endnote_reference()`. The note itself is not removed; use
        `Notes.remove()` or `Note.delete()` for that.
        """
        reference_elm = self._element
        run_elm = reference_elm.getparent()
        run_elm.remove(reference_elm)

        # -- remove the containing run too when it has no remaining content --
        if run_elm.tag == qn("w:r") and not any(child.tag != qn("w:rPr") for child in run_elm):
            run_elm.getparent().remove(run_elm)


Footnotes._note_proxy_class = Footnote
Endnotes._note_proxy_class = Endnote
