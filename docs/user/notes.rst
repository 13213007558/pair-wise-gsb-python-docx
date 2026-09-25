.. _notes:

Working with Footnotes and Endnotes
===================================

Word documents can contain *footnotes*, rendered at the bottom of the page, and
*endnotes*, rendered at the end of the document or section. Both are referred to
here as *notes*; they share the same structure and the same API, so everything
below about footnotes applies equally to endnotes.

**Note anatomy.** Each note has two parts, the *note-reference* and the
*note-content*:

The **note-reference** is the small superscript mark in the document text. In the
XML it is a run containing a ``<w:footnoteReference>`` (or ``<w:endnoteReference>``)
element. A note can be referenced from more than one place; each reference then
displays the same number.

The **note-content** is stored in a separate part (``word/footnotes.xml`` or
``word/endnotes.xml``), not in the main document part. A note is a so-called
*story* and can contain paragraphs and tables, and its paragraphs can contain
rich text, hyperlinks, and images. Relationships for images and hyperlinks
inside a note belong to the part containing the note, never to the main
document part.

**Note id vs. number.** Each note is identified in the XML by a unique ``w:id``
attribute, available as ``Footnote.footnote_id``. This id is an arbitrary
identifier and is distinct from the number displayed for the note in the
rendered document, which Word computes from the position of the note's first
reference. The displayed number is available as ``Footnote.number``. The ids
``-1`` and ``0`` are reserved for the special *separator* and
*continuation-separator* notes, which define layout furniture (like the line
above the footnote area). These special notes are preserved when present but
are not included in the |Footnotes| collection.

Adding a footnote
-----------------

A footnote is anchored at a run boundary. Pass the run after which the
reference mark should appear::

    >>> document = Document()
    >>> paragraph = document.add_paragraph("Some claim.")
    >>> footnote = document.add_footnote(paragraph.runs[0], text="The evidence.")
    >>> footnote.footnote_id
    1
    >>> footnote.number
    1

Rich content can be added to the note using its container API::

    >>> footnote.paragraphs[0].add_run().add_picture("evidence.png")
    >>> paragraph2 = footnote.add_paragraph("A second paragraph.")

The same note can be referenced again elsewhere::

    >>> other = document.add_paragraph("Same claim again.")
    >>> footnote.add_reference(other.runs[0])
    >>> len(footnote.references)
    2

Endnotes work the same way, using ``document.add_endnote()`` and
``document.endnotes``.

Removing notes and references
-----------------------------

A single reference can be removed with ``FootnoteReference.delete()``. When the
deleted reference was the last one pointing to a note, the note itself is
removed as well, since an unreferenced note is dead content::

    >>> footnote.references[0].delete()

A note and all its references can be removed in one step with
``Footnote.delete()``. Notes that have become unreferenced by other means (for
example after deleting text) can be cleaned up with ``Footnotes.prune()``,
which always preserves the special separator notes.
