.. _notes_api:

Footnote and endnote objects
=============================

.. currentmodule:: docx.notes

Footnotes and endnotes (collectively *notes*) are block-item containers that can
hold paragraphs, tables, rich text, hyperlinks, and images. Each note has an
internal *note-id* used to reference it from the document text. The note-id is
distinct from the number displayed to the reader, which Word assigns dynamically
based on the position of the note reference and the note-numbering settings.

Relationships for content added to a note, like image and hyperlink relationships,
belong to the part containing the note (the footnotes or endnotes part), not the
document part.


|Footnotes| objects
-------------------

.. autoclass:: Footnotes()
   :members:
   :inherited-members:


|Endnotes| objects
------------------

.. autoclass:: Endnotes()
   :members:
   :inherited-members:


|Footnote| objects
------------------

.. autoclass:: Footnote()
   :members:
   :inherited-members:
   :exclude-members:
       part


|Endnote| objects
-----------------

.. autoclass:: Endnote()
   :members:
   :inherited-members:
   :exclude-members:
       part


|NoteReference| objects
-----------------------

.. autoclass:: NoteReference()
   :members:
