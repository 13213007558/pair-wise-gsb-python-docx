.. _story_api:


Story objects
=============

Provides unified traversal of all document *stories*: the main document body
and each header and footer part. A story traversal visits every paragraph and
table, including content nested in table cells and inside DrawingML and VML
text boxes, which :attr:`~docx.document.Document.paragraphs` and
:attr:`~docx.document.Document.tables` do not reach.

Each yielded item records the package part that owns it, the chain of nested
containers from the story root, and -- for content inside a text box -- the
paragraph anchoring the text box. Floating text boxes appear at their XML
anchor position in XML anchor order, which is not Word's visual layout order.

Content in ``mc:AlternateContent`` is extracted exactly once: the first
``mc:Choice`` whose ``Requires`` capabilities are all in the caller-provided
capability set is used, otherwise ``mc:Fallback``.


|Stories| objects
-----------------

.. currentmodule:: docx.story

.. autoclass:: Stories
   :members:


|Story| objects
---------------

.. autoclass:: Story
   :members:

.. autoclass:: StoryReference
   :members:


|StoryItem| objects
-------------------

.. autoclass:: StoryItem
   :members:

.. autoclass:: StoryContainer
   :members:

.. autoclass:: TextBoxAnchor
   :members:


|TextBox| objects
-----------------

.. autoclass:: TextBox
   :members:
