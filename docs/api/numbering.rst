.. _numbering_api:

List numbering objects
======================

Objects related to resolving the rendered numbers of list paragraphs.

|Document| provides :meth:`.Document.iter_list_items`, which generates a
|ListItem| for each numbered or bulleted paragraph in document order,
including paragraphs inside tables. Each item reports the concrete numbering
instance (``num_id``), the zero-based ``level``, the computed counter
``value``, and the fully expanded ``label`` (for example ``"2.3)"`` or
``"第十二条"``).

Resolution takes into account ``abstractNum``/``num`` definitions,
``lvlOverride``/``startOverride``, ``lvlRestart``, style links (both a
paragraph style carrying ``w:numPr`` and a level linked to a style via
``w:lvl/w:pStyle``), and upper-level references (like ``%1``) in
``w:lvlText``. Counters are tracked independently for each distinct
``num_id``. The traversal is strictly read-only; automatic numbering is never
converted to literal paragraph text.

The number formats that can be rendered are listed in
:data:`~docx.numbering.SUPPORTED_NUM_FMTS`. Items using any other format keep
their original ``num_fmt`` definition and are reported via
:attr:`~ListItem.is_num_fmt_supported`; their ``label`` is |None| rather than
being coerced to decimal.


|ListItem| objects
------------------

.. autoclass:: docx.numbering.ListItem()
   :members:
   :exclude-members: __init__


Module constants
----------------

.. autodata:: docx.numbering.SUPPORTED_NUM_FMTS
