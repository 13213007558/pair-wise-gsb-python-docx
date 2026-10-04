.. _binding_api:

Content-control data binding objects
------------------------------------

Content controls (w:sdt elements) can be bound to nodes in a custom XML part
with a w:dataBinding element. The objects on this page allow the custom XML
data to be updated and the displayed content of bound controls to be refreshed
explicitly, without reopening the document in Word.

.. currentmodule:: docx.binding

.. autoclass:: DocumentBindings()
   :members:
   :exclude-members:
       part

.. autoclass:: RefreshPlan()
   :members:

.. autoclass:: PlanEntry()
   :members:

.. autoclass:: BindingProblem()
   :members:

.. currentmodule:: docx.customxml

.. autoclass:: CustomXmlStore()
   :members:
