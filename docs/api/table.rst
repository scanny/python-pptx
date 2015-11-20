.. _table_api:

Table-related objects
=====================


.. currentmodule:: pptx.shapes.table


|Table| objects
----------------

A |Table| object is added to a slide using the
:meth:`~.SlideShapeTree.add_table` method on |SlideShapeTree|.

.. autoclass:: Table()
   :members:
   :inherited-members:
   :exclude-members:
      element, has_textframe, notify_height_changed, notify_width_changed,
      part, text, textframe
   :undoc-members:


|_Column| objects
-----------------

.. autoclass:: _Column()
   :members:
   :member-order: bysource
   :undoc-members:


|_Row| objects
--------------

.. autoclass:: _Row()
   :members:
   :member-order: bysource
   :undoc-members:


|_Cell| objects
---------------

A |_Cell| object represents a single table cell at a particular row/column
location in the table. |_Cell| objects are not constructed directly. A
reference to a |_Cell| object is obtained using the :meth:`Table.cell` method,
specifying the cell's row/column location.

.. autoclass:: _Cell
   :members:
   :member-order: bysource
   :undoc-members:
