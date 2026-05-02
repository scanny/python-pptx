.. _table_api:

Table-related objects
=====================


.. currentmodule:: pptx.table


|Table| objects
----------------

A |Table| object is added to a slide using the
:meth:`~.SlideShapes.add_table` method on |SlideShapes|.

.. autoclass:: Table()
   :members:
   :inherited-members:
   :exclude-members:
      notify_height_changed, notify_width_changed, part
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


|_RowCollection| objects
------------------------

A |_RowCollection| object is accessed using the :attr:`.Table.rows` property.
Rows can be iterated, indexed, and added to the bottom of the table using its
:meth:`~._RowCollection.add` method.

.. autoclass:: _RowCollection()
   :members: add
   :member-order: bysource
   :undoc-members:


|_Cell| objects
---------------

A |_Cell| object represents a single table cell at a particular row/column
location in the table. |_Cell| objects are not constructed directly. A
reference to a |_Cell| object is obtained using the :meth:`Table.cell` method,
specifying the cell's row/column location. A cell object can also be obtained
using the :attr:`_Row.cells` collection.

.. autoclass:: _Cell
   :members:
   :member-order: bysource
   :undoc-members:
