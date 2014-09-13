
Charts
======

|pp| provides an API for adding and manipulating charts. A chart object, like
a table, is not a shape. Rather it is a graphical object contained in
a |GraphicFrame| shape. The shape API, such as position, size, shape id, and
name, are provided by the graphic frame shape. The chart itself is accessed
using the :attr:`chart` property on the graphic frame shape.


|Chart| objects
---------------

The |Chart| object is the root of a generally hierarchical graph of component
objects that together provide access to the properties and methods required
to specify and format a chart.

.. autoclass:: pptx.chart.chart.Chart
   :members:
   :member-order: bysource
   :undoc-members:
