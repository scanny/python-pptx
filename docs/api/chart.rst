
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


|Legend| objects
----------------

A legend provides a visual key relating each series of data points to their
assigned meaning by mapping a color, line type, or point shape to each series
name. A legend is optional, but there can be at most one. Most aspects of
a legend are determined automatically, but aspects of its position may be
specified via the API.

.. autoclass:: pptx.chart.chart.Legend()
   :members:
   :member-order: bysource
   :undoc-members:


|Axis| objects
--------------

A chart typically has two axes, a category axis and a value axis. In general,
one of these is horizontal and the other is vertical, where which is which
depends on the chart type. Perhaps the most prominent exception is a pie
chart, which has neither a category or value axis.

.. autoclass:: pptx.chart.axis._BaseAxis()
   :members:
   :member-order: bysource
   :undoc-members:


Value Axes
~~~~~~~~~~

Some axis properties are only relevant to value axes, in particular, those
related to numeric values rather than text category labels.

.. autoclass:: pptx.chart.axis.ValueAxis()
   :members:
   :member-order: bysource
   :undoc-members:


|TickLabels| objects
--------------------

Tick labels are the numbers appearing on a value axis or the category names
appearing on a category axis. Certain formatting options are available for
changing how these labels are displayed.

.. autoclass:: pptx.chart.axis.TickLabels()
   :members:
   :member-order: bysource
   :undoc-members:

