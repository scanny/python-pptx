.. _chart-api:

Charts
======

|pp| provides an API for adding and manipulating charts. A chart object, like
a table, is not a shape. Rather it is a graphical object contained in
a |GraphicFrame| shape. The shape API, such as position, size, shape id, and
name, are provided by the graphic frame shape. The chart itself is accessed
using the :attr:`chart` property on the graphic frame shape.


|ChartData| objects
-------------------

A |ChartData| object is used to specify the categories and series that
together comprise the data depicted in a chart. It is used when creating
a new chart and when replacing the data for an existing chart.

.. autoclass:: pptx.chart.data.ChartData
   :exclude-members: series, xlsx_blob, xml_bytes
   :members:
   :member-order: bysource
   :undoc-members:


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


|Plot| objects
--------------

A *plot* is a group of series all depicted using the same charting type, e.g.
bar, column, line, etc. Most charts have only a single plot; however, a chart
may have multiple, as in where a line plot appears overlaid on a bar plot in
the same chart. In the Microsoft API, this concept has the name *chart
group*. The term *plot* was chosen for |pp| to avoid the common mistake of
understanding a chart group to be a group of chart objects.

Certain properties must be set at the plot level. Some of those properties
are not present on plots of all chart types. For example, :attr:`gap_width`
is only present on a bar or column plot.

.. autoclass:: pptx.chart.plot.Plot()
   :members:
   :member-order: bysource
   :undoc-members:


|BarPlot| objects
~~~~~~~~~~~~~~~~~

The following properties are only present on bar-type plots, which includes
both bar and column charts.

.. autoclass:: pptx.chart.plot.BarPlot()
   :members:
   :member-order: bysource
   :undoc-members:


|DataLabels| objects
--------------------

A *data label* is text that labels a particular data point, usually with its
value, allowing the point to be interpreted more clearly than just visually
comparing its marker with its axis.

A |DataLabels| object is not a collection, such as a sequence, and it does
not provide access to individual data points. Rather, it provides properties
that allow all the data labels in its scope to be formatted at once.

.. autoclass:: pptx.chart.plot.DataLabels()
   :members:
   :member-order: bysource
   :undoc-members:


|Series| objects
----------------

A *series* is a sequence of data points that represent a coherent set of
observations across each of the categories in the chart. For example, on
a chart having regional categories "West", "East", and "Mid-west", a series
might be "Q1 Sales" and have values 42, 120, and 34. The series in this case
coheres around the first quarter time period.

In general, the type (class) of a series object depends upon the chart type.
The following properties are available on series objects of all types.

.. autoclass:: pptx.chart.series._BaseSeries()
   :members:
   :member-order: bysource
   :undoc-members:


|BarSeries| objects
~~~~~~~~~~~~~~~~~~~

These properties are only available on series belonging to a bar-type plot.

.. autoclass:: pptx.chart.series.BarSeries()
   :exclude-members: get_or_add_ln, ln
   :members:
   :member-order: bysource
   :undoc-members:


|LineSeries| objects
~~~~~~~~~~~~~~~~~~~~

These properties are only available on series belonging to a line-type plot.

.. autoclass:: pptx.chart.series.LineSeries()
   :members:
   :member-order: bysource
   :undoc-members:
