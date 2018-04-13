.. _chart-api:

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

.. autoclass:: pptx.chart.chart.ChartTitle()
   :members:
   :member-order: bysource


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
depends on the chart type. For example, the category axis is horizontal on
a column chart, but vertical on a bar chart.

A chart where the independent variable is in a continuous (numeric) range,
such as an XY/scatter chart, does not have a category axis. Rather it has two
value axes.

A category is perhaps most commonly a string label, such as ``'East'`` or
``'Revenue'``; however a category can also be a number or a date (although
all categories in a chart must be the same type).

When a chart's categories are dates, the category axis is generally, but not
necessarily a |DateAxis| object.

A Chart may have zero to four axes. A pie chart, for example, has neither
a category nor a value axis.

.. autoclass:: pptx.chart.axis._BaseAxis()
   :members:
   :member-order: bysource

.. autoclass:: pptx.chart.axis.CategoryAxis()
   :members:
   :member-order: bysource

.. autoclass:: pptx.chart.axis.DateAxis()
   :members:
   :member-order: bysource

.. autoclass:: pptx.chart.axis.AxisTitle()
   :members:
   :member-order: bysource


Value Axes
~~~~~~~~~~

Some axis properties are only relevant to value axes, in particular, those
related to numeric values rather than text category labels.

.. autoclass:: pptx.chart.axis.ValueAxis()
   :members:
   :member-order: bysource
   :undoc-members:


|MajorGridlines| objects
------------------------

Gridlines are the vertical and horizontal lines that extend major tick marks
of an axis across the chart to ease comparison of a data point with the axis
divisions.

.. autoclass:: pptx.chart.axis.MajorGridlines()
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


|_BasePlot| objects
-------------------

A *plot* is a group of series all depicted using the same charting type, e.g.
bar, column, line, etc. Most charts have only a single plot; however, a chart
may have multiple, as in where a line plot appears overlaid on a bar plot in
the same chart. In the Microsoft API, this concept has the name *chart
group*. The term *plot* was chosen for |pp| to avoid the common mistake of
understanding a chart group to be a group of chart objects.

Certain properties must be set at the plot level. Some of those properties
are not present on plots of all chart types. For example, :attr:`gap_width`
is only present on a bar or column plot.

.. autoclass:: pptx.chart.plot._BasePlot()
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


|BubblePlot| objects
~~~~~~~~~~~~~~~~~~~~

The following properties are only present on bubble-type plots.

.. autoclass:: pptx.chart.plot.BubblePlot()
   :members:
   :member-order: bysource
   :undoc-members:


|category.Categories| objects
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Category plots provide access to a |category.Categories| object with their
`.categories` property.

.. autoclass:: pptx.chart.category.Categories()
   :members:
   :member-order: bysource
   :undoc-members:


|category.Category| objects
~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. autoclass:: pptx.chart.category.Category()
   :members:
   :member-order: bysource
   :undoc-members:


|CategoryLevel| objects
~~~~~~~~~~~~~~~~~~~~~~~

.. autoclass:: pptx.chart.category.CategoryLevel()
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

.. autoclass:: pptx.chart.datalabel.DataLabels()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.chart.datalabel.DataLabel()
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


|AreaSeries| objects
~~~~~~~~~~~~~~~~~~~~

These properties are available on a series belonging to an area-type plot
such as AREA_STACKED.

.. autoclass:: pptx.chart.series.AreaSeries()
   :inherited-members:
   :members:
   :member-order: bysource


|BarSeries| objects
~~~~~~~~~~~~~~~~~~~

These properties are available on a series belonging to a bar-type plot. Note
that column charts are also bar-type charts.

.. autoclass:: pptx.chart.series.BarSeries()
   :inherited-members:
   :members:
   :member-order: bysource


|BubbleSeries| objects
~~~~~~~~~~~~~~~~~~~~~~

These properties are available on series belonging to a bubble chart.

.. autoclass:: pptx.chart.series.BubbleSeries()
   :inherited-members:
   :members:
   :member-order: bysource


|LineSeries| objects
~~~~~~~~~~~~~~~~~~~~

These properties are available on a series belonging to a line-type plot.

.. autoclass:: pptx.chart.series.LineSeries()
   :inherited-members:
   :members:
   :member-order: bysource


|Marker| objects
~~~~~~~~~~~~~~~~

A marker is the small shape (e.g. diamond or circle) that "marks" each
individual data point connected by a series line in a line-type plot.

.. autoclass:: pptx.chart.marker.Marker()
   :inherited-members:
   :members:
   :member-order: bysource


|PieSeries| objects
~~~~~~~~~~~~~~~~~~~

These properties are available on a series belonging to a pie chart.

.. autoclass:: pptx.chart.series.PieSeries()
   :inherited-members:
   :members:
   :member-order: bysource


|RadarSeries| objects
~~~~~~~~~~~~~~~~~~~~~

These properties are available on a series belonging to a radar chart.

.. autoclass:: pptx.chart.series.RadarSeries()
   :inherited-members:
   :members:
   :member-order: bysource


|XySeries| objects
~~~~~~~~~~~~~~~~~~

These properties are available on series belonging to an XY plot.

.. autoclass:: pptx.chart.series.XySeries()
   :members:
   :member-order: bysource
   :inherited-members:
   :undoc-members:


|Point| objects
---------------

An XY or bubble chart has a :attr:`points` attribute providing access to a
sequence of |Point| objects. That sequence supports iteration, indexed
access, and ``len()``.

.. autoclass:: pptx.chart.point.CategoryPoints()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.chart.point.BubblePoints()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.chart.point.XyPoints()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.chart.point.Point()
   :members:
   :member-order: bysource
   :undoc-members:
