
Chart Data
==========

The data behind a chart--its category labels and its series names and
values--turns out to be a pivotal object in both the construction of a new
chart and updating the data behind an existing chart.

The first role of |ChartData| is to act as a data transfer object, allowing
the data for a chart to be accumulated over multiple calls such as
`.add_series()`. This avoids the need to assemble and send a complex nested
structure of primitives to method calls.

In addition to this, |ChartData| also takes on the role of broker to
|ChartXmlWriter| and |WorkbookWriter| which know how to assemble the
``<c:chartSpace>`` XML and Excel workbook for a chart, respectively. This is
sensible because neither of these objects can operate without a chart data
instance to provide the data they need and doing so concentrates the coupling
to the latter two objects into one place.


Protocol
--------

A |ChartData| object is constructed directly as needed, and used for either
creating a new chart or for replacing the data behind an existing one.

Creating a new chart::

    >>> chart_data = ChartData()
    >>> chart_data.categories = 'Foo', 'Bar'
    >>> chart_data.add_series('Series 1', (1.2, 2.3))
    >>> chart_data.add_series('Series 2', (3.4, 4.5))

    >>> x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4)
    >>> graphic_frame = shapes.add_chart(
    >>>     XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    >>> )
    >>> chart = graphicFrame.chart

Changing the data behind an existing chart::

    >>> chart_data = ChartData()
    >>> chart_data.categories = 'Foobar', 'Barbaz', 'Bazfoo'
    >>> chart_data.add_series('New Series 1', (5.6, 6.7, 7.8))
    >>> chart_data.add_series('New Series 2', (2.3, 3.4, 4.5))
    >>> chart_data.add_series('New Series 3', (8.9, 9.1, 1.2))

    >>> chart.replace_data(chart_data)

Note that the dimensions of the replacement data can differ from that of the
existing chart.
