
Working with charts
===================

|pp| supports adding charts and modifying existing ones. Most chart types
other than 3D types are supported.


Adding a chart
--------------

The following code adds a single-series column chart in a new presentation::

    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    # create presentation with 1 slide ------
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # define chart data ---------------------
    chart_data = CategoryChartData()
    chart_data.categories = ['East', 'West', 'Midwest']
    chart_data.add_series('Series 1', (19.2, 21.4, 16.7))

    # add chart to slide --------------------
    x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    )

    prs.save('chart-01.pptx')

.. image:: /_static/img/chart-01.png


Customizing things a bit
------------------------

The remaining code will leave out code we've already seen and only show
imports, for example, when they're used for the first time, just to keep the
focus on the new bits. Let's create a multi-series chart to use for these
examples::

    chart_data = ChartData()
    chart_data.categories = ['East', 'West', 'Midwest']
    chart_data.add_series('Q1 Sales', (19.2, 21.4, 16.7))
    chart_data.add_series('Q2 Sales', (22.3, 28.6, 15.2))
    chart_data.add_series('Q3 Sales', (20.4, 26.3, 14.2))

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    )

    chart = graphic_frame.chart

.. image:: /_static/img/chart-02.png

Notice that we captured the shape reference returned by the
:meth:`~.SlideShapes.add_chart` call as ``graphic_frame`` and then extracted
the chart object from the graphic frame using its
:attr:`~.GraphicFrame.chart` property. We'll need the chart reference to get
to the properties we'll need in the next steps. The
:meth:`~.SlideShapes.add_chart` method doesn't directly return the chart
object. That's because a chart is not itself a shape. Rather it's a graphical
(DrawingML) object *contained* in the graphic frame shape. Tables work this
way too, also being contained in a graphic frame shape.


XY and Bubble charts
--------------------

The charts so far use a *discrete* set of values for the independent variable
(the X axis, roughly speaking). These are perfect when your values fall into
a well-defined set of categories. However, there are many cases, particularly
in science and engineering, where the independent variable is a continuous
value, such as temperature or frequency. These are supported in PowerPoint by
XY (aka. scatter) charts. A bubble chart is essentially an XY chart where the
marker size is used to reflect an additional value, effectively adding
a third dimension to the chart.

Because the independent variable is continuous, in general, the series do not
all share the same X values. This requires a somewhat different data
structure and that is provided for by distinct |XyChartData| and
|BubbleChartData| objects used to specify the data behind charts of these
types::

    chart_data = XyChartData()

    series_1 = chart_data.add_series('Model 1')
    series_1.add_data_point(0.7, 2.7)
    series_1.add_data_point(1.8, 3.2)
    series_1.add_data_point(2.6, 0.8)

    series_2 = chart_data.add_series('Model 2')
    series_2.add_data_point(1.3, 3.7)
    series_2.add_data_point(2.7, 2.3)
    series_2.add_data_point(1.6, 1.8)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER, x, y, cx, cy, chart_data
    ).chart

.. image:: /_static/img/chart-08.png

Creation of a bubble chart is very similar, having an additional value for each data point that specifies the bubble size::

    chart_data = BubbleChartData()

    series_1 = chart_data.add_series('Series 1')
    series_1.add_data_point(0.7, 2.7, 10)
    series_1.add_data_point(1.8, 3.2, 4)
    series_1.add_data_point(2.6, 0.8, 8)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BUBBLE, x, y, cx, cy, chart_data
    ).chart

.. image:: /_static/img/chart-09.png



Axes
----

Let's change up the category and value axes a bit::

    from pptx.enum.chart import XL_TICK_MARK
    from pptx.util import Pt

    category_axis = chart.category_axis
    category_axis.has_major_gridlines = True
    category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
    category_axis.tick_labels.font.italic = True
    category_axis.tick_labels.font.size = Pt(24)

    value_axis = chart.value_axis
    value_axis.maximum_scale = 50.0
    value_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
    value_axis.has_minor_gridlines = True

    tick_labels = value_axis.tick_labels
    tick_labels.number_format = '0"%"'
    tick_labels.font.bold = True
    tick_labels.font.size = Pt(14)

.. image:: /_static/img/chart-03.png

Okay, that was probably going a bit too far. But it gives us an idea of the
kinds of things we can do with the value and category axes. Let's undo this
part and go back to the version we had before.


Secondary value axis
~~~~~~~~~~~~~~~~~~~~

A chart may carry a *secondary value axis* — a second value axis (conventionally
rendered on the right side of the chart) against which one or more series may
be plotted. This is useful when two series have widely different magnitudes
and a single axis would cause the smaller series to appear flat.

When a chart read from a presentation file contains a secondary value axis,
access it through :attr:`Chart.secondary_value_axis`::

    if chart.has_secondary_value_axis:
        secondary_axis = chart.secondary_value_axis
        secondary_axis.maximum_scale = 100.0
        secondary_axis.tick_labels.font.size = Pt(12)

:attr:`Chart.has_secondary_value_axis` lets you test for its presence without
raising. Accessing :attr:`Chart.secondary_value_axis` on a chart that has no
secondary value axis raises :class:`ValueError`.

An XY/scatter chart has two value axes (one for X and one for Y), but neither
is considered a *secondary* axis in this sense; both are primary axes, so
:attr:`has_secondary_value_axis` is always ``False`` for XY/scatter charts.

.. note::
   Creating a new secondary value axis (and assigning series to it) in a chart
   that does not already have one requires combo-chart plumbing that is not
   yet exposed by the library. See `issue #141
   <https://github.com/scanny/python-pptx/issues/141>`_ for status.


Data Labels
-----------

Let's add some data labels so we can see exactly what the value for each bar
is::

    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_LABEL_POSITION

    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels

    data_labels.font.size = Pt(13)
    data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
    data_labels.position = XL_LABEL_POSITION.INSIDE_END

.. image:: /_static/img/chart-04.png

Here we needed to access a Plot object to gain access to the data labels.
A plot is like a sub-chart, containing one or more series and drawn as
a particular chart type, like column or line. This distinction is needed for
charts that combine more than one type, like a line chart appearing on top of
a column chart. A chart like this would have two plot objects, one for the
series appearing as columns and the other for the lines. Most charts only
have a single plot and |pp| doesn't yet support creating multi-plot charts,
but you can access multiple plots on a chart that already has them.

In the Microsoft API, the name *ChartGroup* is used for this object. I found
that term confusing for a long time while I was learning about MS Office
charts so I chose the name Plot for that object in |pp|.

Collection-level text-frame properties (such as whether label text should
word-wrap, the autofit strategy, vertical anchor, and internal margins) are
available on :attr:`DataLabels.text_frame`. This |TextFrame| wraps the
``c:dLbls/c:txPr`` element that holds these settings for every data label in
the collection::

    data_labels = plot.data_labels
    data_labels.text_frame.word_wrap = False

Note this is distinct from :attr:`DataLabel.text_frame` on an individual
|DataLabel|, which exposes the *custom text content* for a single data label
(backed by ``c:dLbl/c:tx/c:rich``). The collection-level property is for
formatting, not content.


Error Bars
----------

Each series on a bar, column, line, area, XY, or bubble chart can be annotated
with *error bars* (confidence-interval whiskers). |pp| 1.x adds a minimum-viable
API for attaching a set of error bars to a series::

    from pptx.enum.chart import XL_ERROR_BAR_TYPE, XL_ERROR_BAR_INCLUDE

    series = chart.plots[0].series[0]
    series.set_error_bars(
        type_=XL_ERROR_BAR_TYPE.PERCENT,
        value=10.0,
        include=XL_ERROR_BAR_INCLUDE.BOTH,
    )

    # subsequently, read and adjust:
    assert series.has_error_bars is True
    series.error_bars.value = 5.0          # change magnitude
    series.error_bars.end_cap = False      # drop the "T" caps

The error-bar type (``XL_ERROR_BAR_TYPE``) can be ``FIXED_VALUE``, ``PERCENT``,
``STDEV`` (a multiplier on the series' standard deviation), or ``STERROR``.
``XL_ERROR_BAR_TYPE.CUSTOM`` is writable but in the 1.x series |pp| does not
populate the ``plus``/``minus`` references that hold per-point magnitudes — use
the underlying ``series._element`` to edit those directly if needed. Removing
error bars is as easy as ``series.error_bars = None``.


Legend
------

A legend is often useful to have on a chart, to give a name to each series
and help a reader tell which one is which::

    from pptx.enum.chart import XL_LEGEND_POSITION

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False

.. image:: /_static/img/chart-05.png

Nice! Okay, let's try some other chart types.


Line Chart
----------

A line chart is added pretty much the same way as a bar or column chart, the
main difference being the chart type provided in the :meth:`add_chart` call::

    chart_data = ChartData()
    chart_data.categories = ['Q1 Sales', 'Q2 Sales', 'Q3 Sales']
    chart_data.add_series('West',    (32.2, 28.4, 34.7))
    chart_data.add_series('East',    (24.3, 30.6, 20.2))
    chart_data.add_series('Midwest', (20.4, 18.3, 26.2))

    x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data
    ).chart

    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.series[0].smooth = True

.. image:: /_static/img/chart-06.png

I switched the categories and series data here to better suit a line chart.
You can see the line for the "West" region is *smoothed* into a curve while
the other two have their points connected with straight line segments.


Pie Chart
---------

A pie chart is a little special in that it only ever has a single series and
doesn't have any axes::

    chart_data = ChartData()
    chart_data.categories = ['West', 'East', 'North', 'South', 'Other']
    chart_data.add_series('Series 1', (0.135, 0.324, 0.180, 0.235, 0.126))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data
    ).chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    chart.plots[0].has_data_labels = True
    data_labels = chart.plots[0].data_labels
    data_labels.number_format = '0%'
    data_labels.position = XL_LABEL_POSITION.OUTSIDE_END

.. image:: /_static/img/chart-07.png


Refreshing cached values after an external workbook edit
--------------------------------------------------------

A chart created by python-pptx (or PowerPoint) stores its data in two places:

* an embedded Excel workbook (``.xlsx``) that acts as the authored source
  of truth, and
* a set of cached values in the chart XML itself — the ``<c:numCache>``
  and ``<c:strCache>`` elements under each series reference.

PowerPoint only re-reads the workbook on an explicit *Refresh Data* action
(``Alt+F5`` in the Chart Design ribbon). Until that happens, what you see
in the slide is the cached copy. That means if you edit the embedded
workbook directly — for example by extracting the ``.xlsx`` part,
modifying cell values, and writing it back — the chart in the rendered
slide will keep showing the old numbers and labels.

:meth:`Chart.update_cached_values` rewrites those caches in-place from
the embedded workbook::

    from pptx import Presentation

    prs = Presentation("deck-with-chart.pptx")
    chart = prs.slides[0].shapes[0].chart

    # ... externally modify chart._workbook.xlsx_part.blob, or use a
    # helper that opens the embedded xlsx, edits it, and writes it back.

    chart.update_cached_values()
    prs.save("deck-refreshed.pptx")

The method parses the embedded xlsx, resolves every ``<c:f>`` cell
reference under the chart, and rewrites the neighbouring
``<c:numCache>`` / ``<c:strCache>`` subtree so what PowerPoint renders
matches the workbook. Cell references that cannot be resolved
(multi-range formulas, links to an *external* workbook, unknown
sheet names outside the embedded file) are left untouched. The method
is a no-op for charts without an embedded workbook.


Odds & Ends
-----------

This should be enough to get you started with adding charts to your
presentation with |pp|. There are more details in the API documentation for
charts here: :ref:`chart-api`


About colors
~~~~~~~~~~~~

By default, the colors assigned to each series in a chart are the theme
colors Accent 1 through Accent 6, in that order. If you have more than six
series, darker and lighter versions of those same colors are used. While it's
possible to assign specific colors to data points (bar, line, pie segment,
etc.) for at least some chart types, the best strategy to start with is
changing the theme colors in your starting "template" presentation.

When :meth:`Chart.replace_data` adds more series than the chart originally
had, each new series element is cloned from the last existing one. If the
source series has an explicitly-set sRGB fill color, that explicit color is
replaced on each clone with a theme-accent reference
(``a:schemeClr val="accent1..6"``) rotating through the six theme accents.
This keeps newly-added series rendering with the theme's accent palette
rather than all appearing in the source series's color (see GitHub issue
#529).

Extended chart-style values (Office 2010+)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The :attr:`Chart.chart_style` property accepts two ranges:

* ``1..48`` — the classic *chart style* gallery introduced with Office 2007. Stored
  as a bare ``<c:style val="N"/>`` child of the chart part.
* ``101..148`` (and, in principle, any ``xsd:unsignedByte`` above 48) — the *extended*
  chart styles introduced with Office 2010. These are stored as a ``<c14:style val="N"/>``
  element wrapped in a ``<mc:AlternateContent>`` / ``<mc:Choice Requires="c14">`` element,
  together with a ``<c:style val="M"/>`` fallback (``M = N mod 100``) for Office 2007-era
  readers. The extended styles let PowerPoint pick pure *Accent-N* colours for every
  series in a chart with six or more series instead of darker/lighter shaded alternates of
  Accent 1 through Accent 6. See GitHub issue
  `#516 <https://github.com/scanny/python-pptx/issues/516>`_ for background.

Before |pp| 1.0.1, reading :attr:`Chart.chart_style` returned ``None`` when the chart
was authored by PowerPoint with an extended ``c14:style`` value (which is the common
case for charts edited in modern PowerPoint). It now transparently surfaces the extended
value, and writes the full ``<mc:AlternateContent>`` wrapper when you assign a value
greater than 48.
