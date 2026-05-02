
Working with charts
===================

|pp| supports adding charts and modifying existing ones. Most chart types
other than 3D types are supported; see :ref:`supported-chart-types` below for
a member-by-member status table of the :class:`~pptx.enum.chart.XL_CHART_TYPE`
enumeration.


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

Tick labels can also be rotated, for example to prevent long category-axis
labels from overlapping. :attr:`TickLabels.rotation` is a read/write
clockwise rotation in degrees::

    category_axis.tick_labels.rotation = 45

Both ``int`` and ``float`` values are accepted. Assigning the default
(``0``) removes the explicit rotation; negative values are normalized to
the equivalent positive rotation in ``[0, 360)``.

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
have a single plot, but |pp| now supports adding an additional plot to a
chart via :meth:`Chart.add_plot` (see :ref:`combo-charts` below).

In the Microsoft API, the name *ChartGroup* is used for this object. I found
that term confusing for a long time while I was learning about MS Office
charts so I chose the name Plot for that object in |pp|.


.. _combo-charts:

Combo charts (adding a second plot)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A *combo chart* draws two (or more) chart types in a single plot area —
for example a line plotted on top of a column chart. Create one by first
building an ordinary single-plot chart and then appending a second plot
with :meth:`Chart.add_plot`::

    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    chart_data = CategoryChartData()
    chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
    chart_data.add_series('Revenue', (10, 20, 30, 40))
    chart_data.add_series('Cost',    ( 5, 12, 18, 25))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    ).chart

    line_data = CategoryChartData()
    line_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
    line_data.add_series('Trend', (7, 15, 25, 35))

    chart.add_plot(XL_CHART_TYPE.LINE_MARKERS, line_data)

The new plot reuses the existing plot's category and value axes. The
MVP supports bar/column and line variants for the added plot
(``BAR_CLUSTERED``, ``COLUMN_CLUSTERED``, ``BAR_STACKED``,
``COLUMN_STACKED``, ``BAR_STACKED_100``, ``COLUMN_STACKED_100``,
``LINE``, ``LINE_MARKERS``, ``LINE_STACKED``, ``LINE_STACKED_100``,
``LINE_MARKERS_STACKED`` and ``LINE_MARKERS_STACKED_100``). Other
chart types raise :class:`NotImplementedError`.

.. note::
   In this MVP, values for the added plot are written inline in the
   chart XML (``c:numLit`` / ``c:strLit``) rather than appended to the
   embedded Excel workbook. The chart renders correctly in PowerPoint
   and LibreOffice, but PowerPoint's **Edit Data** dialog will show
   only the original plot's data — the added plot's values are not
   surfaced into the workbook. A follow-up that synchronizes the
   workbook is planned; see ``docs/dev/analysis/combo-chart.rst`` for
   the design notes.

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

.. note::
   A newly-added doughnut or exploded-doughnut chart emits its default
   ``c:dLbls`` block with ``c:showVal val="1"``, so the numeric values appear
   on each slice as soon as the chart is opened (see
   `issue #347 <https://github.com/scanny/python-pptx/issues/347>`_).
   Earlier releases emitted ``showVal val="0"``, which left the data-label
   element wired up but every show-flag turned off — the labels existed in
   the XML but PowerPoint drew nothing. If you need a doughnut chart without
   data labels, assign ``plot.has_data_labels = False`` or toggle individual
   ``plot.data_labels.show_value`` / ``show_percentage`` / ``show_category_name``
   properties.


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

Choosing the right ChartData subclass for replace_data
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:meth:`Chart.replace_data` requires the |ChartData| subclass to match the
chart's type family, because the underlying XML shape (``c:cat``/``c:val`` vs
``c:xVal``/``c:yVal``/``c:bubbleSize``) differs between category, XY and
bubble charts. Passing the wrong subclass would silently write the wrong
shape into the chart XML, producing a file that PowerPoint refuses to open
without the infamous "repair needed" dialog (see GitHub issue
`#396 <https://github.com/scanny/python-pptx/issues/396>`_). As of |pp|
1.0.2 a mismatch raises :class:`ValueError` naming the expected subclass.

* Category-family charts (bar, column, line, pie, area, radar, doughnut,
  stock, …) require :class:`~pptx.chart.data.CategoryChartData` (or the
  legacy alias :class:`~pptx.chart.data.ChartData`).
* XY / scatter charts require :class:`~pptx.chart.data.XyChartData`.
* Bubble charts require :class:`~pptx.chart.data.BubbleChartData`.

If you need to change a chart's *type* as well as its data, replace the
entire shape instead: remove the chart graphic frame and add a fresh one
with :meth:`~pptx.shapes.shapetree.SlideShapes.add_chart`.

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

.. _supported-chart-types:

Chart-type reference
~~~~~~~~~~~~~~~~~~~~

Every member of :class:`~pptx.enum.chart.XL_CHART_TYPE` (the ``XlChartType``
enumeration from the Microsoft API) is listed below with its current
support level in |pp|.

* **create** — can be passed to :meth:`SlideShapes.add_chart` to author a
  new chart.
* **read** — returned by :attr:`Chart.chart_type` / :attr:`GraphicFrame.chart_type`
  when loading a presentation that contains a chart of this type, with chart
  contents readable through the normal |Chart| / :class:`.plot._BasePlot`
  / |Series| object model.
* **round-trip** — recognised on read and written back unchanged on save,
  but the chart's contents are not introspectable through |pp|'s object
  model. Today this only applies to the Office 2016+ "chartex" chart types
  surfaced via the :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX` sentinel; see
  the :ref:`chartex section <chartex-charts>` below.
* **not yet** — the chart type exists in the Microsoft enumeration but no
  |pp| support is wired up. Passing it to :meth:`SlideShapes.add_chart`
  raises :class:`NotImplementedError`. In practice 3D variants, cone /
  cylinder / pyramid variants, stock charts, surface charts, and
  ``BAR_OF_PIE`` / ``PIE_OF_PIE`` fall in this bucket.

.. list-table::
   :header-rows: 1
   :widths: 40 20 40

   * - ``XL_CHART_TYPE`` member
     - Support
     - Notes
   * - ``AREA``
     - create, read
     - Area.
   * - ``AREA_STACKED``
     - create, read
     - Stacked Area.
   * - ``AREA_STACKED_100``
     - create, read
     - 100% Stacked Area.
   * - ``BAR_CLUSTERED``
     - create, read
     - Clustered Bar. Also allowed as a secondary plot in a combo chart.
   * - ``BAR_OF_PIE``
     - not yet
     - Bar of Pie.
   * - ``BAR_STACKED``
     - create, read
     - Stacked Bar. Also allowed as a secondary plot in a combo chart.
   * - ``BAR_STACKED_100``
     - create, read
     - 100% Stacked Bar. Also allowed as a secondary plot in a combo
       chart.
   * - ``BUBBLE``
     - create, read
     - Bubble. Requires :class:`~pptx.chart.data.BubbleChartData`.
   * - ``BUBBLE_THREE_D_EFFECT``
     - create, read
     - Bubble with 3D effects. Requires
       :class:`~pptx.chart.data.BubbleChartData`.
   * - ``COLUMN_CLUSTERED``
     - create, read
     - Clustered Column. Also allowed as a secondary plot in a combo
       chart.
   * - ``COLUMN_STACKED``
     - create, read
     - Stacked Column. Also allowed as a secondary plot in a combo chart.
   * - ``COLUMN_STACKED_100``
     - create, read
     - 100% Stacked Column. Also allowed as a secondary plot in a combo
       chart.
   * - ``CONE_BAR_CLUSTERED``
     - not yet
     - Clustered Cone Bar.
   * - ``CONE_BAR_STACKED``
     - not yet
     - Stacked Cone Bar.
   * - ``CONE_BAR_STACKED_100``
     - not yet
     - 100% Stacked Cone Bar.
   * - ``CONE_COL``
     - not yet
     - 3D Cone Column.
   * - ``CONE_COL_CLUSTERED``
     - not yet
     - Clustered Cone Column.
   * - ``CONE_COL_STACKED``
     - not yet
     - Stacked Cone Column.
   * - ``CONE_COL_STACKED_100``
     - not yet
     - 100% Stacked Cone Column.
   * - ``CYLINDER_BAR_CLUSTERED``
     - not yet
     - Clustered Cylinder Bar.
   * - ``CYLINDER_BAR_STACKED``
     - not yet
     - Stacked Cylinder Bar.
   * - ``CYLINDER_BAR_STACKED_100``
     - not yet
     - 100% Stacked Cylinder Bar.
   * - ``CYLINDER_COL``
     - not yet
     - 3D Cylinder Column.
   * - ``CYLINDER_COL_CLUSTERED``
     - not yet
     - Clustered Cylinder Column.
   * - ``CYLINDER_COL_STACKED``
     - not yet
     - Stacked Cylinder Column.
   * - ``CYLINDER_COL_STACKED_100``
     - not yet
     - 100% Stacked Cylinder Column.
   * - ``DOUGHNUT``
     - create, read
     - Doughnut.
   * - ``DOUGHNUT_EXPLODED``
     - create, read
     - Exploded Doughnut.
   * - ``LINE``
     - create, read
     - Line. Also allowed as a secondary plot in a combo chart.
   * - ``LINE_MARKERS``
     - create, read
     - Line with Markers. Also allowed as a secondary plot in a combo
       chart.
   * - ``LINE_MARKERS_STACKED``
     - create, read
     - Stacked Line with Markers. Also allowed as a secondary plot in a
       combo chart.
   * - ``LINE_MARKERS_STACKED_100``
     - create, read
     - 100% Stacked Line with Markers. Also allowed as a secondary plot
       in a combo chart.
   * - ``LINE_STACKED``
     - create, read
     - Stacked Line. Also allowed as a secondary plot in a combo chart.
   * - ``LINE_STACKED_100``
     - create, read
     - 100% Stacked Line. Also allowed as a secondary plot in a combo
       chart.
   * - ``PIE``
     - create, read
     - Pie.
   * - ``PIE_EXPLODED``
     - create, read
     - Exploded Pie.
   * - ``PIE_OF_PIE``
     - not yet
     - Pie of Pie.
   * - ``PYRAMID_BAR_CLUSTERED``
     - not yet
     - Clustered Pyramid Bar.
   * - ``PYRAMID_BAR_STACKED``
     - not yet
     - Stacked Pyramid Bar.
   * - ``PYRAMID_BAR_STACKED_100``
     - not yet
     - 100% Stacked Pyramid Bar.
   * - ``PYRAMID_COL``
     - not yet
     - 3D Pyramid Column.
   * - ``PYRAMID_COL_CLUSTERED``
     - not yet
     - Clustered Pyramid Column.
   * - ``PYRAMID_COL_STACKED``
     - not yet
     - Stacked Pyramid Column.
   * - ``PYRAMID_COL_STACKED_100``
     - not yet
     - 100% Stacked Pyramid Column.
   * - ``RADAR``
     - create, read
     - Radar.
   * - ``RADAR_FILLED``
     - create, read
     - Filled Radar.
   * - ``RADAR_MARKERS``
     - create, read
     - Radar with Data Markers.
   * - ``STOCK_HLC``
     - not yet
     - High-Low-Close.
   * - ``STOCK_OHLC``
     - not yet
     - Open-High-Low-Close.
   * - ``STOCK_VHLC``
     - not yet
     - Volume-High-Low-Close.
   * - ``STOCK_VOHLC``
     - not yet
     - Volume-Open-High-Low-Close.
   * - ``SURFACE``
     - not yet
     - 3D Surface.
   * - ``SURFACE_TOP_VIEW``
     - not yet
     - Surface (Top View).
   * - ``SURFACE_TOP_VIEW_WIREFRAME``
     - not yet
     - Surface (Top View wireframe).
   * - ``SURFACE_WIREFRAME``
     - not yet
     - 3D Surface (wireframe).
   * - ``THREE_D_AREA``
     - not yet
     - 3D Area.
   * - ``THREE_D_AREA_STACKED``
     - not yet
     - 3D Stacked Area.
   * - ``THREE_D_AREA_STACKED_100``
     - not yet
     - 3D 100% Stacked Area.
   * - ``THREE_D_BAR_CLUSTERED``
     - not yet
     - 3D Clustered Bar.
   * - ``THREE_D_BAR_STACKED``
     - not yet
     - 3D Stacked Bar.
   * - ``THREE_D_BAR_STACKED_100``
     - not yet
     - 3D 100% Stacked Bar.
   * - ``THREE_D_COLUMN``
     - not yet
     - 3D Column.
   * - ``THREE_D_COLUMN_CLUSTERED``
     - not yet
     - 3D Clustered Column.
   * - ``THREE_D_COLUMN_STACKED``
     - not yet
     - 3D Stacked Column.
   * - ``THREE_D_COLUMN_STACKED_100``
     - not yet
     - 3D 100% Stacked Column.
   * - ``THREE_D_LINE``
     - not yet
     - 3D Line.
   * - ``THREE_D_PIE``
     - not yet
     - 3D Pie.
   * - ``THREE_D_PIE_EXPLODED``
     - not yet
     - Exploded 3D Pie.
   * - ``XY_SCATTER``
     - create, read
     - Scatter. Requires :class:`~pptx.chart.data.XyChartData`.
   * - ``XY_SCATTER_LINES``
     - create, read
     - Scatter with Lines. Requires
       :class:`~pptx.chart.data.XyChartData`.
   * - ``XY_SCATTER_LINES_NO_MARKERS``
     - create, read
     - Scatter with Lines and No Data Markers. Requires
       :class:`~pptx.chart.data.XyChartData`.
   * - ``XY_SCATTER_SMOOTH``
     - create, read
     - Scatter with Smoothed Lines. Requires
       :class:`~pptx.chart.data.XyChartData`.
   * - ``XY_SCATTER_SMOOTH_NO_MARKERS``
     - create, read
     - Scatter with Smoothed Lines and No Data Markers. Requires
       :class:`~pptx.chart.data.XyChartData`.
   * - ``UNSUPPORTED_CHARTEX``
     - round-trip
     - python-pptx sentinel (no Microsoft counterpart). Returned by
       :attr:`GraphicFrame.chart_type` for Office 2016+ chartex shapes —
       funnel, treemap, sunburst, waterfall, histogram / Pareto,
       box-and-whisker, and map — so they round-trip unchanged. See the
       :ref:`chartex section <chartex-charts>` below and `issue #386`_ for
       the roadmap that extends |pp| to read and author these types.

.. _issue #386: https://github.com/scanny/python-pptx/issues/386


.. _chartex-charts:

Office 2016+ extended chart types (funnel, treemap, sunburst, waterfall, histogram,
box-and-whisker, map)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

These newer chart types are written in the ``cx:`` ("chartex") XML namespace
introduced with Office 2016. |pp| does not yet read or modify their contents, but
it does surface them on ``slide.shapes`` so you can detect them and so a plain
load/save round-trip preserves their XML unchanged.

On a shape that wraps one of these charts:

* :attr:`GraphicFrame.has_chart` returns ``True``
* :attr:`GraphicFrame.has_chartex` returns ``True``
* :attr:`GraphicFrame.chart_type` returns :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`
* :attr:`GraphicFrame.shape_type` returns :attr:`MSO_SHAPE_TYPE.CHART`
* :attr:`GraphicFrame.chart` raises :exc:`NotImplementedError`

Use ``has_chartex`` to branch your code when you need to skip these shapes, for
example::

    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        if shape.has_chartex:
            # -- Office 2016+ extended chart (funnel, treemap, ...) — not yet
            # -- readable, but it will round-trip unchanged --
            continue
        chart = shape.chart
        ...

PowerPoint frequently wraps these chart shapes in an ``mc:AlternateContent``
envelope (providing a legacy chart fallback for older viewers); |pp| unwraps the
``mc:Choice`` branch transparently so the chartex shape is visible on
``slide.shapes``, and keeps the ``mc:Fallback`` subtree intact on save.

Roadmap. The passthrough landed here is the MVP step of
`issue #386`_ ("chartex passthrough"). Incremental follow-ups — typed
element classes for the ``cx:`` shapes, a ``Chart``-compatible
introspection surface, then authoring support for each chartex type
(funnel / treemap / sunburst / waterfall / histogram / Pareto /
box-and-whisker / map) — are tracked on that issue; see
``docs/dev/analysis/chartex-foundation.rst`` for the staged design
notes. Until those land the sentinel :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`
is what :attr:`GraphicFrame.chart_type` returns for these shapes.
