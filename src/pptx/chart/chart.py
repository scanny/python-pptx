"""Chart-related objects such as Chart and ChartTitle."""

from __future__ import annotations

from collections.abc import Sequence

from lxml import etree

from pptx.chart.axis import CategoryAxis, DateAxis, ValueAxis
from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData
from pptx.chart.legend import Legend
from pptx.chart.plot import PlotFactory, PlotTypeInspector
from pptx.chart.series import SeriesCollection
from pptx.chart.xlsx import WorkbookReader, parse_sheet_range_ref
from pptx.chart.xmlwriter import SeriesXmlRewriterFactory, _PlotFragmentBuilder
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.shared import ElementProxy, PartElementProxy
from pptx.text.text import Font, TextFrame
from pptx.util import lazyproperty


# -- chart-type families for Chart.replace_data() compatibility check (#396). --
# -- Bubble and XY chart types use distinct XML shapes (c:xVal/c:yVal and --
# -- c:bubbleSize) that are incompatible with the c:cat/c:val shape produced --
# -- for category charts; writing the wrong shape leaves PowerPoint unable to --
# -- read the file, producing a "repair needed" dialog. --
_BUBBLE_CHART_TYPES = frozenset(
    {
        XL_CHART_TYPE.BUBBLE,
        XL_CHART_TYPE.BUBBLE_THREE_D_EFFECT,
    }
)
_XY_CHART_TYPES = frozenset(
    {
        XL_CHART_TYPE.XY_SCATTER,
        XL_CHART_TYPE.XY_SCATTER_LINES,
        XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
        XL_CHART_TYPE.XY_SCATTER_SMOOTH,
        XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
    }
)


class Chart(PartElementProxy):
    """A chart object."""

    def __init__(self, chartSpace, chart_part):
        super(Chart, self).__init__(chartSpace, chart_part)
        self._chartSpace = chartSpace

    def add_plot(self, chart_type, chart_data):
        """Append a new plot of *chart_type* to this chart (issue #338).

        This is the entry-point for creating a *combo chart* — a chart that
        displays two (or more) chart types sharing a single plot area, such
        as a column plot with a line plot overlaid. *chart_type* must be a
        member of :ref:`XlChartType`; only the bar/column and line variants
        are supported in this MVP release (``BAR_CLUSTERED``,
        ``COLUMN_CLUSTERED``, ``BAR_STACKED``, ``COLUMN_STACKED``,
        ``BAR_STACKED_100``, ``COLUMN_STACKED_100``, ``LINE``,
        ``LINE_MARKERS``, ``LINE_STACKED``, ``LINE_STACKED_100``,
        ``LINE_MARKERS_STACKED`` and ``LINE_MARKERS_STACKED_100``). Other
        chart types raise |NotImplementedError|.

        *chart_data* is a |CategoryChartData| instance whose category
        labels must match those already on this chart (they are written
        only for schema conformance — the axis labels are drawn from the
        first plot). Series values are stored inline in the chart XML
        (`c:numLit` / `c:strLit`) rather than in the embedded Excel
        workbook. This means the new plot renders correctly in PowerPoint
        and LibreOffice but its values are not surfaced to PowerPoint's
        "Edit Data" dialog for that plot; see
        :file:`docs/dev/analysis/combo-chart.rst` for discussion of the
        F5-dependent workbook synchronization that would lift this limit.

        The added plot reuses the `c:axId` values of the chart's existing
        back-most plot, so both plots share the category and value axes.
        The new plot is appended after existing xCharts so it is drawn in
        front of them.

        Returns the newly-created |BarPlot| or |LinePlot| object.
        """
        plotArea = self._chartSpace.plotArea
        xCharts = plotArea.xCharts
        if not xCharts:
            raise ValueError("cannot add plot: chart has no existing plot")
        first_xChart = xCharts[0]
        axId_elms = first_xChart.findall(qn("c:axId"))
        if len(axId_elms) < 2:
            raise ValueError(
                "cannot add plot: existing plot is missing axId elements"
            )
        cat_axId = axId_elms[0].get("val")
        val_axId = axId_elms[1].get("val")
        base_idx = len(plotArea.sers)
        builder = _PlotFragmentBuilder.new(chart_type, chart_data, cat_axId, val_axId)
        new_xChart = builder.element
        # --- re-index series to avoid idx/order collisions with existing series ---
        for offset, ser in enumerate(new_xChart.findall(qn("c:ser"))):
            idx_elm = ser.find(qn("c:idx"))
            order_elm = ser.find(qn("c:order"))
            if idx_elm is not None:
                idx_elm.set("val", str(base_idx + offset))
            if order_elm is not None:
                order_elm.set("val", str(base_idx + offset))
        # --- insert after the last existing xChart but before any axis elements ---
        last_xChart = xCharts[-1]
        last_xChart.addnext(new_xChart)
        return PlotFactory(new_xChart, self)

    @property
    def category_axis(self):
        """
        The category axis of this chart. In the case of an XY or Bubble
        chart, this is the X axis. Raises |ValueError| if no category
        axis is defined (as is the case for a pie chart, for example).
        """
        catAx_lst = self._chartSpace.catAx_lst
        if catAx_lst:
            return CategoryAxis(catAx_lst[0])

        dateAx_lst = self._chartSpace.dateAx_lst
        if dateAx_lst:
            return DateAxis(dateAx_lst[0])

        valAx_lst = self._chartSpace.valAx_lst
        if valAx_lst:
            return ValueAxis(valAx_lst[0])

        raise ValueError("chart has no category axis")

    @property
    def chart_style(self):
        """Read/write integer index of chart style used to format this chart.

        Plain chart styles are in the range 1 to 48. Starting with Office 2010, an *extended*
        style index in the range 101 to 148 is also supported; this index is serialised in a
        `c14:style` element wrapped in `mc:AlternateContent` and is what PowerPoint writes
        when you pick one of the "Accent-N" chart-style variants that keep each series in a
        pure accent colour (instead of shaded alternates) for charts with six or more series
        (issue #516).

        Returns the extended (`c14:style`) value when present, otherwise the plain
        (`c:style`) value, otherwise `None` (the default style is used).

        Assigning `None` removes any explicit setting (plain and extended). Assigning a value
        in 1..48 writes a plain `c:style`; assigning 49..255 writes an `mc:AlternateContent`
        wrapper carrying the extended `c14:style` value and a `c:style` fallback set to
        `value % 100` (so Office 2007-era readers still render the chart).
        """
        return self._chartSpace.chart_style_ex_val

    @chart_style.setter
    def chart_style(self, value):
        self._chartSpace.set_chart_style_ex_val(value)

    @property
    def chart_title(self):
        """A |ChartTitle| object providing access to title properties.

        Calling this property is destructive in the sense it adds a chart
        title element (`c:title`) to the chart XML if one is not already
        present. Use :attr:`has_title` to test for presence of a chart title
        non-destructively.
        """
        return ChartTitle(self._element.get_or_add_title())

    @property
    def chart_type(self):
        """Member of :ref:`XlChartType` enumeration specifying type of this chart.

        If the chart has two plots, for example, a line plot overlayed on a bar plot,
        the type reported is for the first (back-most) plot. Read-only.
        """
        first_plot = self.plots[0]
        return PlotTypeInspector.chart_type(first_plot)

    @lazyproperty
    def font(self):
        """Font object controlling text format defaults for this chart."""
        defRPr = self._chartSpace.get_or_add_txPr().p_lst[0].get_or_add_pPr().get_or_add_defRPr()
        return Font(defRPr)

    @property
    def has_legend(self):
        """
        Read/write boolean, |True| if the chart has a legend. Assigning
        |True| causes a legend to be added to the chart if it doesn't already
        have one. Assigning False removes any existing legend definition
        along with any existing legend settings.
        """
        return self._chartSpace.chart.has_legend

    @has_legend.setter
    def has_legend(self, value):
        self._chartSpace.chart.has_legend = bool(value)

    @property
    def has_title(self):
        """Read/write boolean, specifying whether this chart has a title.

        Assigning |True| causes a title to be added if not already present.
        Assigning |False| removes any existing title along with its text and
        settings.
        """
        title = self._chartSpace.chart.title
        if title is None:
            return False
        return True

    @has_title.setter
    def has_title(self, value):
        chart = self._chartSpace.chart
        if bool(value) is False:
            chart._remove_title()
            autoTitleDeleted = chart.get_or_add_autoTitleDeleted()
            autoTitleDeleted.val = True
            return
        chart.get_or_add_title()

    @property
    def legend(self):
        """
        A |Legend| object providing access to the properties of the legend
        for this chart.
        """
        legend_elm = self._chartSpace.chart.legend
        if legend_elm is None:
            return None
        return Legend(legend_elm)

    @lazyproperty
    def plots(self):
        """
        The sequence of plots in this chart. A plot, called a *chart group*
        in the Microsoft API, is a distinct sequence of one or more series
        depicted in a particular charting type. For example, a chart having
        a series plotted as a line overlaid on three series plotted as
        columns would have two plots; the first corresponding to the three
        column series and the second to the line series. Plots are sequenced
        in the order drawn, i.e. back-most to front-most. Supports *len()*,
        membership (e.g. ``p in plots``), iteration, slicing, and indexed
        access (e.g. ``plot = plots[i]``).
        """
        plotArea = self._chartSpace.chart.plotArea
        return _Plots(plotArea, self)

    def replace_data(self, chart_data):
        """Replace chart data with categories/values from the |ChartData| object *chart_data*.

        Existing series formatting (fill, line, etc.) is preserved on series that already
        exist in the chart. When *chart_data* adds more series than currently exist, new
        ``c:ser`` elements are cloned from the last existing series. Any explicit sRGB
        color on the source series is replaced on each clone with a cycling theme-accent
        reference (``a:schemeClr val="accent1..6"``) so newly added series adopt the
        theme's accent color rotation rather than all appearing in the source series's
        color (see GitHub issue #529).

        Raises |ValueError| when *chart_data* is incompatible with the chart-type of
        this chart (GitHub issue #396). XY (scatter) charts require an
        |XyChartData| instance and bubble charts require a |BubbleChartData|
        instance; passing a |CategoryChartData| (or plain |ChartData|) to one of
        those chart types would otherwise produce a file that PowerPoint cannot
        open without offering to "repair" it. Category-type charts (bar, line,
        pie, area, radar, doughnut, etc.) require a |CategoryChartData|.
        """
        self._validate_chart_data_type(chart_data)
        rewriter = SeriesXmlRewriterFactory(self.chart_type, chart_data)
        rewriter.replace_series_data(self._chartSpace)
        self._workbook.update_from_xlsx_blob(chart_data.xlsx_blob)

    def _validate_chart_data_type(self, chart_data):
        """Raise |ValueError| when *chart_data* is incompatible with the chart-type.

        See :meth:`replace_data` for rationale (GitHub issue #396). The error
        message names the expected |ChartData| subclass so the caller knows
        exactly which type to construct.
        """
        chart_type = self.chart_type
        if chart_type in _BUBBLE_CHART_TYPES:
            if not isinstance(chart_data, BubbleChartData):
                raise ValueError(
                    "bubble chart requires BubbleChartData instance for "
                    "Chart.replace_data(), got %s" % type(chart_data).__name__
                )
            return
        if chart_type in _XY_CHART_TYPES:
            # -- BubbleChartData is a subclass of XyChartData but carries --
            # -- per-point bubble-size values that the XY rewriter ignores; --
            # -- treat it as incompatible so callers get a clear error rather --
            # -- than silently dropped data. --
            if not isinstance(chart_data, XyChartData) or isinstance(
                chart_data, BubbleChartData
            ):
                raise ValueError(
                    "XY (scatter) chart requires XyChartData instance for "
                    "Chart.replace_data(), got %s" % type(chart_data).__name__
                )
            return
        # -- all other chart types expect a category-shaped data set. --
        # -- XyChartData / BubbleChartData are siblings of CategoryChartData --
        # -- (both inherit from _BaseChartData), so an isinstance check for --
        # -- CategoryChartData here excludes them cleanly. --
        if not isinstance(chart_data, CategoryChartData):
            raise ValueError(
                "category chart requires CategoryChartData instance for "
                "Chart.replace_data(), got %s" % type(chart_data).__name__
            )

    def update_cached_values(self):
        """Rewrite cached chart values from the embedded Excel worksheet.

        PowerPoint stores the values displayed in a chart in two places: the
        embedded `.xlsx` workbook that is the authored data source, and a
        set of XML caches (``c:numCache`` / ``c:strCache``) living under the
        chart XML. When the workbook is edited by an external tool, only the
        embedded copy changes; the cached XML still holds the old values and
        keeps getting displayed until PowerPoint itself refreshes the chart
        (typically on open via an F5 / "refresh data" action). Users that
        round-trip a file through python-pptx therefore see stale values.

        This method re-reads the embedded workbook, looks up each cell
        reference recorded in a ``<c:f>`` formula element under the chart,
        and rewrites the sibling ``c:numCache`` / ``c:strCache`` so the
        cached values match the workbook again. Categories, series names,
        and numeric values are all refreshed. When a cell reference cannot
        be resolved — for example because the chart is linked to an
        external workbook that is not embedded, or the referenced sheet is
        absent — the corresponding cache is left untouched.

        This is a no-op for charts that do not have an embedded workbook
        (i.e. ``<c:externalData>`` is absent from the chart XML).
        """
        xlsx_part = self._workbook.xlsx_part
        if xlsx_part is None:
            return
        with WorkbookReader(xlsx_part.blob) as reader:
            _ChartCacheRefresher(self._chartSpace, reader).refresh()

    @lazyproperty
    def series(self):
        """
        A |SeriesCollection| object containing all the series in this
        chart. When the chart has multiple plots, all the series for the
        first plot appear before all those for the second, and so on. Series
        within a plot have an explicit ordering and appear in that sequence.
        """
        return SeriesCollection(self._chartSpace.plotArea)

    @property
    def value_axis(self):
        """
        The |ValueAxis| object providing access to properties of the value
        axis of this chart. Raises |ValueError| if the chart has no value
        axis.
        """
        valAx_lst = self._chartSpace.valAx_lst
        if not valAx_lst:
            raise ValueError("chart has no value axis")

        idx = 1 if len(valAx_lst) > 1 else 0
        return ValueAxis(valAx_lst[idx])

    @property
    def has_secondary_value_axis(self):
        """Read-only |bool| specifying whether this chart has a secondary value axis.

        Returns |True| when this chart has a second `c:valAx` element designating
        a secondary value axis (typically rendered on the right side of the chart
        in a category-based chart). Always |False| for an XY/scatter chart, where
        a second `c:valAx` identifies the X-axis rather than a secondary axis.
        """
        return self._chartSpace.plotArea.secondary_valAx is not None

    @property
    def secondary_value_axis(self):
        """The |ValueAxis| object for the secondary value axis of this chart.

        Raises |ValueError| if the chart has no secondary value axis. Use
        :attr:`has_secondary_value_axis` to test for its presence
        non-destructively.

        A secondary value axis is an additional value axis (conventionally
        rendered on the right side of the chart) against which one or more
        series may be plotted, allowing two data ranges on different scales to
        be compared on the same chart. An XY/scatter chart has two value axes
        but neither is considered "secondary" in this sense; both are primary.
        """
        secondary_valAx = self._chartSpace.plotArea.secondary_valAx
        if secondary_valAx is None:
            raise ValueError("chart has no secondary value axis")
        return ValueAxis(secondary_valAx)

    @property
    def _workbook(self):
        """
        The |ChartWorkbook| object providing access to the Excel source data
        for this chart.
        """
        return self.part.chart_workbook


class ChartTitle(ElementProxy):
    """Provides properties for manipulating a chart title."""

    # This shares functionality with AxisTitle, which could be factored out
    # into a base class, perhaps pptx.chart.shared.BaseTitle. I suspect they
    # actually differ in certain fuller behaviors, but at present they're
    # essentially identical.

    def __init__(self, title):
        super(ChartTitle, self).__init__(title)
        self._title = title

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line and fill formatting.

        Return the |ChartFormat| object providing shape formatting properties
        for this chart title, such as its line color and fill.
        """
        return ChartFormat(self._title)

    @property
    def has_text_frame(self):
        """Read/write Boolean specifying whether this title has a text frame.

        Return |True| if this chart title has a text frame, and |False|
        otherwise. Assigning |True| causes a text frame to be added if not
        already present. Assigning |False| causes any existing text frame to
        be removed along with its text and formatting.
        """
        if self._title.tx_rich is None:
            return False
        return True

    @has_text_frame.setter
    def has_text_frame(self, value):
        if bool(value) is False:
            self._title._remove_tx()
            return
        self._title.get_or_add_tx_rich()

    @property
    def text_frame(self):
        """|TextFrame| instance for this chart title.

        Return a |TextFrame| instance allowing read/write access to the text
        of this chart title and its text formatting properties. Accessing this
        property is destructive in the sense it adds a text frame if one is
        not present. Use :attr:`has_text_frame` to test for the presence of
        a text frame non-destructively.
        """
        rich = self._title.get_or_add_tx_rich()
        return TextFrame(rich, self)


class _ChartCacheRefresher(object):
    """Rewrite `c:numCache`/`c:strCache` trees under a `c:chartSpace` from xlsx.

    Iterates every `c:numRef` and `c:strRef` descendant of `chartSpace`,
    reads the sibling `c:f` formula reference, resolves the referenced
    cells in `reader`, and replaces the cache subtree with fresh
    `c:numCache` / `c:strCache` elements. A reference that cannot be
    resolved (unknown sheet, multi-range `f`, etc.) is left untouched so
    the chart still renders with whatever was there before.
    """

    def __init__(self, chartSpace, reader):
        self._chartSpace = chartSpace
        self._reader = reader

    def refresh(self):
        for numRef in self._chartSpace.iter(qn("c:numRef")):
            self._refresh_numRef(numRef)
        for strRef in self._chartSpace.iter(qn("c:strRef")):
            self._refresh_strRef(strRef)

    # -- numRef -----------------------------------------------------------

    def _refresh_numRef(self, numRef):
        values, format_code = self._resolve_numRef(numRef)
        if values is None:
            return
        self._replace_numCache(numRef, values, format_code)

    def _resolve_numRef(self, numRef):
        """Return `(values, format_code)` for `numRef` or `(None, None)`.

        `values` is a list whose elements are floats or |None| (for empty
        cells). `format_code` preserves any existing ``c:formatCode`` so the
        refresh doesn't silently drop number-format metadata.
        """
        f_elm = numRef.find(qn("c:f"))
        if f_elm is None or not f_elm.text:
            return None, None
        parsed = parse_sheet_range_ref(f_elm.text)
        if parsed is None:
            return None, None
        sheet_name, cells = parsed
        values = []
        for row, col in cells:
            v = self._reader.cell_value(sheet_name, row, col)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                values.append(float(v))
            elif v is None:
                values.append(None)
            else:
                # --- non-numeric value in a numRef; try to coerce, else keep None ---
                try:
                    values.append(float(v))
                except (TypeError, ValueError):
                    values.append(None)
        # --- preserve existing format code if present ---
        old_cache = numRef.find(qn("c:numCache"))
        format_code = None
        if old_cache is not None:
            fmt_elm = old_cache.find(qn("c:formatCode"))
            if fmt_elm is not None:
                format_code = fmt_elm.text
        return values, format_code

    def _replace_numCache(self, numRef, values, format_code):
        # --- remove existing numCache/numLit siblings (replace) ---
        for tag in ("c:numCache", "c:numLit"):
            for child in numRef.findall(qn(tag)):
                numRef.remove(child)
        cache = etree.SubElement(numRef, qn("c:numCache"))
        if format_code is not None:
            fmt = etree.SubElement(cache, qn("c:formatCode"))
            fmt.text = format_code
        ptCount = etree.SubElement(cache, qn("c:ptCount"))
        ptCount.set("val", str(len(values)))
        for idx, v in enumerate(values):
            if v is None:
                continue
            pt = etree.SubElement(cache, qn("c:pt"))
            pt.set("idx", str(idx))
            v_elm = etree.SubElement(pt, qn("c:v"))
            v_elm.text = _format_numeric(v)

    # -- strRef -----------------------------------------------------------

    def _refresh_strRef(self, strRef):
        values = self._resolve_strRef(strRef)
        if values is None:
            return
        self._replace_strCache(strRef, values)

    def _resolve_strRef(self, strRef):
        f_elm = strRef.find(qn("c:f"))
        if f_elm is None or not f_elm.text:
            return None
        parsed = parse_sheet_range_ref(f_elm.text)
        if parsed is None:
            return None
        sheet_name, cells = parsed
        values = []
        for row, col in cells:
            v = self._reader.cell_value(sheet_name, row, col)
            if v is None:
                values.append(None)
            elif isinstance(v, bool):
                values.append("TRUE" if v else "FALSE")
            elif isinstance(v, float):
                # --- format without trailing ".0" for whole numbers, to
                #     match PowerPoint's own strCache output ---
                if v.is_integer():
                    values.append(str(int(v)))
                else:
                    values.append(repr(v))
            else:
                values.append(str(v))
        return values

    def _replace_strCache(self, strRef, values):
        for tag in ("c:strCache", "c:strLit"):
            for child in strRef.findall(qn(tag)):
                strRef.remove(child)
        cache = etree.SubElement(strRef, qn("c:strCache"))
        ptCount = etree.SubElement(cache, qn("c:ptCount"))
        ptCount.set("val", str(len(values)))
        for idx, v in enumerate(values):
            if v is None:
                continue
            pt = etree.SubElement(cache, qn("c:pt"))
            pt.set("idx", str(idx))
            v_elm = etree.SubElement(pt, qn("c:v"))
            v_elm.text = v


def _format_numeric(value):
    """Serialize `value` to the form PowerPoint writes into `c:v`."""
    if value == 0:
        return "0"
    if float(value).is_integer():
        return str(int(value))
    return repr(float(value))


class _Plots(Sequence):
    """
    The sequence of plots in a chart, such as a bar plot or a line plot. Most
    charts have only a single plot. The concept is necessary when two chart
    types are displayed in a single set of axes, like a bar plot with
    a superimposed line plot.
    """

    def __init__(self, plotArea, chart):
        super(_Plots, self).__init__()
        self._plotArea = plotArea
        self._chart = chart

    def __getitem__(self, index):
        xCharts = self._plotArea.xCharts
        if isinstance(index, slice):
            plots = [PlotFactory(xChart, self._chart) for xChart in xCharts]
            return plots[index]
        else:
            xChart = xCharts[index]
            return PlotFactory(xChart, self._chart)

    def __len__(self):
        return len(self._plotArea.xCharts)
