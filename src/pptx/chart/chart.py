"""Chart-related objects such as Chart and ChartTitle."""

from __future__ import annotations

from collections.abc import Sequence

from lxml import etree

from pptx.chart.axis import CategoryAxis, DateAxis, ValueAxis
from pptx.chart.legend import Legend
from pptx.chart.plot import PlotFactory, PlotTypeInspector
from pptx.chart.series import SeriesCollection
from pptx.chart.xlsx import WorkbookReader, parse_sheet_range_ref
from pptx.chart.xmlwriter import SeriesXmlRewriterFactory
from pptx.dml.chtfmt import ChartFormat
from pptx.oxml.ns import qn
from pptx.shared import ElementProxy, PartElementProxy
from pptx.text.text import Font, TextFrame
from pptx.util import lazyproperty


class Chart(PartElementProxy):
    """A chart object."""

    def __init__(self, chartSpace, chart_part):
        super(Chart, self).__init__(chartSpace, chart_part)
        self._chartSpace = chartSpace

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
        """
        Read/write integer index of chart style used to format this chart.
        Range is from 1 to 48. Value is |None| if no explicit style has been
        assigned, in which case the default chart style is used. Assigning
        |None| causes any explicit setting to be removed. The integer index
        corresponds to the style's position in the chart style gallery in the
        PowerPoint UI.
        """
        style = self._chartSpace.style
        if style is None:
            return None
        return style.val

    @chart_style.setter
    def chart_style(self, value):
        self._chartSpace._remove_style()
        if value is None:
            return
        self._chartSpace._add_style(val=value)

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
        """
        Use the categories and series values in the |ChartData| object
        *chart_data* to replace those in the XML and Excel worksheet for this
        chart.
        """
        rewriter = SeriesXmlRewriterFactory(self.chart_type, chart_data)
        rewriter.replace_series_data(self._chartSpace)
        self._workbook.update_from_xlsx_blob(chart_data.xlsx_blob)

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
