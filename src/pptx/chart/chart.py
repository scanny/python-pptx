"""Chart-related objects such as Chart and ChartTitle."""

from __future__ import annotations

import io
import os
import zipfile
from collections.abc import Sequence
from copy import deepcopy

from lxml import etree

from pptx.chart.axis import CategoryAxis, DateAxis, ValueAxis
from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData
from pptx.chart.legend import Legend
from pptx.chart.plot import PlotFactory, PlotTypeInspector
from pptx.chart.series import SeriesCollection
from pptx.chart.xlsx import (
    WorkbookReader,
    WorkbookUpdater,
    parse_a1_cell,
    parse_sheet_range_ref,
)
from pptx.chart.xmlwriter import SeriesXmlRewriterFactory, _PlotFragmentBuilder
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_CHART_TYPE, XL_DISPLAY_BLANKS_AS
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
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

        .. versionadded:: 2026.05.0
        """
        plotArea = self._chartSpace.plotArea
        xCharts = plotArea.xCharts
        if not xCharts:
            raise ValueError("cannot add plot: chart has no existing plot")
        first_xChart = xCharts[0]
        axId_elms = first_xChart.findall(qn("c:axId"))
        if len(axId_elms) < 2:
            raise ValueError("cannot add plot: existing plot is missing axId elements")
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

    def apply_template(self, template):
        """Apply the chart formatting from a ``.crtx`` chart template.

        A ``.crtx`` chart-template file is a ZIP package containing a single
        chart ``c:chartSpace`` XML part (typically at ``chart/chart1.xml``).
        This method copies the *formatting* elements from that template onto
        this chart while leaving the chart's own *data* (series values,
        categories, embedded workbook) unchanged. It is the programmatic
        counterpart to PowerPoint's "Change Chart Type > Templates" UI
        command (GitHub issue #243).

        Parameters
        ----------
        template : str | bytes | file-like
            Path to a ``.crtx`` file, the bytes of one, or any file-like
            object (opened in binary mode) the :class:`zipfile.ZipFile`
            constructor accepts.

        What is copied
        --------------
        The following formatting elements are copied from the template's
        ``c:chartSpace`` onto this chart's ``c:chartSpace``:

        * ``c:style`` and/or the ``mc:AlternateContent`` wrapper holding a
          ``c14:style`` — the chart-style index (e.g. 118 for "Accent-N").
        * ``c:spPr`` — the chart-space fill / line / effect formatting.
        * ``c:txPr`` — the default text properties (font family, size,
          color) applied chart-wide.
        * On each axis in this chart's plot area, when the template has a
          same-type axis (``c:catAx`` ↔ ``c:catAx``, ``c:valAx`` ↔
          ``c:valAx``, ``c:dateAx`` ↔ ``c:dateAx``): the axis
          ``c:majorGridlines``, ``c:minorGridlines``, ``c:numFmt``,
          ``c:majorTickMark``, ``c:minorTickMark``, ``c:tickLblPos``,
          ``c:spPr`` and ``c:txPr`` sub-elements are replaced.
        * ``c:legend`` — when the template has a legend, its entire legend
          element (position, layout, and formatting) replaces this
          chart's legend. When the template has no legend, this chart's
          legend is left alone.

        What is preserved
        -----------------
        * All ``c:ser`` series elements (data, references, caches) — the
          whole point of templating is to re-skin *these* data.
        * The chart's embedded ``.xlsx`` workbook (``c:externalData``).
        * The chart's own ``c:title`` text; a template title element is
          copied only when this chart has no title. (A template's title
          text is usually a placeholder such as "Chart Title" — see the
          Notes section below.)
        * Axis ``c:axId`` values and scaling / cross references (so the
          target's plotted series remain connected to their axes).
        * The underlying plot type (``c:barChart``, ``c:lineChart``,
          etc.). This method does *not* change chart type; pair it with
          :meth:`~pptx.chart.chart.Chart.add_plot` or a fresh
          :meth:`~pptx.shapes.shapetree.SlideShapes.add_chart` when the
          template is for a different chart type than the target.

        Notes
        -----
        A ``.crtx`` is always a ZIP package even though PowerPoint shows it
        as a single file. Its embedded chart XML may reference cells in an
        embedded Excel workbook that is *not* included in the template
        (the cells serve only as placeholders for the author when the
        template was saved). This method ignores those references and any
        ``c:externalData`` in the template — only the formatting elements
        listed above cross over.

        Raises :class:`ValueError` when the template is not a valid ZIP
        package or does not contain a recognizable ``c:chartSpace`` XML
        part.

        .. versionadded:: 2026.05.0
        """
        template_chartSpace = _CrtxReader(template).chartSpace
        _ChartTemplateApplier(self._chartSpace, template_chartSpace).apply()

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

        Chart styles live in two tiers in the OOXML schema, and this property spans both:

        * **Plain styles (1-48)** — the `ST_Style` range defined in ISO/IEC 29500-1
          §21.2.3.46. These are serialised as a bare ``c:style val="N"`` child of
          ``c:chartSpace`` and correspond to the 48-slot chart-style gallery shipped
          with Office 2007 / PowerPoint 2010-2011. Assigning a value in 1..48 writes a
          plain ``c:style``.
        * **Extended styles (49-255)** — an Office 2010+ range carried in a
          ``c14:style`` element (namespace
          ``http://schemas.microsoft.com/office/drawing/2007/8/2/chart``). To stay
          backward-compatible the ``c14:style`` is wrapped in
          ``mc:AlternateContent`` with an ``mc:Choice Requires="c14"`` holding the
          extended value and an ``mc:Fallback`` holding a plain ``c:style`` so
          Office 2007-era readers still render the chart. Assigning a value in
          49..255 writes this wrapper; the fallback ``c:style`` is derived from
          the extended value (``value - 100`` when ``value > 100``, else ``value``,
          clamped to 1..48). PowerPoint emits values in this tier — for example
          118 = the "Accent-N" variant of plain style 18 — to keep each series in a
          pure accent colour for charts with six or more series (issue #516).

        Returns the extended (``c14:style``) value when an ``mc:AlternateContent``
        wrapper is present, otherwise the plain (``c:style``) value, otherwise
        ``None`` (no explicit style — PowerPoint applies its default).

        Assigning ``None`` removes any explicit setting, dropping both the bare
        ``c:style`` and any ``mc:AlternateContent`` wrapper.

        Examples of common PowerPoint-UI-displayed values in the plain tier
        (as typically seen in the Office 2007 / PowerPoint 2010-2011 48-slot
        gallery, which lays the styles out as six rows of eight slots, each row
        a different format level and each column a different accent slot):

        =====  =====================================================
        Value  Visual (approximate — see Note below)
        =====  =====================================================
        1      Monochrome (theme ``dk1`` / ``lt1``)
        2      Colourful, each series filled with a different accent
        3      Accent 1 shaded across series
        4      Accent 2 shaded across series
        5      Accent 3 shaded across series
        6      Accent 4 shaded across series
        7      Accent 5 shaded across series
        8      Accent 6 shaded across series
        42     Darker fills, Accent 2
        118    Extended: pure-Accent variant of plain style 18
        =====  =====================================================

        So, for example, ``chart.chart_style = 118`` writes the Office-2010+
        ``mc:AlternateContent`` / ``c14:style val="118"`` wrapper with a
        ``c:style val="18"`` fallback; ``chart.chart_style = 6`` writes a
        plain ``c:style val="6"``.

        Note
        ----
        The numeric-to-visual mapping **depends on the theme applied to the
        presentation** (the six ``a:accent1..a:accent6`` colours and the
        ``a:lt1`` / ``a:dk1`` pair from ``theme1.xml``): the style index selects
        a recipe, not fixed RGB values, so style 6 on a presentation themed in
        blues produces a different rendering than style 6 on a presentation
        themed in greens. The mapping also varies between PowerPoint versions
        (PowerPoint 2011 for Mac exposes all 48 plain slots, while PowerPoint
        2016+ groups them behind a shorter gallery whose entries depend on the
        chart type — see issue #407). When you need a specific visual, the most
        reliable workflow is to set the style in PowerPoint, then read back
        ``chart.chart_style`` (or inspect the chart part's XML) to learn the
        integer to use in code.
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

    def set_title(self, text):
        """Set the chart title text in a single call (GitHub issue #764).

        Convenience wrapper that ensures a title is present and rewrites
        its text, collapsing the common three-step idiom
        ``chart.has_title = True; chart.chart_title.text_frame.text = "…"``
        into one line. Assigning ``None`` (or an empty string) removes the
        chart title entirely, equivalent to setting
        :attr:`has_title` to ``False``.

        Returns this chart to support chaining, e.g.
        ``chart.set_title("Q4 Revenue").set_axis_title("value", "USD")``.

        .. versionadded:: 2026.05.1
        """
        if text is None or text == "":
            self.has_title = False
            return self
        self.chart_title.text_frame.text = text
        return self

    def set_axis_title(self, axis, text):
        """Set an axis's title text in a single call (GitHub issue #764).

        *axis* is one of the strings ``"category"``, ``"value"``, or
        ``"secondary_value"``, matching the corresponding
        :attr:`category_axis`, :attr:`value_axis`, and
        :attr:`secondary_value_axis` accessors on this chart. Assigning
        ``None`` (or an empty string) for *text* removes the axis title
        by setting the axis's ``has_title`` to ``False``; any non-empty
        string ensures the axis title element is present and rewrites its
        text. An unknown *axis* value raises :class:`ValueError`; a call
        with ``"secondary_value"`` on a chart without a secondary value
        axis propagates the :class:`ValueError` raised by
        :attr:`secondary_value_axis`.

        Returns this chart to support chaining.

        .. versionadded:: 2026.05.1
        """
        if axis == "category":
            axis_obj = self.category_axis
        elif axis == "value":
            axis_obj = self.value_axis
        elif axis == "secondary_value":
            axis_obj = self.secondary_value_axis
        else:
            raise ValueError(
                "axis must be one of 'category', 'value', 'secondary_value', " "got %r" % (axis,)
            )
        if text is None or text == "":
            axis_obj.has_title = False
            return self
        axis_obj.axis_title.text_frame.text = text
        return self

    @property
    def chart_type(self):
        """Member of :ref:`XlChartType` enumeration specifying type of this chart.

        If the chart has two plots, for example, a line plot overlayed on a bar plot,
        the type reported is for the first (back-most) plot. Read-only.
        """
        first_plot = self.plots[0]
        return PlotTypeInspector.chart_type(first_plot)

    @property
    def display_blanks_as(self):
        """Member of :ref:`XlDisplayBlanksAs` for blank-cell display treatment.

        Controls the PowerPoint "Hidden and Empty Cells" dialog option "Show
        empty cells as", which determines how gaps in the chart data are
        rendered — for example whether a stacked bar chart draws a
        zero-height bar for a missing value or omits it entirely. Implements
        GitHub issue #859, where zero values in a stacked bar chart are
        suppressed by setting this to :attr:`~XL_DISPLAY_BLANKS_AS.GAPS`.

        Returns :attr:`~XL_DISPLAY_BLANKS_AS.ZERO` when the chart has no
        ``c:dispBlanksAs`` child (the XSD-declared default). Assigning a
        non-default value writes a ``c:dispBlanksAs val="..."`` child;
        assigning :attr:`~XL_DISPLAY_BLANKS_AS.ZERO` (the default) removes
        any existing ``c:dispBlanksAs`` element so the XML stays minimal.
        Assigning a value that is not a member of :ref:`XlDisplayBlanksAs`
        raises :class:`ValueError`.
        """
        dispBlanksAs = self._chartSpace.chart.dispBlanksAs
        if dispBlanksAs is None:
            return XL_DISPLAY_BLANKS_AS.ZERO
        return dispBlanksAs.val

    @display_blanks_as.setter
    def display_blanks_as(self, value):
        XL_DISPLAY_BLANKS_AS.validate(value)
        chart = self._chartSpace.chart
        if value == XL_DISPLAY_BLANKS_AS.ZERO:
            chart._remove_dispBlanksAs()
            return
        dispBlanksAs = chart.get_or_add_dispBlanksAs()
        dispBlanksAs.val = value

    @lazyproperty
    def font(self):
        """Font object controlling text format defaults for this chart."""
        defRPr = self._chartSpace.get_or_add_txPr().p_lst[0].get_or_add_pPr().get_or_add_defRPr()
        return Font(defRPr)

    @property
    def has_data_table(self):
        """Read/write |bool| specifying whether a data table is shown beneath the chart.

        Assigning |True| adds a default ``c:plotArea/c:dTable`` element
        (with all four ``c:show*`` flags set to ``1``) if one is not
        already present. Assigning |False| removes any existing data
        table. Implements GitHub issue #373.

        .. versionadded:: 2026.05.0
        """
        return self._chartSpace.plotArea.dTable is not None

    @has_data_table.setter
    def has_data_table(self, value):
        plotArea = self._chartSpace.plotArea
        if bool(value) is False:
            plotArea._remove_dTable()
            return
        if plotArea.dTable is None:
            # -- create a default c:dTable (all four c:show* flags = 1) --
            # -- matching what PowerPoint writes when the user toggles --
            # -- "Data Table" on from the chart's Add Chart Element menu. --
            plotArea._insert_dTable(_new_default_dTable())

    @property
    def data_table(self):
        """A |_DataTable| proxy for the data table beneath this chart, or ``None``.

        Returns ``None`` when the chart has no ``c:plotArea/c:dTable``
        element. Use :attr:`has_data_table` to probe for presence
        non-destructively, and assign |True| to add a default data table.
        Implements GitHub issue #373.

        .. versionadded:: 2026.05.0
        """
        dTable = self._chartSpace.plotArea.dTable
        if dTable is None:
            return None
        return _DataTable(dTable)

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
        # -- GitHub issue #643: when re-enabling the title, clear any
        # -- prior `c:autoTitleDeleted val="1"` so PowerPoint's auto-title
        # -- logic engages (for a single-series chart, the title renders
        # -- as the series name; otherwise PowerPoint shows the "Chart Title"
        # -- placeholder). Leaving `autoTitleDeleted="1"` alongside an empty
        # -- `c:title` suppresses the auto-title entirely. --
        autoTitleDeleted = chart.autoTitleDeleted
        if autoTitleDeleted is not None and autoTitleDeleted.val is True:
            autoTitleDeleted.val = False

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
    def plot_area(self):
        """|PlotArea| instance providing access to plot-area formatting.

        The plot area is the rectangular region of a chart that contains the
        plotted data — bars, columns, lines, points, etc. — bounded by the
        chart's axes. This property returns a |PlotArea| proxy whose
        :attr:`~pptx.chart.chart.PlotArea.format` provides the
        :class:`~pptx.dml.chtfmt.ChartFormat` object that carries the plot
        area's fill, line, and effect formatting (``c:plotArea/c:spPr``).

        Implements issue #298. Mirrors the ``format`` pattern already in
        place on :class:`ChartTitle`, :class:`~pptx.chart.datalabel.DataLabel`,
        and :class:`~pptx.chart.datalabel.DataLabels`.

        .. versionadded:: 2026.05.0
        """
        return PlotArea(self._chartSpace.chart.plotArea)

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

        Author-set number formats on existing ``c:val`` / ``c:xVal`` / ``c:yVal`` /
        ``c:bubbleSize`` (and numeric ``c:cat``) elements are preserved across the
        rewrite so callers don't lose their ``0.00%`` / ``#,##0`` / etc. formatting
        on every ``replace_data`` round-trip (GitHub issue #666). Pass an explicit
        ``number_format=`` on *chart_data* to override the preserved format — the
        default ``"General"`` is treated as "no opinion" and the existing format is
        kept.

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

    def replace_data_preserve_formulas(self, chart_data):
        """Refresh data cells in the embedded workbook, skipping formulas (#239).

        A targeted-refresh counterpart to :meth:`replace_data`: instead of
        re-authoring the workbook from scratch (which discards any
        author-entered formulas), this method walks the cells that
        *chart_data* would write and updates each one individually via
        :func:`update_embedded_xlsx_cell` — but only for cells that do
        **not** carry an ``<f>`` formula element in the existing workbook.
        Formula cells are left alone so their computed values continue to
        drive the chart after the refresh.

        *chart_data* must be a |ChartData| subtype compatible with this
        chart's type (same validation as :meth:`replace_data`); see that
        method's docstring and GitHub issue #396 for the compatibility
        matrix. Raises |ValueError| when the chart has no embedded
        workbook (``c:externalData`` missing or its relationship
        stripped — see issue #490).

        Limitations
        -----------
        This is a *data-only refresh*, not a structural rewrite:

        * Chart dimensions are fixed by the existing ``c:ser`` / ``c:cat``
          / ``c:val`` cell ranges. If *chart_data* carries **more** series
          or categories than the chart currently has, the excess cells are
          written to the workbook but no new ``c:ser`` elements are added
          to the chart XML — the chart will not render them until the
          workbook is opened in Excel / PowerPoint and the chart refreshes
          its range. If *chart_data* carries **fewer** series or
          categories than the chart has, the surplus cells in the workbook
          are left at their previous values (they are not cleared). Use
          :meth:`replace_data` when you need to change chart dimensions.
        * Cached chart values (``c:numCache`` / ``c:strCache``) are
          refreshed only for cells that were written. Cells skipped
          because they carry a formula retain their existing cache entry
          — call :meth:`update_cached_values` afterward to reconcile
          caches against the post-refresh formula results.

        Returns the number of cells that were written (i.e. excluding
        formula cells that were skipped).

        .. versionadded:: 2026.05.0
        """
        self._validate_chart_data_type(chart_data)
        workbook_bytes = self.workbook
        if workbook_bytes is None:
            raise ValueError(
                "chart has no embedded workbook to refresh; use "
                "Chart.replace_data() to create one"
            )
        # --- inspect the current workbook to find formula cells that must
        # --- be skipped. Use a read-only WorkbookReader so we touch each
        # --- sheet at most once regardless of how many cells the writer
        # --- enumerates. --
        writer = chart_data._workbook_writer
        with WorkbookReader(workbook_bytes) as reader:
            writes = []
            for sheet, row, col, value in writer.iter_cell_writes():
                if reader.cell_has_formula(sheet, row, col):
                    continue
                writes.append((sheet, row, col, value))
        # --- apply non-formula writes to the workbook blob in one pass,
        # --- then refresh matching chart caches for each written cell. --
        updater = WorkbookUpdater(workbook_bytes)
        for sheet, row, col, value in writes:
            updater.set_cell(sheet, row, col, value)
        new_blob = updater.blob()
        self._workbook.update_from_xlsx_blob(new_blob)
        for sheet, row, col, value in writes:
            _SingleCellCacheRefresher(self._chartSpace, sheet, row, col, value).refresh()
        return len(writes)

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
            if not isinstance(chart_data, XyChartData) or isinstance(chart_data, BubbleChartData):
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

        .. versionadded:: 2026.05.0
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
    def series_in_rows(self):
        """Read-only |bool| indicating whether series data are in rows.

        Corresponds to the "Switch Row/Column" state of the chart's source
        data in PowerPoint's Edit Data dialog (GitHub issue #828). OOXML does
        not persist the Switch Row/Column flag directly — the orientation is
        implicit in the shape of the cell references under each ``c:ser``.
        This property inspects those references and reports:

        * ``False`` — series are laid out in **columns** of the embedded
          worksheet. This is PowerPoint's default (and the layout
          python-pptx writes when authoring a chart from
          :class:`CategoryChartData`): each series name lives in row 1 of
          its own column and the category labels share a single column to
          its left.
        * ``True`` — series are laid out in **rows**. Each series name is
          one cell of a single column (typically column A, rows 2..N) and
          the category labels span row 1 across multiple columns. This is
          the state a user who applied *Chart Design > Switch Row/Column*
          in PowerPoint would produce.
        * ``None`` — the orientation cannot be determined from the chart
          XML. Common causes: the chart has no series, the first series
          has no ``c:cat`` reference (e.g. an XY/scatter or bubble chart,
          which have no category axis), or the references are inline
          literals / unparseable. Callers should treat ``None`` as
          "unknown" rather than "default".

        Detection uses the first series's ``c:cat//c:f`` range — a
        single-column range means series-in-columns (``False``); a
        single-row range means series-in-rows (``True``). A single-cell
        ``c:cat`` (one category) falls back to the series-name ref:
        ``c:tx/c:strRef/c:f`` at row 1 ⇒ ``False``; elsewhere ⇒ ``True``.
        Read-only: python-pptx does not offer a writer for the switch
        because PowerPoint rewrites every ``c:ser`` sub-reference in
        response to the UI toggle, and emulating that without a live
        workbook would silently mismatch what PowerPoint re-authors on
        next save. Users who need to flip orientation programmatically
        should re-author the chart via :meth:`replace_data` with the
        desired data transposed.
        """
        plotArea = self._chartSpace.plotArea
        # -- `c:ser` is a descendant (e.g. `c:barChart/c:ser`), not a direct
        # -- child of `c:plotArea`; use `.//` to reach any chart-type. --
        first_ser = plotArea.find(".//" + qn("c:ser"))
        if first_ser is None:
            return None
        # --- examine the c:cat/c:strRef (or numRef) first — its range ---
        # --- shape is the strongest signal for series orientation. --
        cat = first_ser.find(qn("c:cat"))
        if cat is not None:
            parsed = _parse_first_ref(cat)
            if parsed is not None:
                cells = parsed
                rows = {r for r, _ in cells}
                cols = {c for _, c in cells}
                # -- a multi-row, single-column range means categories run
                # -- vertically — series are in columns (default). --
                if len(rows) > 1 and len(cols) == 1:
                    return False
                # -- a single-row, multi-column range means categories run
                # -- horizontally — series are in rows (switched). --
                if len(cols) > 1 and len(rows) == 1:
                    return True
                # -- single-cell c:cat range: inconclusive from c:cat alone;
                # -- fall through to examine the series-name ref. --
        # --- fall back to c:tx/c:strRef/c:f; row 1 => default, else switched ---
        tx = first_ser.find(qn("c:tx"))
        if tx is not None:
            parsed = _parse_first_ref(tx)
            if parsed is not None:
                rows = {r for r, _ in parsed}
                cols = {c for _, c in parsed}
                if len(rows) == 1 and len(cols) == 1:
                    row, _col = next(iter(parsed))
                    # -- series name in row 1 ⇒ default (series in cols);
                    # -- series name elsewhere on a later row ⇒ switched
                    # -- (series in rows). --
                    return row != 1
        return None

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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
        """
        secondary_valAx = self._chartSpace.plotArea.secondary_valAx
        if secondary_valAx is None:
            raise ValueError("chart has no secondary value axis")
        return ValueAxis(secondary_valAx)

    def clone_to(self, shapes, x, y, cx, cy):
        """Duplicate this chart onto `shapes` at (`x`, `y`) with size (`cx`, `cy`).

        Implements the cross-slide chart-copy primitive (issue #877). `shapes`
        is a slide shape-tree such as :attr:`Slide.shapes` on the same
        presentation as this chart *or* on a different one. A new chart part
        is created in `shapes`'s presentation containing a deep copy of this
        chart's ``c:chartSpace`` XML and a freshly-duplicated
        :class:`EmbeddedXlsxPart`, and a new ``p:graphicFrame`` referencing it
        is appended to `shapes`. Every relationship carried by the source
        chart part (chart image, theme override, ...) is re-established on
        the new chart part as a side-effect — no cross-package or cross-part
        references escape.

        In the same-presentation case (copy from slide A to slide B within
        one file), image/theme-override parts are reused by package
        reference; only the chart part itself and the embedded workbook are
        forcibly duplicated so PowerPoint's "Edit Data" dialog keeps working
        per-chart. In the cross-presentation case every referenced part is
        materialised in the destination package so the target file can be
        saved and opened independently.

        `x`, `y`, `cx`, `cy` are the position and extents of the new chart's
        ``p:graphicFrame`` on `shapes`'s slide, in English Metric Units. They
        are required: a chart's current position lives on its enclosing
        ``p:graphicFrame`` (not on the chart XML itself), so callers that
        want to preserve the source position should read those values from
        the ``GraphicFrame`` shape wrapping the source chart (e.g.
        ``source_gf.left``, ``source_gf.top``, ``source_gf.width``,
        ``source_gf.height``).

        Returns the newly-created :class:`~pptx.shapes.graphfrm.GraphicFrame`
        containing the duplicated chart; reach the duplicate :class:`Chart`
        via the returned graphic-frame's
        :attr:`~pptx.shapes.graphfrm.GraphicFrame.chart` property.

        .. versionadded:: 2026.05.0
        """
        return shapes.clone_chart(self, x, y, cx, cy)

    @property
    def has_user_shapes(self):
        """Read-only |bool| specifying whether this chart has any user-shape annotations.

        ``True`` when the chart has a ``chartUserShapes`` relationship (the
        ``c:userShapes`` element) carrying one or more
        ``cdr:relSizeAnchor`` / ``cdr:absSizeAnchor`` annotation anchors.
        ``False`` when the relationship is missing or the related part
        contains no anchors. Use this to probe for annotation shapes
        non-destructively before touching :attr:`user_shapes`.

        MVP scope (issue #351) is read-only inspection; see
        :attr:`user_shapes` and the
        ``docs/dev/analysis/chart-user-shapes.rst`` analysis for context.

        .. versionadded:: 2026.05.0
        """
        user_shapes = self.user_shapes
        if user_shapes is None:
            return False
        return user_shapes.anchor_count > 0

    @property
    def user_shapes(self):
        """The |ChartDrawingPart| carrying this chart's user-shape annotations, or ``None``.

        A user-shapes drawing overlays annotation shapes (arrows, call-outs,
        text boxes) onto the chart plot area — PowerPoint authors them via
        *Insert > Shapes* while the chart is selected, and stores them in
        a separate ``cdr:userShapes`` XML part linked from the chart via a
        ``chartUserShapes`` relationship.

        This MVP exposes read-only access: the returned part offers
        :meth:`~pptx.parts.chartdrawing.ChartDrawingPart.iter_anchor_elements`
        and :attr:`~pptx.parts.chartdrawing.ChartDrawingPart.anchor_count`
        for enumerating the raw anchor elements. Authoring annotation
        shapes is not yet supported; see
        ``docs/dev/analysis/chart-user-shapes.rst`` for the full
        design note and the reason the proxy hierarchy is deferred.

        Returns ``None`` when the chart has no ``chartUserShapes``
        relationship (the common case for charts authored through this
        library, since python-pptx does not add one).

        .. versionadded:: 2026.05.0
        """
        try:
            return self.part.part_related_by(RT.CHART_USER_SHAPES)
        except KeyError:
            return None

    @property
    def workbook(self):
        """Bytes of this chart's embedded Excel workbook, or ``None``.

        Reads the raw bytes of the ``.xlsx`` part referenced by this chart's
        ``c:externalData`` element. Returns ``None`` when the chart has no
        embedded workbook — for example when it is linked to an external
        Excel file via ``c:externalData/@externalDataId`` without an
        embedded copy, or when the embedded-workbook relationship is
        missing (e.g. a chart pasted from a pre-2007 ``.xls`` source; see
        issue #490).

        Assigning a bytes object replaces the embedded workbook in-place —
        the existing :class:`EmbeddedXlsxPart` keeps its part name and
        relationship id, only its ``blob`` is overwritten. If the chart
        does not yet have an embedded workbook, a new
        :class:`EmbeddedXlsxPart` is created and attached via a ``PACKAGE``
        relationship, and a ``c:externalData`` element pointing at it is
        added to the chart XML. Setting ``None`` is not supported (raise
        :class:`TypeError`). This accessor is the F5 cross-part handler
        used by combo-chart, cross-slide chart copy, and targeted cell
        updates via :func:`update_embedded_xlsx_cell`.
        """
        xlsx_part = self._workbook.xlsx_part
        if xlsx_part is None:
            return None
        return xlsx_part.blob

    @workbook.setter
    def workbook(self, xlsx_blob):
        if not isinstance(xlsx_blob, (bytes, bytearray)):
            raise TypeError(
                "Chart.workbook must be set to a bytes object, got %s" % type(xlsx_blob).__name__
            )
        self._workbook.update_from_xlsx_blob(bytes(xlsx_blob))

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

    @property
    def position(self):
        """Read/write ``(x, y)`` tuple of fractional manual-layout offsets.

        Returns ``None`` when the title is laid out automatically (the
        PowerPoint default). When non-``None``, the tuple contains two
        ``float`` values in the 0.0-1.0 relative-coordinate space defined by
        ``c:title/c:layout/c:manualLayout/c:x`` and ``c:y``: ``(0.0, 0.0)``
        pins the top-left of the title to the top-left of the chart, and
        ``(1.0, 1.0)`` to the bottom-right; values outside that range are
        accepted and render as off-chart positions just as PowerPoint allows.

        Assigning a 2-tuple writes ``c:layout/c:manualLayout`` with both
        ``c:xMode`` and ``c:yMode`` set to ``"factor"`` (the relative-
        coordinate mode PowerPoint writes when a user drags the title). This
        pins the title to the manual position — PowerPoint will not
        automatically reposition it on subsequent edits. Assigning ``None``
        removes the ``c:layout`` entirely, restoring auto-layout. Assigning
        a non-2-tuple raises :class:`ValueError`.

        Implements GitHub issue #1030.
        """
        layout = self._title.layout
        if layout is None:
            return None
        manualLayout = layout.manualLayout
        if manualLayout is None:
            return None
        return manualLayout.position

    @position.setter
    def position(self, value):
        if value is None:
            self._title._remove_layout()
            return
        try:
            x, y = value
        except (TypeError, ValueError):
            raise ValueError(
                "ChartTitle.position must be a 2-tuple (x, y) or None, got %r" % (value,)
            )
        layout = self._title.get_or_add_layout()
        manualLayout = layout.get_or_add_manualLayout()
        manualLayout.position = (x, y)


class PlotArea(ElementProxy):
    """Proxy for a chart's ``c:plotArea`` element.

    Provides access to the plot-area's shape-formatting block (fill, line,
    and shadow) via its :attr:`format` property. The plot area is the
    rectangular region of a chart that contains the plotted data, bounded
    by the chart's axes.

    Access via :attr:`Chart.plot_area` (issue #298).

    .. versionadded:: 2026.05.0
    """

    def __init__(self, plotArea):
        super(PlotArea, self).__init__(plotArea)
        self._plotArea = plotArea

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line and fill formatting.

        Returns the |ChartFormat| proxy for this plot area's ``c:spPr``
        shape-properties element, giving read/write access to its
        :attr:`~pptx.dml.chtfmt.ChartFormat.fill`,
        :attr:`~pptx.dml.chtfmt.ChartFormat.line`, and
        :attr:`~pptx.dml.chtfmt.ChartFormat.shadow`.

        .. versionadded:: 2026.05.0
        """
        return ChartFormat(self._plotArea)


class _DataTable(ElementProxy):
    """Proxy for a chart's ``c:dTable`` (data-table) element.

    Exposes the four ``c:show*`` border/outline/keys booleans as read/write
    properties, and a :attr:`format` property returning a |ChartFormat| for
    the data table's ``c:spPr`` shape-properties child (fill, line, and
    effect formatting). Implements GitHub issue #373.

    Access via :attr:`Chart.data_table` (returns ``None`` when no data
    table is present; pair with :attr:`Chart.has_data_table` for a
    non-destructive probe, and assign ``Chart.has_data_table = True`` to
    add a default data table).

    .. versionadded:: 2026.05.0
    """

    def __init__(self, dTable):
        super(_DataTable, self).__init__(dTable)
        self._dTable = dTable

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line and fill formatting.

        Returns the |ChartFormat| proxy for this data table's ``c:spPr``
        shape-properties element, giving read/write access to its
        :attr:`~pptx.dml.chtfmt.ChartFormat.fill`,
        :attr:`~pptx.dml.chtfmt.ChartFormat.line`, and
        :attr:`~pptx.dml.chtfmt.ChartFormat.shadow`.

        .. versionadded:: 2026.05.0
        """
        return ChartFormat(self._dTable)

    @property
    def show_horz_border(self):
        """Read/write |bool| specifying whether horizontal borders are shown.

        Corresponds to the ``c:showHorzBorder`` child. Returns |True| when
        the child is absent (the schema-declared default). Assigning |None|
        removes the child (restoring the default).
        """
        return self._show_val("showHorzBorder")

    @show_horz_border.setter
    def show_horz_border(self, value):
        self._set_show_val("showHorzBorder", value)

    @property
    def show_vert_border(self):
        """Read/write |bool| specifying whether vertical borders are shown.

        Corresponds to the ``c:showVertBorder`` child. Returns |True| when
        the child is absent (the schema-declared default). Assigning |None|
        removes the child (restoring the default).
        """
        return self._show_val("showVertBorder")

    @show_vert_border.setter
    def show_vert_border(self, value):
        self._set_show_val("showVertBorder", value)

    @property
    def show_outline(self):
        """Read/write |bool| specifying whether the data-table outline is shown.

        Corresponds to the ``c:showOutline`` child. Returns |True| when the
        child is absent (the schema-declared default). Assigning |None|
        removes the child (restoring the default).
        """
        return self._show_val("showOutline")

    @show_outline.setter
    def show_outline(self, value):
        self._set_show_val("showOutline", value)

    @property
    def show_keys(self):
        """Read/write |bool| specifying whether legend keys are shown in the data table.

        Corresponds to the ``c:showKeys`` child. Returns |True| when the
        child is absent (the schema-declared default). Assigning |None|
        removes the child (restoring the default).
        """
        return self._show_val("showKeys")

    @show_keys.setter
    def show_keys(self, value):
        self._set_show_val("showKeys", value)

    def _show_val(self, attr_name):
        """Return the effective |bool| value of child ``c:<attr_name>``.

        Missing child resolves to |True| per ``CT_Boolean``'s default.
        """
        child = getattr(self._dTable, attr_name)
        if child is None:
            return True
        return bool(child.val)

    def _set_show_val(self, attr_name, value):
        if value is None:
            getattr(self._dTable, "_remove_%s" % attr_name)()
            return
        child = getattr(self._dTable, "get_or_add_%s" % attr_name)()
        child.val = bool(value)


def _new_default_dTable():
    """Return a newly-created default ``c:dTable`` element.

    Thin wrapper over :meth:`CT_DTable.new_dTable` that keeps the
    `Chart.has_data_table` setter free of an `oxml`-module import at the
    top of this file (there's already a ripple of ``from pptx.oxml …``
    lines above; one more helper function keeps the seam local).
    """
    from pptx.oxml.chart.chart import CT_DTable

    return CT_DTable.new_dTable()


def update_embedded_xlsx_cell(chart, sheet, a1_ref, value):
    """Write `value` to cell `a1_ref` on `sheet` of `chart`'s embedded workbook.

    This is the F5 targeted-cell helper: it rewrites the single cell
    ``sheet!a1_ref`` in the chart's embedded ``.xlsx`` workbook *and*
    synchronously rewrites every ``c:numCache`` / ``c:strCache`` entry in
    the chart XML that points at that cell — both sides are updated as a
    single transaction so the chart never renders with stale cached values
    after a cell edit.

    Parameters
    ----------
    chart : :class:`Chart`
        The chart whose embedded workbook will be updated.
    sheet : str
        The worksheet name (e.g. ``"Sheet1"``). Chart workbooks
        authored by python-pptx always use ``"Sheet1"``; files authored
        elsewhere may have been renamed. When `sheet` is not found in
        the workbook, the updater falls back to the first sheet — this
        matches PowerPoint's own behaviour for chart-data edits.
    a1_ref : str
        A single-cell A1-style reference (e.g. ``"B2"`` or ``"$B$2"``).
        Multi-cell ranges (``"A1:B2"``) are not supported; use several
        calls instead. A sheet-qualified reference (``"Sheet1!B2"``)
        raises :class:`ValueError`.
    value : None | bool | int | float | str
        The new cell value. ``None`` clears the cell. Numeric values
        become numeric cells; strings become inline strings
        (``t="inlineStr"``); booleans become ``t="b"`` cells.

    Behaviour
    ---------
    * When the chart has no embedded workbook (``c:externalData`` missing
      or its rel stripped; see issue #490), this function raises
      :class:`ValueError` rather than silently succeeding — callers can
      use :attr:`Chart.workbook` first to probe presence.
    * The function does **not** re-author the entire workbook. Only the
      bytes of the single cell change; styles, shared strings, other
      cells, and docProps are preserved byte-for-byte modulo any XML
      re-serialization of the target worksheet.
    * After the blob is rewritten, each chart ``c:numRef``/``c:strRef``
      whose ``c:f`` resolves to this cell has its sibling cache
      replaced with a single-point ``c:numCache`` or ``c:strCache`` so
      the chart renders the new value immediately, without requiring a
      ``Chart.update_cached_values()`` pass.

    Returns the updated xlsx blob bytes (also now stored on the chart).

    .. versionadded:: 2026.05.0
    """
    workbook_bytes = chart.workbook
    if workbook_bytes is None:
        raise ValueError("chart has no embedded workbook to update")
    if "!" in a1_ref:
        raise ValueError(
            "a1_ref must be a single-cell reference without sheet qualifier, " "got %r" % a1_ref
        )
    row, col = parse_a1_cell(a1_ref)
    updater = WorkbookUpdater(workbook_bytes)
    updater.set_cell(sheet, row, col, value)
    new_blob = updater.blob()
    chart.workbook = new_blob
    # --- refresh any caches that point at this cell ---
    _SingleCellCacheRefresher(chart._chartSpace, sheet, row, col, value).refresh()
    return new_blob


class _SingleCellCacheRefresher(object):
    """Rewrite ``c:numCache`` / ``c:strCache`` entries targeting one cell.

    Scans every ``c:numRef`` / ``c:strRef`` descendant of a ``c:chartSpace``
    and, when its ``c:f`` resolves to (sheet, row, col), overwrites the
    corresponding entry in the sibling cache with `value`. Ranges that
    *include* the target cell have only the one matching ``c:pt`` updated;
    other points in the range are left alone so that a cell-write doesn't
    destructively truncate a multi-cell cache.
    """

    def __init__(self, chartSpace, sheet, row, col, value):
        self._chartSpace = chartSpace
        self._sheet = sheet
        self._row = row
        self._col = col
        self._value = value

    def refresh(self):
        for numRef in self._chartSpace.iter(qn("c:numRef")):
            self._update_numRef(numRef)
        for strRef in self._chartSpace.iter(qn("c:strRef")):
            self._update_strRef(strRef)

    def _matching_index(self, ref_elm):
        """Return the point index within `ref_elm` that targets the cell.

        Returns ``None`` when the cell is not covered by this ref's range
        or the range can't be parsed. Returns an int index (0-based into
        the range's cell list, which is row-major) otherwise.
        """
        f_elm = ref_elm.find(qn("c:f"))
        if f_elm is None or not f_elm.text:
            return None
        parsed = parse_sheet_range_ref(f_elm.text)
        if parsed is None:
            return None
        ref_sheet, cells = parsed
        # --- compare sheet names case-sensitively; PowerPoint preserves
        # --- case. Fall back to first-sheet semantics only when updater
        # --- couldn't find the named sheet either. --
        if ref_sheet != self._sheet:
            return None
        for idx, (r, c) in enumerate(cells):
            if r == self._row and c == self._col:
                return idx
        return None

    def _update_numRef(self, numRef):
        idx = self._matching_index(numRef)
        if idx is None:
            return
        numeric = self._as_number(self._value)
        self._set_cache_pt(numRef, "c:numCache", "c:numLit", idx, numeric)

    def _update_strRef(self, strRef):
        idx = self._matching_index(strRef)
        if idx is None:
            return
        text = self._as_text(self._value)
        self._set_cache_pt(strRef, "c:strCache", "c:strLit", idx, text)

    @staticmethod
    def _as_number(value):
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _as_text(value):
        if value is None:
            return None
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def _set_cache_pt(self, ref_elm, cache_tag, lit_tag, idx, value):
        """Replace ``c:pt[@idx=idx]/c:v`` in the ref's cache with `value`.

        Creates the cache + ``c:pt`` entries if they don't exist.
        Removes any sibling ``c:numLit`` / ``c:strLit`` since those would
        otherwise take precedence over the cache for some readers.
        """
        # --- drop any literal counterpart; cache is the authoritative form ---
        for child in ref_elm.findall(qn(lit_tag)):
            ref_elm.remove(child)
        cache = ref_elm.find(qn(cache_tag))
        if cache is None:
            cache = etree.SubElement(ref_elm, qn(cache_tag))
            # --- ptCount is required by the schema; keep it at least idx+1 ---
            ptCount = etree.SubElement(cache, qn("c:ptCount"))
            ptCount.set("val", str(idx + 1))
        else:
            ptCount = cache.find(qn("c:ptCount"))
            if ptCount is None:
                ptCount = etree.SubElement(cache, qn("c:ptCount"))
                ptCount.set("val", str(idx + 1))
            else:
                try:
                    cur = int(ptCount.get("val", "0"))
                except ValueError:
                    cur = 0
                if cur < idx + 1:
                    ptCount.set("val", str(idx + 1))
        pt = None
        for existing in cache.findall(qn("c:pt")):
            try:
                if int(existing.get("idx", "-1")) == idx:
                    pt = existing
                    break
            except ValueError:
                continue
        if value is None:
            if pt is not None:
                cache.remove(pt)
            return
        if pt is None:
            pt = etree.SubElement(cache, qn("c:pt"))
            pt.set("idx", str(idx))
        # --- rewrite the c:v child in place ---
        for child in list(pt):
            pt.remove(child)
        v_elm = etree.SubElement(pt, qn("c:v"))
        if cache_tag == "c:numCache":
            v_elm.text = _format_numeric(value)
        else:
            v_elm.text = value


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


def _parse_first_ref(parent):
    """Return the parsed cell list for the first ``c:f`` under *parent*.

    *parent* is expected to be a ``c:cat``, ``c:val``, ``c:tx``,
    ``c:strRef``, or ``c:numRef`` element. Returns the list of
    ``(row, col)`` tuples parsed by :func:`parse_sheet_range_ref`, or
    ``None`` when no parseable reference is found. The sheet name is
    discarded — callers of this helper only care about the range
    geometry for orientation detection (issue #828).
    """
    f_elm = parent.find(".//" + qn("c:f"))
    if f_elm is None or not f_elm.text:
        return None
    parsed = parse_sheet_range_ref(f_elm.text)
    if parsed is None:
        return None
    _sheet, cells = parsed
    return cells


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


class _CrtxReader(object):
    """Open a ``.crtx`` chart-template package and expose its ``c:chartSpace``.

    A ``.crtx`` is a ZIP package; its chart XML typically lives at
    ``chart/chart1.xml`` but we scan the archive for any entry whose root
    element is ``c:chartSpace`` to be robust against future layouts.
    """

    def __init__(self, template):
        self._template = template

    @property
    def chartSpace(self):
        """Return the ``c:chartSpace`` element parsed from the template."""
        raw = self._open_zip()
        try:
            with zipfile.ZipFile(raw) as zf:
                for name in zf.namelist():
                    # -- ignore rels, content-types, theme, etc. --
                    if not name.lower().endswith(".xml"):
                        continue
                    if "_rels" in name.lower() or name.lower().endswith(".rels"):
                        continue
                    blob = zf.read(name)
                    if b"chartSpace" not in blob:
                        continue
                    try:
                        elm = parse_xml(blob)
                    except etree.XMLSyntaxError:
                        continue
                    if elm.tag == qn("c:chartSpace"):
                        return elm
        except zipfile.BadZipFile as exc:
            raise ValueError("apply_template: template is not a valid .crtx ZIP package: %s" % exc)
        raise ValueError("apply_template: no c:chartSpace XML part found in template")

    def _open_zip(self):
        """Return a binary file-like opened on the template data."""
        tmpl = self._template
        if isinstance(tmpl, (bytes, bytearray)):
            return io.BytesIO(bytes(tmpl))
        if isinstance(tmpl, (str, os.PathLike)):
            return open(tmpl, "rb")
        # -- assume an already-open file-like object --
        return tmpl


class _ChartTemplateApplier(object):
    """Copy formatting elements from a template ``c:chartSpace`` onto a target.

    Data (c:ser, c:externalData) on the target is left untouched; axis IDs
    and cross-references are preserved so the copy cannot break the chart's
    own series-to-axis plumbing. See :meth:`Chart.apply_template` for the
    full list of what is and isn't copied.
    """

    # -- chartSpace-level children that carry only formatting --
    _CHARTSPACE_FMT_TAGS = ("c:style", "c:spPr", "c:txPr")

    # -- axis-level children that carry only formatting (no axId / scaling / --
    # -- cross references, so copying them cannot break axis plumbing) --
    _AXIS_FMT_TAGS = (
        "c:majorGridlines",
        "c:minorGridlines",
        "c:numFmt",
        "c:majorTickMark",
        "c:minorTickMark",
        "c:tickLblPos",
        "c:spPr",
        "c:txPr",
    )

    _AXIS_TAGS = ("c:catAx", "c:dateAx", "c:valAx", "c:serAx")

    def __init__(self, target_chartSpace, template_chartSpace):
        self._target = target_chartSpace
        self._template = template_chartSpace

    def apply(self):
        self._apply_style()
        self._apply_chartSpace_formatting()
        self._apply_axes_formatting()
        self._apply_legend()
        self._apply_title_if_absent()

    # -- chart-style index (c:style or mc:AlternateContent/c14:style) --

    def _apply_style(self):
        """Copy the chart-style index from the template to the target.

        Uses :meth:`CT_ChartSpace.set_chart_style_ex_val` which already
        handles the plain-vs-extended split (plain ``c:style`` 1-48 vs.
        ``mc:AlternateContent/c14:style`` 49-255).
        """
        # -- read template's effective style value (extended preferred) --
        style_val = self._template.chart_style_ex_val
        if style_val is None:
            return
        self._target.set_chart_style_ex_val(style_val)

    # -- chartSpace-level c:spPr / c:txPr --

    def _apply_chartSpace_formatting(self):
        """Replace target's ``c:spPr`` / ``c:txPr`` with the template's copies."""
        for tag in ("c:spPr", "c:txPr"):
            template_elm = self._template.find(qn(tag))
            # -- remove any existing child of this tag on the target --
            for existing in self._target.findall(qn(tag)):
                self._target.remove(existing)
            if template_elm is None:
                continue
            self._insert_in_schema_order(self._target, deepcopy(template_elm), tag)

    # -- axis formatting (per axis type) --

    def _apply_axes_formatting(self):
        """For each target axis with a matching template axis of the same type,
        replace the target axis's formatting sub-elements with the template's.

        "Matching" means same tag name and same ordinal position among
        axes of that type (1st valAx ↔ 1st valAx, 2nd valAx ↔ 2nd valAx).
        """
        target_pa = self._target.find(qn("c:chart") + "/" + qn("c:plotArea"))
        template_pa = self._template.find(qn("c:chart") + "/" + qn("c:plotArea"))
        if target_pa is None or template_pa is None:
            return
        for axis_tag in self._AXIS_TAGS:
            target_axes = target_pa.findall(qn(axis_tag))
            template_axes = template_pa.findall(qn(axis_tag))
            for target_axis, template_axis in zip(target_axes, template_axes):
                self._apply_axis_formatting(target_axis, template_axis)

    def _apply_axis_formatting(self, target_axis, template_axis):
        """Copy each formatting child of `template_axis` onto `target_axis`.

        The target's existing same-tag child is replaced. Non-formatting
        children (axId, scaling, delete, axPos, crosses, crossesAt, etc.)
        on the target are untouched.
        """
        for tag in self._AXIS_FMT_TAGS:
            template_child = template_axis.find(qn(tag))
            # -- remove any existing child on the target --
            for existing in target_axis.findall(qn(tag)):
                target_axis.remove(existing)
            if template_child is None:
                continue
            # -- use ZeroOrOne accessor to respect schema order when possible --
            attr = tag.split(":", 1)[1]
            # -- some schema attribute names are python keywords (delete_) --
            # -- but none of the _AXIS_FMT_TAGS are; direct name works. --
            if hasattr(target_axis, attr):
                # -- the ZeroOrOne descriptor's setter handles insertion --
                # -- but to keep the copy deep we set via direct insert --
                self._insert_axis_child(target_axis, deepcopy(template_child))
            else:
                # -- fall back to plain append; caller can re-serialize --
                target_axis.append(deepcopy(template_child))

    def _insert_axis_child(self, target_axis, new_elm):
        """Insert `new_elm` into `target_axis` in schema-declared order.

        Uses the axis element's declared ``_tag_seq`` when available; else
        appends.
        """
        tag_seq = getattr(target_axis.__class__, "_tag_seq", None)
        if tag_seq is None:
            target_axis.append(new_elm)
            return
        try:
            new_idx = tag_seq.index(
                new_elm.tag.replace(
                    "{http://schemas.openxmlformats.org/drawingml/2006/chart}",
                    "c:",
                )
            )
        except ValueError:
            target_axis.append(new_elm)
            return
        # -- find first existing child whose schema index is > new_idx --
        insert_before = None
        for child in target_axis:
            child_short = child.tag.replace(
                "{http://schemas.openxmlformats.org/drawingml/2006/chart}",
                "c:",
            )
            try:
                child_idx = tag_seq.index(child_short)
            except ValueError:
                continue
            if child_idx > new_idx:
                insert_before = child
                break
        if insert_before is not None:
            insert_before.addprevious(new_elm)
        else:
            target_axis.append(new_elm)

    # -- legend --

    def _apply_legend(self):
        """Replace target's ``c:legend`` with a deep copy of the template's.

        When the template has no legend, the target's legend is left alone
        (callers who want to drop the legend can set ``chart.has_legend =
        False`` separately).
        """
        template_legend = self._template.find(qn("c:chart") + "/" + qn("c:legend"))
        if template_legend is None:
            return
        target_chart = self._target.find(qn("c:chart"))
        if target_chart is None:
            return
        for existing in target_chart.findall(qn("c:legend")):
            target_chart.remove(existing)
        new_legend = deepcopy(template_legend)
        # -- insert in schema order: before c:plotVisOnly, after c:plotArea --
        plotArea = target_chart.find(qn("c:plotArea"))
        if plotArea is not None:
            plotArea.addnext(new_legend)
        else:
            target_chart.append(new_legend)

    # -- title (only when target has no title) --

    def _apply_title_if_absent(self):
        """Copy the template's ``c:title`` to the target only when the target
        does not already have one.

        Template titles typically contain placeholder text ("Chart Title"),
        so overwriting an authored title would be surprising. When the
        target has no title, though, the template's title element carries
        useful formatting (font / fill / layout) that is worth adopting.
        """
        target_chart = self._target.find(qn("c:chart"))
        template_chart = self._template.find(qn("c:chart"))
        if target_chart is None or template_chart is None:
            return
        if target_chart.find(qn("c:title")) is not None:
            return
        template_title = template_chart.find(qn("c:title"))
        if template_title is None:
            return
        new_title = deepcopy(template_title)
        # -- title is the first child of c:chart in schema order --
        if len(target_chart) == 0:
            target_chart.append(new_title)
        else:
            target_chart[0].addprevious(new_title)
        # -- ensure autoTitleDeleted is not set to True (if present) --
        autoTitleDeleted = target_chart.find(qn("c:autoTitleDeleted"))
        if autoTitleDeleted is not None:
            autoTitleDeleted.set("val", "0")

    # -- helpers --

    def _insert_in_schema_order(self, parent, new_elm, tag):
        """Insert `new_elm` into `parent` respecting the chartSpace tag order.

        The target is a ``c:chartSpace`` whose ``_tag_seq`` declares the
        canonical child order. Falls back to append when the tag isn't in
        the declared sequence.
        """
        tag_seq = (
            "c:date1904",
            "c:lang",
            "c:roundedCorners",
            "c:style",
            "c:clrMapOvr",
            "c:pivotSource",
            "c:protection",
            "c:chart",
            "c:spPr",
            "c:txPr",
            "c:externalData",
            "c:printSettings",
            "c:userShapes",
            "c:extLst",
        )
        try:
            new_idx = tag_seq.index(tag)
        except ValueError:
            parent.append(new_elm)
            return
        insert_before = None
        for child in parent:
            child_short = child.tag.replace(
                "{http://schemas.openxmlformats.org/drawingml/2006/chart}",
                "c:",
            )
            try:
                child_idx = tag_seq.index(child_short)
            except ValueError:
                continue
            if child_idx > new_idx:
                insert_before = child
                break
        if insert_before is not None:
            insert_before.addprevious(new_elm)
        else:
            parent.append(new_elm)
