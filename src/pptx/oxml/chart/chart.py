"""Custom element classes for top-level chart-related XML elements."""

from __future__ import annotations

from typing import cast

from pptx.enum.chart import XL_DISPLAY_BLANKS_AS
from pptx.oxml import parse_xml
from pptx.oxml.chart.shared import CT_Title
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import ST_Style, ST_StyleEx, XsdString
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_DTable(BaseOxmlElement):
    """`c:dTable` element; the data-table displayed beneath a chart's plot area.

    When present, PowerPoint renders a tabular grid of category and series
    values below the chart's plot area. The four ``c:show*`` children are
    independent booleans controlling border lines, the surrounding outline,
    and the inclusion of a legend-key column (GitHub issue #373).

    See ``CT_DTable`` in ``spec/ISO-IEC-29500-1/schemas/xsd/dml-chart.xsd``.
    """

    _tag_seq = (
        "c:showHorzBorder",
        "c:showVertBorder",
        "c:showOutline",
        "c:showKeys",
        "c:spPr",
        "c:txPr",
        "c:extLst",
    )
    showHorzBorder = ZeroOrOne("c:showHorzBorder", successors=_tag_seq[1:])
    showVertBorder = ZeroOrOne("c:showVertBorder", successors=_tag_seq[2:])
    showOutline = ZeroOrOne("c:showOutline", successors=_tag_seq[3:])
    showKeys = ZeroOrOne("c:showKeys", successors=_tag_seq[4:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[5:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[6:])
    del _tag_seq

    @classmethod
    def new_dTable(cls):
        """Return a newly-created default ``c:dTable`` element.

        PowerPoint writes all four ``c:show*`` children set to ``1`` by
        default (horizontal and vertical borders, outline, and legend keys
        all visible). Matching that default keeps round-trip output aligned
        with what the user sees when they toggle the data-table on via
        PowerPoint's "Data Table" menu.
        """
        return parse_xml(
            "<c:dTable %s>\n"
            '  <c:showHorzBorder val="1"/>\n'
            '  <c:showVertBorder val="1"/>\n'
            '  <c:showOutline val="1"/>\n'
            '  <c:showKeys val="1"/>\n'
            "</c:dTable>" % nsdecls("c")
        )

    def _new_txPr(self):
        return CT_TextBody.new_txPr()


class CT_Chart(BaseOxmlElement):
    """`c:chart` custom element class."""

    _tag_seq = (
        "c:title",
        "c:autoTitleDeleted",
        "c:pivotFmts",
        "c:view3D",
        "c:floor",
        "c:sideWall",
        "c:backWall",
        "c:plotArea",
        "c:legend",
        "c:plotVisOnly",
        "c:dispBlanksAs",
        "c:showDLblsOverMax",
        "c:extLst",
    )
    title = ZeroOrOne("c:title", successors=_tag_seq[1:])
    autoTitleDeleted = ZeroOrOne("c:autoTitleDeleted", successors=_tag_seq[2:])
    plotArea = OneAndOnlyOne("c:plotArea")
    legend = ZeroOrOne("c:legend", successors=_tag_seq[9:])
    plotVisOnly = ZeroOrOne("c:plotVisOnly", successors=_tag_seq[10:])
    dispBlanksAs = ZeroOrOne("c:dispBlanksAs", successors=_tag_seq[11:])
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]

    @property
    def has_legend(self):
        """
        True if this chart has a legend defined, False otherwise.
        """
        legend = self.legend
        if legend is None:
            return False
        return True

    @has_legend.setter
    def has_legend(self, bool_value):
        """
        Add, remove, or leave alone the ``<c:legend>`` child element depending
        on current state and *bool_value*. If *bool_value* is |True| and no
        ``<c:legend>`` element is present, a new default element is added.
        When |False|, any existing legend element is removed.
        """
        if bool(bool_value) is False:
            self._remove_legend()
        else:
            if self.legend is None:
                self._add_legend()

    @staticmethod
    def new_chart(rId: str) -> CT_Chart:
        """Return a new `c:chart` element."""
        return cast(CT_Chart, parse_xml(f'<c:chart {nsdecls("c")} {nsdecls("r")} r:id="{rId}"/>'))

    def _new_title(self):
        return CT_Title.new_title()


class CT_ChartSpace(BaseOxmlElement):
    """`c:chartSpace` root element of a chart part."""

    _tag_seq = (
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
    date1904 = ZeroOrOne("c:date1904", successors=_tag_seq[1:])
    style = ZeroOrOne("c:style", successors=_tag_seq[4:])
    chart = OneAndOnlyOne("c:chart")
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[10:])
    externalData = ZeroOrOne("c:externalData", successors=_tag_seq[11:])

    # NOTE: `mc:AlternateContent` may wrap a `c14:style` (Office 2010+ extended chart-style
    # index used to enable pure-accent series colours past the 6th series) together with a
    # `c:style` fallback for Office 2007-era readers. The wrapper, when present, appears in
    # the same slot as `c:style`; the `style` accessor defined above still resolves the
    # plain-schema child (the `mc:Fallback` contents are not in `style`'s scope), while
    # `chart_style_ex_val` / `set_chart_style_ex_val` below expose the extended value and
    # manage both the wrapper and the bare `c:style` element.
    _chart_style_ex_successors = (
        qn("c:clrMapOvr"),
        qn("c:pivotSource"),
        qn("c:protection"),
        qn("c:chart"),
        qn("c:spPr"),
        qn("c:txPr"),
        qn("c:externalData"),
        qn("c:printSettings"),
        qn("c:userShapes"),
        qn("c:extLst"),
    )
    del _tag_seq

    @property
    def catAx_lst(self):
        return self.chart.plotArea.catAx_lst

    @property
    def chart_style_ex_val(self) -> int | None:
        """The effective chart-style index, preferring the `c14:style` extended value.

        Returns the integer from the `c14:style` child of the first `mc:Choice` (inside the
        `mc:AlternateContent` wrapper) when present, otherwise the `c:style` value, otherwise
        `None`. This lets callers transparently observe the extended-style index that MS
        PowerPoint uses to pick pure-accent series colors past the 6th series.
        """
        style_ex = self._c14_style_elm
        if style_ex is not None:
            return style_ex.val
        if self.style is not None:
            return self.style.val
        return None

    def set_chart_style_ex_val(self, value: int | None) -> None:
        """Set the chart-style index, writing an `mc:AlternateContent` wrapper when `value > 48`.

        When `value` is in the plain 1-48 range a bare `c:style` child is written (removing any
        existing `mc:AlternateContent` style wrapper). When `value` is in the extended 49-255
        range an `mc:AlternateContent` wrapper is written containing a `c14:style val="value"`
        in an `mc:Choice Requires="c14"` and an `mc:Fallback/c:style` set to the base style
        (`value - 100` when `value > 100`, else `value`; clamped to 1..48) so the chart still
        renders in readers that do not know about `c14`. When `value` is `None`, both the plain
        `c:style` and any `mc:AlternateContent` wrapper are removed.
        """
        self._remove_style()
        self._remove_alt_content_style()
        if value is None:
            return
        if 1 <= value <= 48:
            self._add_style(val=value)
            return
        # -- extended range: wrap in mc:AlternateContent with a `c:style` fallback. The
        # -- extended-index convention PowerPoint uses is `base + 100` (e.g. 118 = accent
        # -- variant of plain style 18), so the best plain-reader fallback is usually
        # -- `value - 100`. Clamp anything outside 1..48 to 1 so the fallback always
        # -- satisfies `ST_Style`.
        fallback_val = value - 100 if value > 100 else value
        if fallback_val < 1 or fallback_val > 48:
            fallback_val = 1
        ac_xml = (
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-co'
            'mpatibility/2006">\n'
            '  <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8'
            '/2/chart" Requires="c14">\n'
            '    <c14:style val="%d"/>\n'
            "  </mc:Choice>\n"
            "  <mc:Fallback>\n"
            '    <c:style xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/cha'
            'rt" val="%d"/>\n'
            "  </mc:Fallback>\n"
            "</mc:AlternateContent>"
        ) % (value, fallback_val)
        ac_elm = parse_xml(ac_xml)
        # -- insert in the same slot as a bare c:style would occupy --
        insert_before = None
        for child in self.iterchildren():
            if child.tag in self._chart_style_ex_successors:
                insert_before = child
                break
        if insert_before is not None:
            insert_before.addprevious(ac_elm)
        else:
            self.append(ac_elm)

    @property
    def _c14_style_elm(self) -> CT_StyleEx | None:
        """The `c14:style` element from the first `mc:Choice` of the style `mc:AlternateContent`.

        Returns `None` if no such wrapper is present, or if its first `mc:Choice` contains no
        `c14:style` child.
        """
        # -- only consider mc:AlternateContent children that sit in the `c:style` slot and
        # -- contain a c14:style child in their first mc:Choice; this avoids mis-resolving
        # -- other mc:AlternateContent wrappers that may appear elsewhere in the chartSpace.
        for ac in self.iterchildren(qn("mc:AlternateContent")):
            choices = list(ac.iterchildren(qn("mc:Choice")))
            if not choices:
                continue
            style_ex = choices[0].find(qn("c14:style"))
            if style_ex is not None:
                return cast("CT_StyleEx", style_ex)
        return None

    def _remove_alt_content_style(self) -> None:
        """Remove any `mc:AlternateContent` wrapper whose first `mc:Choice` contains `c14:style`.

        A no-op when no such wrapper is present.
        """
        for ac in list(self.iterchildren(qn("mc:AlternateContent"))):
            choices = list(ac.iterchildren(qn("mc:Choice")))
            if not choices:
                continue
            if choices[0].find(qn("c14:style")) is not None:
                self.remove(ac)

    @property
    def date_1904(self):
        """
        Return |True| if the `c:date1904` child element resolves truthy,
        |False| otherwise. This value indicates whether date number values
        are based on the 1900 or 1904 epoch.
        """
        date1904 = self.date1904
        if date1904 is None:
            return False
        return date1904.val

    @property
    def dateAx_lst(self):
        return self.xpath("c:chart/c:plotArea/c:dateAx")

    def get_or_add_title(self):
        """Return the `c:title` grandchild, newly created if not present."""
        return self.chart.get_or_add_title()

    @property
    def plotArea(self):
        """
        Return the required `c:chartSpace/c:chart/c:plotArea` grandchild
        element.
        """
        return self.chart.plotArea

    @property
    def valAx_lst(self):
        return self.chart.plotArea.valAx_lst

    @property
    def xlsx_part_rId(self):
        """
        The string in the required ``r:id`` attribute of the
        `<c:externalData>` child, or |None| if no externalData element is
        present.
        """
        externalData = self.externalData
        if externalData is None:
            return None
        return externalData.rId

    def _add_externalData(self):
        """
        Always add a ``<c:autoUpdate val="0"/>`` child so auto-updating
        behavior is off by default.
        """
        externalData = self._new_externalData()
        externalData._add_autoUpdate(val=False)
        self._insert_externalData(externalData)
        return externalData

    def _new_txPr(self):
        return CT_TextBody.new_txPr()


class CT_ExternalData(BaseOxmlElement):
    """
    `<c:externalData>` element, defining link to embedded Excel package part
    containing the chart data.
    """

    autoUpdate = ZeroOrOne("c:autoUpdate")
    rId = RequiredAttribute("r:id", XsdString)


class CT_PlotArea(BaseOxmlElement):
    """
    ``<c:plotArea>`` element.
    """

    catAx = ZeroOrMore("c:catAx")
    valAx = ZeroOrMore("c:valAx")
    # -- `c:dTable` is the optional "data-table" grandchild that PowerPoint --
    # -- renders beneath the plot area; in the CT_PlotArea sequence it --
    # -- follows the axis choice and precedes `c:spPr` / `c:extLst`. (Issue --
    # -- #373). --
    dTable = ZeroOrOne("c:dTable", successors=("c:spPr", "c:extLst"))
    # -- `c:spPr` is the optional shape-properties child that carries the --
    # -- plot-area's fill / line / effect formatting. In the CT_PlotArea --
    # -- sequence it follows `c:dTable` and is followed only by `c:extLst`. --
    # -- See ISO/IEC 29500-1 `spec/…/dml-chart.xsd`. (Issue #298). --
    spPr = ZeroOrOne("c:spPr", successors=("c:extLst",))

    @property
    def dateAx_lst(self):
        """Return all ``c:dateAx`` child elements of this plotArea."""
        return self.xpath("./c:dateAx")

    @property
    def has_category_axis(self):
        """True if this plot-area contains at least one category-type axis.

        A "category-type axis" is any of ``c:catAx`` or ``c:dateAx``. Charts whose
        independent variable is categorical (bar, line, column, area, etc.) carry
        one; an XY/scatter chart does not.
        """
        return bool(self.catAx_lst or self.dateAx_lst)

    @property
    def primary_valAx(self):
        """Return the primary `c:valAx` element, or |None| if none is present.

        The primary value axis is the first `c:valAx` child in document order.
        """
        valAx_lst = self.valAx_lst
        return valAx_lst[0] if valAx_lst else None

    @property
    def secondary_valAx(self):
        """Return the secondary `c:valAx` element, or |None| if not present.

        A secondary value axis is only meaningful for a chart that also has a
        category-type axis (`c:catAx` or `c:dateAx`). In that case, the second
        `c:valAx` element in document order is the secondary value axis. For
        an XY/scatter chart both `c:valAx` elements are primary axes (one for
        X and one for Y), so this always returns |None| in that case.
        """
        if not self.has_category_axis:
            return None
        valAx_lst = self.valAx_lst
        if len(valAx_lst) < 2:
            return None
        return valAx_lst[1]

    @property
    def secondary_catAx(self):
        """Return the secondary `c:catAx` element, or |None| if not present.

        The secondary category axis is the second `c:catAx` child in document
        order. It is typically emitted by PowerPoint as a hidden companion to
        a secondary value axis to satisfy the `c:crossAx` pairing requirement.
        """
        catAx_lst = self.catAx_lst
        if len(catAx_lst) < 2:
            return None
        return catAx_lst[1]

    def iter_sers(self):
        """
        Generate each of the `c:ser` elements in this chart, ordered first by
        the document order of the containing xChart element, then by their
        ordering within the xChart element (not necessarily document order).
        """
        for xChart in self.iter_xCharts():
            for ser in xChart.iter_sers():
                yield ser

    def iter_xCharts(self):
        """
        Generate each xChart child element in document.
        """
        plot_tags = (
            qn("c:area3DChart"),
            qn("c:areaChart"),
            qn("c:bar3DChart"),
            qn("c:barChart"),
            qn("c:bubbleChart"),
            qn("c:doughnutChart"),
            qn("c:line3DChart"),
            qn("c:lineChart"),
            qn("c:ofPieChart"),
            qn("c:pie3DChart"),
            qn("c:pieChart"),
            qn("c:radarChart"),
            qn("c:scatterChart"),
            qn("c:stockChart"),
            qn("c:surface3DChart"),
            qn("c:surfaceChart"),
        )

        for child in self.iterchildren():
            if child.tag not in plot_tags:
                continue
            yield child

    @property
    def last_ser(self):
        """
        Return the last `<c:ser>` element in the last xChart element, based
        on series order (not necessarily the same element as document order).
        """
        last_xChart = self.xCharts[-1]
        sers = last_xChart.sers
        if not sers:
            return None
        return sers[-1]

    @property
    def next_idx(self):
        """
        Return the next available `c:ser/c:idx` value within the scope of
        this chart, the maximum idx value found on existing series,
        incremented by one.
        """
        idx_vals = [s.idx.val for s in self.sers]
        if not idx_vals:
            return 0
        return max(idx_vals) + 1

    @property
    def next_order(self):
        """
        Return the next available `c:ser/c:order` value within the scope of
        this chart, the maximum order value found on existing series,
        incremented by one.
        """
        order_vals = [s.order.val for s in self.sers]
        if not order_vals:
            return 0
        return max(order_vals) + 1

    @property
    def sers(self):
        """
        Return a sequence containing all the `c:ser` elements in this chart,
        ordered first by the document order of the containing xChart element,
        then by their ordering within the xChart element (not necessarily
        document order).
        """
        return tuple(self.iter_sers())

    @property
    def xCharts(self):
        """
        Return a sequence containing all the `c:{x}Chart` elements in this
        chart, in document order.
        """
        return tuple(self.iter_xCharts())


class CT_Style(BaseOxmlElement):
    """
    ``<c:style>`` element; defines the chart style.
    """

    val = RequiredAttribute("val", ST_Style)


class CT_StyleEx(BaseOxmlElement):
    """`c14:style` element; the Office 2010+ extended chart-style index.

    Appears inside an `mc:Choice Requires="c14"` child of an `mc:AlternateContent` wrapper that
    sits in the `c:style` slot of `c:chartSpace`. The `val` attribute carries an extended
    chart-style index (typically the plain 1-48 index plus 100, e.g. `118 = AccentN-coloured
    variant of style 18`) that lets MS PowerPoint render pure-accent series colours for charts
    with six or more series instead of shaded alternates of the first six accents.
    """

    val = RequiredAttribute("val", ST_StyleEx)


class CT_DispBlanksAs(BaseOxmlElement):
    """`c:dispBlanksAs` element; how blank cells are displayed in the chart.

    Maps to the "Show empty cells as" radio control in PowerPoint's "Hidden and
    Empty Cells" dialog (``ST_DispBlanksAs`` in ``dml-chart.xsd``). The ``val``
    attribute defaults to ``"zero"`` per the XSD when the attribute is omitted.
    """

    val = OptionalAttribute("val", XL_DISPLAY_BLANKS_AS, default=XL_DISPLAY_BLANKS_AS.ZERO)
