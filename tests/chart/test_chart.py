# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.chart.chart` module."""

from __future__ import annotations

import pytest

from pptx.chart.axis import CategoryAxis, DateAxis, ValueAxis
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.chart.chart import Chart, ChartTitle, Legend, PlotArea, _Plots
from pptx.chart.data import BubbleChartData, CategoryChartData, ChartData, XyChartData
from pptx.chart.plot import _BasePlot
from pptx.chart.series import SeriesCollection
from pptx.chart.xmlwriter import _BaseSeriesXmlRewriter
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_CHART_TYPE
from pptx.parts.chart import ChartWorkbook
from pptx.text.text import Font

from ..unitutil.cxml import element, xml
from ..unitutil.mock import (
    call,
    class_mock,
    function_mock,
    instance_mock,
    property_mock,
)


class DescribeChart(object):
    """Unit-test suite for `pptx.chart.chart.Chart` objects."""

    def it_provides_access_to_its_font(self, font_fixture, Font_, font_):
        chartSpace, expected_xml = font_fixture
        Font_.return_value = font_
        chart = Chart(chartSpace, None)

        font = chart.font

        assert chartSpace.xml == expected_xml
        Font_.assert_called_once_with(chartSpace.xpath("./c:txPr/a:p/a:pPr/a:defRPr")[0])
        assert font is font_

    def it_knows_whether_it_has_a_title(self, has_title_get_fixture):
        chart, expected_value = has_title_get_fixture
        assert chart.has_title is expected_value

    def it_can_change_whether_it_has_a_title(self, has_title_set_fixture):
        chart, new_value, expected_xml = has_title_set_fixture
        chart.has_title = new_value
        assert chart._chartSpace.chart.xml == expected_xml

    def it_provides_access_to_the_chart_title(self, title_fixture):
        chart, expected_xml, ChartTitle_, chart_title_ = title_fixture

        chart_title = chart.chart_title

        assert chart.element.xpath("c:chart/c:title")[0].xml == expected_xml
        ChartTitle_.assert_called_once_with(chart.element.chart.title)
        assert chart_title is chart_title_

    def it_provides_access_to_the_category_axis(self, category_axis_fixture):
        chart, category_axis_, AxisCls_, xAx = category_axis_fixture
        category_axis = chart.category_axis
        AxisCls_.assert_called_once_with(xAx)
        assert category_axis is category_axis_

    def it_raises_when_no_category_axis(self, cat_ax_raise_fixture):
        chart = cat_ax_raise_fixture
        with pytest.raises(ValueError):
            chart.category_axis

    def it_provides_access_to_the_value_axis(self, val_ax_fixture):
        chart, ValueAxis_, valAx, value_axis_ = val_ax_fixture
        value_axis = chart.value_axis
        ValueAxis_.assert_called_once_with(valAx)
        assert value_axis is value_axis_

    def it_raises_when_no_value_axis(self, val_ax_raise_fixture):
        chart = val_ax_raise_fixture
        with pytest.raises(ValueError):
            chart.value_axis

    def it_knows_whether_it_has_a_secondary_value_axis(self, has_sec_val_ax_fixture):
        chart, expected_value = has_sec_val_ax_fixture
        assert chart.has_secondary_value_axis is expected_value

    def it_provides_access_to_the_secondary_value_axis(self, sec_val_ax_fixture):
        chart, ValueAxis_, sec_valAx, value_axis_ = sec_val_ax_fixture
        secondary_value_axis = chart.secondary_value_axis
        ValueAxis_.assert_called_once_with(sec_valAx)
        assert secondary_value_axis is value_axis_

    def it_raises_when_no_secondary_value_axis(self, sec_val_ax_raise_fixture):
        chart = sec_val_ax_raise_fixture
        with pytest.raises(ValueError):
            chart.secondary_value_axis

    def it_provides_access_to_its_series(self, series_fixture):
        chart, SeriesCollection_, plotArea, series_ = series_fixture
        series = chart.series
        SeriesCollection_.assert_called_once_with(plotArea)
        assert series is series_

    def it_provides_access_to_its_plots(self, plots_fixture):
        chart, plots_, _Plots_, plotArea = plots_fixture
        plots = chart.plots
        _Plots_.assert_called_once_with(plotArea, chart)
        assert plots is plots_

    def it_provides_access_to_its_plot_area(self, request):
        """`Chart.plot_area` returns a `PlotArea` on the `c:plotArea` element.

        Addresses issue #298 so callers can reach `chart.plot_area.format.fill`
        and `chart.plot_area.format.line` without dropping into XML.
        """
        chartSpace = element("c:chartSpace/c:chart/c:plotArea")
        plotArea = chartSpace.xpath("./c:chart/c:plotArea")[0]
        plot_area_ = instance_mock(request, PlotArea)
        PlotArea_ = class_mock(request, "pptx.chart.chart.PlotArea", return_value=plot_area_)
        chart = Chart(chartSpace, None)

        plot_area = chart.plot_area

        PlotArea_.assert_called_once_with(plotArea)
        assert plot_area is plot_area_

    def it_knows_whether_it_has_a_legend(self, has_legend_get_fixture):
        chart, expected_value = has_legend_get_fixture
        assert chart.has_legend == expected_value

    def it_can_change_whether_it_has_a_legend(self, has_legend_set_fixture):
        chart, new_value, expected_xml = has_legend_set_fixture
        chart.has_legend = new_value
        assert chart._chartSpace.xml == expected_xml

    def it_provides_access_to_its_legend(self, legend_fixture):
        chart, Legend_, expected_calls, expected_value = legend_fixture
        legend = chart.legend
        assert Legend_.call_args_list == expected_calls
        assert legend is expected_value

    def it_knows_its_chart_type(self, request, PlotTypeInspector_, plot_):
        property_mock(request, Chart, "plots", return_value=[plot_])
        PlotTypeInspector_.chart_type.return_value = XL_CHART_TYPE.PIE
        chart = Chart(None, None)

        chart_type = chart.chart_type

        PlotTypeInspector_.chart_type.assert_called_once_with(plot_)
        assert chart_type == XL_CHART_TYPE.PIE

    def it_knows_its_style(self, style_get_fixture):
        chart, expected_value = style_get_fixture
        assert chart.chart_style == expected_value

    def it_can_change_its_style(self, style_set_fixture):
        chart, new_value, expected_xml = style_set_fixture
        chart.chart_style = new_value
        assert chart._chartSpace.xml == expected_xml

    def it_reads_the_c14_extended_style_when_mc_AlternateContent_wraps_it(self):
        # -- #516 — chart style wrapped in mc:AlternateContent should be surfaced --
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006'
            '/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibili'
            'ty/2006">\n'
            "  <mc:AlternateContent>\n"
            '    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/'
            '2007/8/2/chart" Requires="c14">\n'
            '      <c14:style val="118"/>\n'
            "    </mc:Choice>\n"
            "    <mc:Fallback>\n"
            '      <c:style val="18"/>\n'
            "    </mc:Fallback>\n"
            "  </mc:AlternateContent>\n"
            "  <c:chart><c:plotArea/></c:chart>\n"
            "</c:chartSpace>"
        )
        from pptx.oxml import parse_xml

        chart = Chart(parse_xml(chartSpace_xml), None)
        assert chart.chart_style == 118

    def it_falls_back_to_plain_style_when_no_alternate_content_wrapper(self):
        chart = Chart(element("c:chartSpace/c:style{val=34}"), None)
        assert chart.chart_style == 34

    def it_returns_None_when_neither_style_nor_c14_style_present(self):
        chart = Chart(element("c:chartSpace"), None)
        assert chart.chart_style is None

    def it_can_set_an_extended_chart_style_value(self):
        chart = Chart(element("c:chartSpace/c:chart"), None)

        chart.chart_style = 118

        # -- reading back yields the extended value --
        assert chart.chart_style == 118
        # -- wrapper was inserted --
        ac = chart._chartSpace.find(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
            "AlternateContent"
        )
        assert ac is not None
        choice = ac.find(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice"
        )
        assert choice.get("Requires") == "c14"
        c14_style = choice.find(
            "{http://schemas.microsoft.com/office/drawing/2007/8/2/chart}style"
        )
        assert c14_style is not None
        assert c14_style.get("val") == "118"
        fallback = ac.find(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
        )
        c_style = fallback.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}style"
        )
        # -- fallback is the base style (118 % 100 = 18) --
        assert c_style.get("val") == "18"

    def it_drops_the_AlternateContent_wrapper_when_set_to_a_plain_value(self):
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006'
            '/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibili'
            'ty/2006">\n'
            "  <mc:AlternateContent>\n"
            '    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/'
            '2007/8/2/chart" Requires="c14">\n'
            '      <c14:style val="118"/>\n'
            "    </mc:Choice>\n"
            "    <mc:Fallback>\n"
            '      <c:style val="18"/>\n'
            "    </mc:Fallback>\n"
            "  </mc:AlternateContent>\n"
            "  <c:chart><c:plotArea/></c:chart>\n"
            "</c:chartSpace>"
        )
        from pptx.oxml import parse_xml

        chart = Chart(parse_xml(chartSpace_xml), None)

        chart.chart_style = 4

        assert chart.chart_style == 4
        assert (
            chart._chartSpace.find(
                "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
                "AlternateContent"
            )
            is None
        )

    def it_drops_both_plain_and_extended_style_when_set_to_None(self):
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006'
            '/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibili'
            'ty/2006">\n'
            "  <mc:AlternateContent>\n"
            '    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/'
            '2007/8/2/chart" Requires="c14">\n'
            '      <c14:style val="118"/>\n'
            "    </mc:Choice>\n"
            "    <mc:Fallback>\n"
            '      <c:style val="18"/>\n'
            "    </mc:Fallback>\n"
            "  </mc:AlternateContent>\n"
            "  <c:chart><c:plotArea/></c:chart>\n"
            "</c:chartSpace>"
        )
        from pptx.oxml import parse_xml

        chart = Chart(parse_xml(chartSpace_xml), None)

        chart.chart_style = None

        assert chart.chart_style is None
        assert (
            chart._chartSpace.find(
                "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
                "AlternateContent"
            )
            is None
        )
        assert chart._chartSpace.style is None

    def it_can_replace_the_chart_data(self, replace_fixture):
        (
            chart,
            chart_data_,
            SeriesXmlRewriterFactory_,
            chart_type,
            rewriter_,
            chartSpace,
            workbook_,
            xlsx_blob,
        ) = replace_fixture

        chart.replace_data(chart_data_)

        SeriesXmlRewriterFactory_.assert_called_once_with(chart_type, chart_data_)
        rewriter_.replace_series_data.assert_called_once_with(chartSpace)
        workbook_.update_from_xlsx_blob.assert_called_once_with(xlsx_blob)

    @pytest.mark.parametrize(
        "chart_type, chart_data_factory, expected_substr",
        [
            # -- XY chart rejects CategoryChartData / ChartData (#396) --
            (XL_CHART_TYPE.XY_SCATTER, CategoryChartData, r"XY \(scatter\) chart"),
            (XL_CHART_TYPE.XY_SCATTER_LINES, ChartData, r"XY \(scatter\) chart"),
            (XL_CHART_TYPE.XY_SCATTER_SMOOTH, CategoryChartData, r"XY \(scatter\) chart"),
            # -- XY chart also rejects BubbleChartData (wrong shape) --
            (XL_CHART_TYPE.XY_SCATTER, BubbleChartData, r"XY \(scatter\) chart"),
            # -- Bubble chart rejects CategoryChartData / XyChartData --
            (XL_CHART_TYPE.BUBBLE, CategoryChartData, "bubble chart"),
            (XL_CHART_TYPE.BUBBLE, XyChartData, "bubble chart"),
            (XL_CHART_TYPE.BUBBLE_THREE_D_EFFECT, ChartData, "bubble chart"),
            # -- Category chart rejects XyChartData / BubbleChartData --
            (XL_CHART_TYPE.BAR_CLUSTERED, XyChartData, "category chart"),
            (XL_CHART_TYPE.LINE, BubbleChartData, "category chart"),
            (XL_CHART_TYPE.PIE, XyChartData, "category chart"),
        ],
    )
    def it_raises_on_replace_data_chart_type_mismatch(
        self,
        request,
        chart_type,
        chart_data_factory,
        expected_substr,
        workbook_prop_,
        workbook_,
        SeriesXmlRewriterFactory_,
    ):
        """Chart.replace_data() rejects mismatched ChartData types (#396)."""
        chartSpace = element("c:chartSpace/c:chart/c:plotArea")
        chart = Chart(chartSpace, None)
        property_mock(
            request, Chart, "chart_type", return_value=chart_type
        )
        chart_data = chart_data_factory()

        with pytest.raises(ValueError, match=expected_substr):
            chart.replace_data(chart_data)

        # -- validation short-circuits before any rewrite happens --
        SeriesXmlRewriterFactory_.assert_not_called()
        workbook_.update_from_xlsx_blob.assert_not_called()

    @pytest.mark.parametrize(
        "chart_type, chart_data_cls",
        [
            # -- Category-family types accept CategoryChartData / ChartData --
            (XL_CHART_TYPE.BAR_CLUSTERED, CategoryChartData),
            (XL_CHART_TYPE.LINE, ChartData),
            (XL_CHART_TYPE.PIE, CategoryChartData),
            # -- XY accepts XyChartData --
            (XL_CHART_TYPE.XY_SCATTER, XyChartData),
            # -- Bubble accepts BubbleChartData --
            (XL_CHART_TYPE.BUBBLE, BubbleChartData),
        ],
    )
    def it_accepts_matching_chart_data_type(
        self,
        request,
        chart_type,
        chart_data_cls,
        workbook_prop_,
        workbook_,
        SeriesXmlRewriterFactory_,
        series_rewriter_,
    ):
        """Chart.replace_data() accepts matching ChartData subclasses (#396)."""
        chartSpace = element("c:chartSpace/c:chart/c:plotArea")
        chart = Chart(chartSpace, None)
        property_mock(
            request, Chart, "chart_type", return_value=chart_type
        )
        chart_data = instance_mock(request, chart_data_cls)

        chart.replace_data(chart_data)

        SeriesXmlRewriterFactory_.assert_called_once_with(chart_type, chart_data)
        series_rewriter_.replace_series_data.assert_called_once_with(chartSpace)
        workbook_.update_from_xlsx_blob.assert_called_once_with(chart_data.xlsx_blob)

    def it_refreshes_cached_values_from_the_embedded_xlsx(
        self, update_cached_fixture
    ):
        chart, expected_xml = update_cached_fixture

        chart.update_cached_values()

        assert chart._chartSpace.xml == expected_xml

    # -- Chart.apply_template (issue #243) ---------------------------

    def it_can_apply_a_chart_template_243(self):
        """``Chart.apply_template`` copies chart-space formatting from a
        ``.crtx`` package onto an existing chart, leaving data alone.
        """
        import io
        import zipfile

        target_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml'
            '/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocu'
            'ment/2006/relationships">'
            '<c:style val="2"/>'
            "<c:chart><c:plotArea>"
            "<c:barChart>"
            '<c:ser><c:idx val="0"/><c:order val="0"/>'
            '<c:val><c:numRef><c:f>Sheet1!$A$1</c:f></c:numRef></c:val>'
            "</c:ser>"
            '<c:axId val="1"/><c:axId val="2"/>'
            "</c:barChart>"
            '<c:catAx><c:axId val="1"/></c:catAx>'
            '<c:valAx><c:axId val="2"/></c:valAx>'
            "</c:plotArea></c:chart>"
            "</c:chartSpace>"
        )
        template_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml'
            '/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-co'
            'mpatibility/2006">'
            "<mc:AlternateContent>"
            '<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com'
            '/office/drawing/2007/8/2/chart"><c14:style val="118"/></mc:Choice>'
            '<mc:Fallback><c:style val="18"/></mc:Fallback>'
            "</mc:AlternateContent>"
            "<c:chart><c:plotArea>"
            "<c:barChart>"
            '<c:axId val="100"/><c:axId val="200"/>'
            "</c:barChart>"
            '<c:catAx><c:axId val="100"/><c:majorTickMark val="out"/></c:catAx>'
            '<c:valAx><c:axId val="200"/><c:numFmt formatCode="0.00%" source'
            'Linked="0"/></c:valAx>'
            "</c:plotArea>"
            '<c:legend><c:legendPos val="r"/></c:legend>'
            "</c:chart>"
            '<c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></c:spPr>'
            "<c:txPr><a:bodyPr/><a:lstStyle/>"
            '<a:p><a:pPr><a:defRPr sz="1200"/></a:pPr></a:p>'
            "</c:txPr>"
            "</c:chartSpace>"
        )
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("chart/chart1.xml", template_xml)

        chartSpace = parse_xml(target_xml)
        chart = Chart(chartSpace, None)

        chart.apply_template(buf.getvalue())

        # -- chart-style upgraded to 118 (extended) --
        assert chartSpace.chart_style_ex_val == 118
        # -- chartSpace-level spPr / txPr copied --
        assert len(chartSpace.findall(qn("c:spPr"))) == 1
        assert (
            chartSpace.find(qn("c:spPr") + "/" + qn("a:solidFill")) is not None
        )
        assert len(chartSpace.findall(qn("c:txPr"))) == 1
        # -- legend copied onto target (target had none) --
        legend = chartSpace.find(qn("c:chart") + "/" + qn("c:legend"))
        assert legend is not None
        # -- catAx majorTickMark copied, axId preserved --
        catAx = chartSpace.find(
            qn("c:chart") + "/" + qn("c:plotArea") + "/" + qn("c:catAx")
        )
        assert catAx.find(qn("c:majorTickMark")).get("val") == "out"
        assert catAx.find(qn("c:axId")).get("val") == "1"
        # -- valAx numFmt copied --
        valAx = chartSpace.find(
            qn("c:chart") + "/" + qn("c:plotArea") + "/" + qn("c:valAx")
        )
        assert valAx.find(qn("c:numFmt")).get("formatCode") == "0.00%"
        # -- series data preserved (not overwritten) --
        ser = chartSpace.find(
            qn("c:chart") + "/" + qn("c:plotArea")
            + "/" + qn("c:barChart") + "/" + qn("c:ser")
        )
        val_f = ser.find(
            qn("c:val") + "/" + qn("c:numRef") + "/" + qn("c:f")
        )
        assert val_f.text == "Sheet1!$A$1"

    def it_raises_when_template_is_not_a_zip_243(self):
        chartSpace = parse_xml(
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocu'
            'ment/2006/relationships">'
            "<c:chart><c:plotArea/></c:chart>"
            "</c:chartSpace>"
        )
        chart = Chart(chartSpace, None)
        with pytest.raises(ValueError, match="not a valid .crtx"):
            chart.apply_template(b"not a zip file")

    def it_raises_when_template_has_no_chartSpace_243(self):
        import io
        import zipfile

        chartSpace = parse_xml(
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocu'
            'ment/2006/relationships">'
            "<c:chart><c:plotArea/></c:chart>"
            "</c:chartSpace>"
        )
        chart = Chart(chartSpace, None)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("not-a-chart.xml", "<root/>")
        with pytest.raises(ValueError, match="no c:chartSpace"):
            chart.apply_template(buf.getvalue())

    # -- Chart.add_plot (combo chart, issue #338) --------------------

    def it_can_add_a_line_plot_to_an_existing_bar_chart_338(self):
        from pptx.chart.data import CategoryChartData

        # -- a one-plot barChart with category/value axes sharing axIds --
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea>"
            "<c:barChart>"
            '<c:barDir val="col"/>'
            '<c:grouping val="clustered"/>'
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            '<c:axId val="111"/>'
            '<c:axId val="222"/>'
            "</c:barChart>"
            '<c:catAx><c:axId val="111"/></c:catAx>'
            '<c:valAx><c:axId val="222"/></c:valAx>'
            "</c:plotArea></c:chart></c:chartSpace>"
        )
        chart = Chart(parse_xml(chartSpace_xml), None)
        data = CategoryChartData()
        data.categories = ["Q1", "Q2", "Q3"]
        data.add_series("Trend", (10, 20, 30))

        new_plot = chart.add_plot(XL_CHART_TYPE.LINE, data)

        # -- a second plot is appended --
        assert len(chart.plots) == 2
        # -- the new plot is a LinePlot --
        assert type(new_plot).__name__ == "LinePlot"
        # -- both plots share the same two axIds --
        xCharts = chart._chartSpace.plotArea.xCharts
        bar_axIds = [e.get("val") for e in xCharts[0].findall(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}axId"
        )]
        line_axIds = [e.get("val") for e in xCharts[1].findall(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}axId"
        )]
        assert bar_axIds == line_axIds == ["111", "222"]
        # -- the new series index/order was offset past existing series --
        new_ser = xCharts[1].find(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}ser"
        )
        new_idx = new_ser.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}idx"
        )
        assert new_idx.get("val") == "1"
        # -- inline numLit is used rather than numRef (no workbook rewrite) --
        C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        numLits = xCharts[1].findall(f".//{C_NS}numLit")
        assert len(numLits) == 1
        # -- inline strLit used for categories --
        strLits = xCharts[1].findall(f".//{C_NS}cat/{C_NS}strLit")
        assert len(strLits) == 1

    def it_can_add_a_bar_plot_to_an_existing_line_chart_338(self):
        from pptx.chart.data import CategoryChartData

        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea>"
            "<c:lineChart>"
            '<c:grouping val="standard"/>'
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            '<c:ser><c:idx val="1"/><c:order val="1"/></c:ser>'
            '<c:axId val="111"/>'
            '<c:axId val="222"/>'
            "</c:lineChart>"
            '<c:catAx><c:axId val="111"/></c:catAx>'
            '<c:valAx><c:axId val="222"/></c:valAx>'
            "</c:plotArea></c:chart></c:chartSpace>"
        )
        chart = Chart(parse_xml(chartSpace_xml), None)
        data = CategoryChartData()
        data.categories = ["a", "b"]
        data.add_series("Bars", (5.0, 6.0))

        plot = chart.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED, data)

        assert type(plot).__name__ == "BarPlot"
        assert len(chart.plots) == 2
        # -- new series idx starts at 2 (count of existing series) --
        bar_xChart = chart._chartSpace.plotArea.xCharts[1]
        new_ser = bar_xChart.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}ser"
        )
        assert new_ser.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}idx"
        ).get("val") == "2"

    def it_raises_when_add_plot_called_on_empty_plot_area(self):
        from pptx.chart.data import CategoryChartData

        chart = Chart(element("c:chartSpace/c:chart/c:plotArea"), None)
        data = CategoryChartData()
        data.categories = ["a"]
        data.add_series("s", (1,))
        with pytest.raises(ValueError, match="no existing plot"):
            chart.add_plot(XL_CHART_TYPE.LINE, data)

    def it_raises_when_add_plot_chart_type_is_unsupported(self):
        from pptx.chart.data import CategoryChartData

        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea>"
            "<c:barChart>"
            '<c:barDir val="col"/>'
            '<c:grouping val="clustered"/>'
            '<c:axId val="1"/><c:axId val="2"/>'
            "</c:barChart>"
            "</c:plotArea></c:chart></c:chartSpace>"
        )
        chart = Chart(parse_xml(chartSpace_xml), None)
        data = CategoryChartData()
        data.categories = ["a"]
        data.add_series("s", (1,))
        with pytest.raises(NotImplementedError, match="combo-chart plot fragment"):
            chart.add_plot(XL_CHART_TYPE.PIE, data)

    # -- Chart.update_cached_values follows below --------------------

    def it_is_a_noop_when_the_chart_has_no_embedded_xlsx(
        self, request, workbook_prop_, workbook_
    ):
        workbook_.xlsx_part = None
        chartSpace = element("c:chartSpace/c:chart")
        chart = Chart(chartSpace, None)
        # -- make sure no WorkbookReader is constructed in the no-xlsx path --
        WorkbookReader_ = class_mock(request, "pptx.chart.chart.WorkbookReader")
        original_xml = chartSpace.xml

        chart.update_cached_values()

        assert WorkbookReader_.call_args_list == []
        assert chartSpace.xml == original_xml

    # -- Chart.clone_to (F5 + F1 cross-slide chart copy, #877) -------

    def it_can_clone_itself_onto_another_shape_tree(self, request):
        """Chart.clone_to delegates to shapes.clone_chart, returning the new GraphicFrame."""
        from unittest.mock import MagicMock

        chart = Chart(element("c:chartSpace"), None)
        shapes_ = MagicMock()
        clone_result_ = MagicMock()
        shapes_.clone_chart.return_value = clone_result_

        result = chart.clone_to(shapes_, 11, 22, 33, 44)

        shapes_.clone_chart.assert_called_once_with(chart, 11, 22, 33, 44)
        assert result is clone_result_

    # -- Chart.user_shapes / has_user_shapes (issue #351) -------------

    def it_returns_the_related_chart_drawing_part_for_user_shapes(self, request):
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        from pptx.parts.chart import ChartPart
        from pptx.parts.chartdrawing import ChartDrawingPart

        chart_drawing_part_ = instance_mock(request, ChartDrawingPart)
        chart_part_ = instance_mock(request, ChartPart)
        chart_part_.part_related_by.return_value = chart_drawing_part_
        part_prop_ = property_mock(request, Chart, "part")
        part_prop_.return_value = chart_part_
        chart = Chart(element("c:chartSpace"), None)

        result = chart.user_shapes

        chart_part_.part_related_by.assert_called_once_with(RT.CHART_USER_SHAPES)
        assert result is chart_drawing_part_

    def but_user_shapes_returns_None_when_no_such_relationship_exists(self, request):
        from pptx.parts.chart import ChartPart

        chart_part_ = instance_mock(request, ChartPart)
        chart_part_.part_related_by.side_effect = KeyError("no such rel")
        part_prop_ = property_mock(request, Chart, "part")
        part_prop_.return_value = chart_part_
        chart = Chart(element("c:chartSpace"), None)

        assert chart.user_shapes is None

    def it_reports_True_for_has_user_shapes_when_anchors_exist(self, request):
        from pptx.parts.chartdrawing import ChartDrawingPart

        chart_drawing_part_ = instance_mock(request, ChartDrawingPart)
        chart_drawing_part_.anchor_count = 2
        user_shapes_prop_ = property_mock(request, Chart, "user_shapes")
        user_shapes_prop_.return_value = chart_drawing_part_
        chart = Chart(element("c:chartSpace"), None)

        assert chart.has_user_shapes is True

    def but_has_user_shapes_is_False_when_part_absent(self, request):
        user_shapes_prop_ = property_mock(request, Chart, "user_shapes")
        user_shapes_prop_.return_value = None
        chart = Chart(element("c:chartSpace"), None)

        assert chart.has_user_shapes is False

    def and_has_user_shapes_is_False_when_part_has_no_anchors(self, request):
        from pptx.parts.chartdrawing import ChartDrawingPart

        chart_drawing_part_ = instance_mock(request, ChartDrawingPart)
        chart_drawing_part_.anchor_count = 0
        user_shapes_prop_ = property_mock(request, Chart, "user_shapes")
        user_shapes_prop_.return_value = chart_drawing_part_
        chart = Chart(element("c:chartSpace"), None)

        assert chart.has_user_shapes is False

    # -- Chart.workbook property (F5) --------------------------------

    def it_reads_embedded_workbook_bytes_via_workbook_property(
        self, workbook_prop_, workbook_
    ):
        workbook_.xlsx_part = _XlsxPartStub(b"blob-bytes")
        chart = Chart(element("c:chartSpace"), None)

        assert chart.workbook == b"blob-bytes"

    def but_workbook_returns_None_when_chart_has_no_embedded_xlsx(
        self, workbook_prop_, workbook_
    ):
        workbook_.xlsx_part = None
        chart = Chart(element("c:chartSpace"), None)

        assert chart.workbook is None

    def it_updates_the_embedded_workbook_bytes_on_assignment(
        self, workbook_prop_, workbook_
    ):
        chart = Chart(element("c:chartSpace"), None)

        chart.workbook = b"new-bytes"

        workbook_.update_from_xlsx_blob.assert_called_once_with(b"new-bytes")

    def it_raises_TypeError_when_assigning_non_bytes_to_workbook(
        self, workbook_prop_, workbook_
    ):
        chart = Chart(element("c:chartSpace"), None)

        with pytest.raises(TypeError, match="must be set to a bytes object"):
            chart.workbook = "not-bytes"

    # -- update_embedded_xlsx_cell helper (F5) ------------------------

    def it_updates_workbook_and_matching_numCache_via_helper(
        self, request, workbook_prop_, workbook_
    ):
        from pptx.chart.chart import update_embedded_xlsx_cell

        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea><c:barChart><c:ser>"
            "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>"
            '<c:ptCount val="2"/>'
            '<c:pt idx="0"><c:v>1</c:v></c:pt>'
            '<c:pt idx="1"><c:v>2</c:v></c:pt>'
            "</c:numCache></c:numRef></c:val>"
            "</c:ser></c:barChart></c:plotArea></c:chart>"
            "</c:chartSpace>"
        )
        chartSpace = parse_xml(chartSpace_xml)
        chart = Chart(chartSpace, None)
        workbook_.xlsx_part = _XlsxPartStub(
            _build_update_cached_xlsx_blob({}, {("Sheet1", 2, 2): 1.0, ("Sheet1", 3, 2): 2.0})
        )

        update_embedded_xlsx_cell(chart, "Sheet1", "B2", 42.0)

        # --- blob was refreshed via ChartWorkbook.update_from_xlsx_blob ---
        workbook_.update_from_xlsx_blob.assert_called_once()
        # --- the numCache cell for B2 (idx 0) now reads 42 ---
        C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        numCache = chartSpace.find(f".//{C}numCache")
        assert numCache is not None
        # locate pt idx=0
        for pt in numCache.findall(f"{C}pt"):
            if pt.get("idx") == "0":
                assert pt.find(f"{C}v").text == "42"
                break
        else:
            pytest.fail("numCache pt idx=0 not found")

    def it_updates_workbook_and_strCache_for_a_string_cell(
        self, request, workbook_prop_, workbook_
    ):
        from pptx.chart.chart import update_embedded_xlsx_cell

        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea><c:barChart><c:ser>"
            "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>"
            '<c:ptCount val="1"/>'
            '<c:pt idx="0"><c:v>Old</c:v></c:pt>'
            "</c:strCache></c:strRef></c:tx>"
            "</c:ser></c:barChart></c:plotArea></c:chart>"
            "</c:chartSpace>"
        )
        chartSpace = parse_xml(chartSpace_xml)
        chart = Chart(chartSpace, None)
        workbook_.xlsx_part = _XlsxPartStub(
            _build_update_cached_xlsx_blob({("Sheet1", 1, 2): "Old"}, {})
        )

        update_embedded_xlsx_cell(chart, "Sheet1", "B1", "New")

        C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        v = chartSpace.find(f".//{C}strCache/{C}pt/{C}v")
        assert v.text == "New"

    def it_raises_when_chart_has_no_embedded_workbook(
        self, workbook_prop_, workbook_
    ):
        from pptx.chart.chart import update_embedded_xlsx_cell

        workbook_.xlsx_part = None
        chart = Chart(element("c:chartSpace"), None)
        with pytest.raises(ValueError, match="no embedded workbook"):
            update_embedded_xlsx_cell(chart, "Sheet1", "B2", 7)

    def it_raises_when_a1_ref_is_sheet_qualified(
        self, workbook_prop_, workbook_
    ):
        from pptx.chart.chart import update_embedded_xlsx_cell

        workbook_.xlsx_part = _XlsxPartStub(
            _build_update_cached_xlsx_blob({}, {("Sheet1", 1, 1): 1.0})
        )
        chart = Chart(element("c:chartSpace"), None)
        with pytest.raises(ValueError, match="single-cell"):
            update_embedded_xlsx_cell(chart, "Sheet1", "Sheet1!B2", 1)

    def it_does_not_truncate_other_cells_in_a_range_cache(
        self, request, workbook_prop_, workbook_
    ):
        from pptx.chart.chart import update_embedded_xlsx_cell

        # --- range B2:B4, updating only B3 (idx=1) ---
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea><c:barChart><c:ser>"
            "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache>"
            '<c:ptCount val="3"/>'
            '<c:pt idx="0"><c:v>1</c:v></c:pt>'
            '<c:pt idx="1"><c:v>2</c:v></c:pt>'
            '<c:pt idx="2"><c:v>3</c:v></c:pt>'
            "</c:numCache></c:numRef></c:val>"
            "</c:ser></c:barChart></c:plotArea></c:chart>"
            "</c:chartSpace>"
        )
        chartSpace = parse_xml(chartSpace_xml)
        chart = Chart(chartSpace, None)
        workbook_.xlsx_part = _XlsxPartStub(
            _build_update_cached_xlsx_blob(
                {},
                {
                    ("Sheet1", 2, 2): 1.0,
                    ("Sheet1", 3, 2): 2.0,
                    ("Sheet1", 4, 2): 3.0,
                },
            )
        )

        update_embedded_xlsx_cell(chart, "Sheet1", "B3", 99)

        C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        pts = {pt.get("idx"): pt.find(f"{C}v").text for pt in chartSpace.iter(f"{C}pt")}
        # -- other idx values are preserved --
        assert pts["0"] == "1"
        assert pts["1"] == "99"
        assert pts["2"] == "3"

    # -- replace_data_preserve_formulas (issue #239) -----------------

    def it_refreshes_data_cells_but_skips_formula_cells_239(
        self, request, workbook_prop_, workbook_
    ):
        """Data cells get new values; formula cells are left untouched."""
        from pptx.chart.chart import Chart as _Chart  # local to use property_mock

        # -- chart: 1 series, 2 categories; B3 is a formula that must NOT
        # -- be overwritten when new series values (10, 20) are supplied. --
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea><c:barChart><c:ser>"
            "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>"
            '<c:ptCount val="1"/>'
            '<c:pt idx="0"><c:v>Old</c:v></c:pt>'
            "</c:strCache></c:strRef></c:tx>"
            "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache>"
            '<c:ptCount val="2"/>'
            '<c:pt idx="0"><c:v>CA</c:v></c:pt>'
            '<c:pt idx="1"><c:v>CB</c:v></c:pt>'
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>"
            '<c:ptCount val="2"/>'
            '<c:pt idx="0"><c:v>1</c:v></c:pt>'
            '<c:pt idx="1"><c:v>2</c:v></c:pt>'
            "</c:numCache></c:numRef></c:val>"
            "</c:ser></c:barChart></c:plotArea></c:chart>"
            "</c:chartSpace>"
        )
        chartSpace = parse_xml(chartSpace_xml)
        chart = Chart(chartSpace, None)
        property_mock(
            request, _Chart, "chart_type", return_value=XL_CHART_TYPE.BAR_CLUSTERED
        )
        # -- existing workbook has B3 as a formula =B2*2 (cached value 2). --
        workbook_.xlsx_part = _XlsxPartStub(
            _build_update_cached_xlsx_blob(
                {("Sheet1", 1, 2): "Old", ("Sheet1", 2, 1): "CA", ("Sheet1", 3, 1): "CB"},
                {("Sheet1", 2, 2): 1.0},
                formula_cells={("Sheet1", 3, 2): ("B2*2", 2.0)},
            )
        )

        cd = CategoryChartData()
        cd.categories = ["NewA", "NewB"]
        cd.add_series("NewName", (10.0, 20.0))

        written = chart.replace_data_preserve_formulas(cd)

        # -- 5 cells in layout (A2, A3, B1, B2, B3); B3 is a formula so
        # -- it's skipped, leaving 4 writes. --
        assert written == 4
        # -- B2 got its new numCache value; B3 kept its original cache entry --
        C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        pts = {
            pt.get("idx"): pt.find(f"{C}v").text
            for pt in chartSpace.find(f".//{C}numCache").findall(f"{C}pt")
        }
        assert pts["0"] == "10"
        assert pts["1"] == "2"
        # -- series-name strCache was refreshed --
        strCaches = chartSpace.findall(f".//{C}strCache")
        # first strCache is the series name cell (B1); second is categories
        name_pts = {
            pt.get("idx"): pt.find(f"{C}v").text for pt in strCaches[0].findall(f"{C}pt")
        }
        assert name_pts["0"] == "NewName"
        # -- workbook was updated (bytes differ from input blob) --
        workbook_.update_from_xlsx_blob.assert_called_once()

    def it_raises_when_chart_has_no_embedded_workbook_on_preserve_formulas_239(
        self, request, workbook_prop_, workbook_
    ):
        from pptx.chart.chart import Chart as _Chart

        workbook_.xlsx_part = None
        chart = Chart(element("c:chartSpace/c:chart/c:plotArea"), None)
        property_mock(
            request, _Chart, "chart_type", return_value=XL_CHART_TYPE.BAR_CLUSTERED
        )
        cd = CategoryChartData()
        cd.categories = ["A"]
        cd.add_series("S", (1.0,))

        with pytest.raises(ValueError, match="no embedded workbook to refresh"):
            chart.replace_data_preserve_formulas(cd)

    def it_validates_chart_data_type_on_preserve_formulas_239(
        self, request, workbook_prop_, workbook_
    ):
        """Mismatched ChartData subclass raises ValueError (#396 reuse)."""
        from pptx.chart.chart import Chart as _Chart

        chart = Chart(element("c:chartSpace/c:chart/c:plotArea"), None)
        property_mock(
            request, _Chart, "chart_type", return_value=XL_CHART_TYPE.XY_SCATTER
        )
        # -- passing CategoryChartData to an XY chart is invalid --
        with pytest.raises(ValueError, match=r"XY \(scatter\) chart"):
            chart.replace_data_preserve_formulas(CategoryChartData())
        # -- validation short-circuits before any workbook touch --
        workbook_.update_from_xlsx_blob.assert_not_called()

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=["c:catAx", "c:dateAx", "c:valAx"])
    def category_axis_fixture(self, request, CategoryAxis_, DateAxis_, ValueAxis_):
        ax_tag = request.param
        chartSpace_cxml = "c:chartSpace/c:chart/c:plotArea/%s" % ax_tag
        chartSpace = element(chartSpace_cxml)
        chart = Chart(chartSpace, None)
        AxisCls_ = {
            "c:catAx": CategoryAxis_,
            "c:dateAx": DateAxis_,
            "c:valAx": ValueAxis_,
        }[ax_tag]
        axis_ = AxisCls_.return_value
        xAx = chartSpace.xpath(".//%s" % ax_tag)[0]
        return chart, axis_, AxisCls_, xAx

    @pytest.fixture
    def cat_ax_raise_fixture(self):
        chart = Chart(element("c:chartSpace/c:chart/c:plotArea"), None)
        return chart

    @pytest.fixture(
        params=[
            (
                "c:chartSpace{a:b=c}",
                "c:chartSpace{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr" ")",
            ),
            ("c:chartSpace/c:txPr/a:p", "c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr"),
            (
                "c:chartSpace/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)",
                "c:chartSpace/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)",
            ),
        ]
    )
    def font_fixture(self, request):
        chartSpace_cxml, expected_cxml = request.param
        chartSpace = element(chartSpace_cxml)
        expected_xml = xml(expected_cxml)
        return chartSpace, expected_xml

    @pytest.fixture(
        params=[
            ("c:chartSpace/c:chart", False),
            ("c:chartSpace/c:chart/c:legend", True),
        ]
    )
    def has_legend_get_fixture(self, request):
        chartSpace_cxml, expected_value = request.param
        chart = Chart(element(chartSpace_cxml), None)
        return chart, expected_value

    @pytest.fixture(params=[("c:chartSpace/c:chart", True, "c:chartSpace/c:chart/c:legend")])
    def has_legend_set_fixture(self, request):
        chartSpace_cxml, new_value, expected_chartSpace_cxml = request.param
        chart = Chart(element(chartSpace_cxml), None)
        expected_xml = xml(expected_chartSpace_cxml)
        return chart, new_value, expected_xml

    @pytest.fixture(
        params=[("c:chartSpace/c:chart", False), ("c:chartSpace/c:chart/c:title", True)]
    )
    def has_title_get_fixture(self, request):
        chartSpace_cxml, expected_value = request.param
        chart = Chart(element(chartSpace_cxml), None)
        return chart, expected_value

    @pytest.fixture(
        params=[
            ("c:chart", True, "c:chart/c:title/(c:layout,c:overlay{val=0})"),
            ("c:chart/c:title", True, "c:chart/c:title"),
            ("c:chart/c:title", False, "c:chart/c:autoTitleDeleted{val=1}"),
            ("c:chart", False, "c:chart/c:autoTitleDeleted{val=1}"),
        ]
    )
    def has_title_set_fixture(self, request):
        chart_cxml, new_value, expected_cxml = request.param
        chart = Chart(element("c:chartSpace/%s" % chart_cxml), None)
        expected_xml = xml(expected_cxml)
        return chart, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:chartSpace/c:chart", False),
            ("c:chartSpace/c:chart/c:legend", True),
        ]
    )
    def legend_fixture(self, request, Legend_, legend_):
        chartSpace_cxml, has_legend = request.param
        chartSpace = element(chartSpace_cxml)
        chart = Chart(chartSpace, None)
        expected_value, expected_calls = None, []
        if has_legend:
            expected_value = legend_
            legend_elm = chartSpace.chart.legend
            expected_calls.append(call(legend_elm))
        return chart, Legend_, expected_calls, expected_value

    @pytest.fixture
    def plots_fixture(self, _Plots_, plots_):
        chartSpace = element("c:chartSpace/c:chart/c:plotArea")
        plotArea = chartSpace.xpath("./c:chart/c:plotArea")[0]
        chart = Chart(chartSpace, None)
        return chart, plots_, _Plots_, plotArea

    @pytest.fixture
    def replace_fixture(
        self,
        chart_data_,
        SeriesXmlRewriterFactory_,
        series_rewriter_,
        workbook_,
        workbook_prop_,
    ):
        chartSpace = element("c:chartSpace/c:chart/c:plotArea/c:pieChart")
        chart = Chart(chartSpace, None)
        chart_type = XL_CHART_TYPE.PIE
        xlsx_blob = "fooblob"
        chart_data_.xlsx_blob = xlsx_blob
        return (
            chart,
            chart_data_,
            SeriesXmlRewriterFactory_,
            chart_type,
            series_rewriter_,
            chartSpace,
            workbook_,
            xlsx_blob,
        )

    @pytest.fixture
    def update_cached_fixture(self, request, workbook_prop_, workbook_):
        # -- chartSpace with one numRef and one strRef resolving to Sheet1 --
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea><c:barChart><c:ser>"
            "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>"
            '<c:ptCount val="1"/>'
            '<c:pt idx="0"><c:v>OldName</c:v></c:pt>'
            "</c:strCache></c:strRef></c:tx>"
            "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>"
            "<c:formatCode>0.00</c:formatCode>"
            '<c:ptCount val="2"/>'
            '<c:pt idx="0"><c:v>1</c:v></c:pt>'
            '<c:pt idx="1"><c:v>2</c:v></c:pt>'
            "</c:numCache></c:numRef></c:val>"
            "</c:ser></c:barChart></c:plotArea></c:chart>"
            "</c:chartSpace>"
        )
        chartSpace = parse_xml(chartSpace_xml)
        chart = Chart(chartSpace, None)
        # -- fake xlsx part containing a workbook with the new cell values --
        workbook_.xlsx_part = _XlsxPartStub(
            _build_update_cached_xlsx_blob(
                {("Sheet1", 1, 2): "NewName"},
                {("Sheet1", 2, 2): 10.0, ("Sheet1", 3, 2): 20.5},
            )
        )
        expected_xml = parse_xml(
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea><c:barChart><c:ser>"
            "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>"
            '<c:ptCount val="1"/>'
            '<c:pt idx="0"><c:v>NewName</c:v></c:pt>'
            "</c:strCache></c:strRef></c:tx>"
            "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>"
            "<c:formatCode>0.00</c:formatCode>"
            '<c:ptCount val="2"/>'
            '<c:pt idx="0"><c:v>10</c:v></c:pt>'
            '<c:pt idx="1"><c:v>20.5</c:v></c:pt>'
            "</c:numCache></c:numRef></c:val>"
            "</c:ser></c:barChart></c:plotArea></c:chart>"
            "</c:chartSpace>"
        ).xml
        return chart, expected_xml

    @pytest.fixture
    def series_fixture(self, SeriesCollection_, series_collection_):
        chartSpace = element("c:chartSpace/c:chart/c:plotArea")
        plotArea = chartSpace.xpath(".//c:plotArea")[0]
        chart = Chart(chartSpace, None)
        return chart, SeriesCollection_, plotArea, series_collection_

    @pytest.fixture(params=[("c:chartSpace/c:style{val=42}", 42), ("c:chartSpace", None)])
    def style_get_fixture(self, request):
        chartSpace_cxml, expected_value = request.param
        chart = Chart(element(chartSpace_cxml), None)
        return chart, expected_value

    @pytest.fixture(
        params=[
            ("c:chartSpace", 4, "c:chartSpace/c:style{val=4}"),
            ("c:chartSpace", None, "c:chartSpace"),
            ("c:chartSpace/c:style{val=4}", 2, "c:chartSpace/c:style{val=2}"),
            ("c:chartSpace/c:style{val=4}", None, "c:chartSpace"),
        ]
    )
    def style_set_fixture(self, request):
        chartSpace_cxml, new_value, expected_chartSpace_cxml = request.param
        chart = Chart(element(chartSpace_cxml), None)
        expected_xml = xml(expected_chartSpace_cxml)
        return chart, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:chartSpace/c:chart", "c:title/(c:layout,c:overlay{val=0})"),
            ("c:chartSpace/c:chart/c:title/c:layout", "c:title/c:layout"),
        ]
    )
    def title_fixture(self, request, ChartTitle_, chart_title_):
        chartSpace_cxml, expected_cxml = request.param
        chart = Chart(element(chartSpace_cxml), None)
        expected_xml = xml(expected_cxml)
        return chart, expected_xml, ChartTitle_, chart_title_

    @pytest.fixture(
        params=[
            ("c:chartSpace/c:chart/c:plotArea/(c:catAx,c:valAx)", 0),
            ("c:chartSpace/c:chart/c:plotArea/(c:valAx,c:valAx)", 1),
        ]
    )
    def val_ax_fixture(self, request, ValueAxis_, value_axis_):
        chartSpace_xml, idx = request.param
        chartSpace = element(chartSpace_xml)
        chart = Chart(chartSpace, None)
        valAx = chartSpace.xpath(".//c:valAx")[idx]
        return chart, ValueAxis_, valAx, value_axis_

    @pytest.fixture
    def val_ax_raise_fixture(self):
        chart = Chart(element("c:chartSpace/c:chart/c:plotArea"), None)
        return chart

    @pytest.fixture(
        params=[
            # -- no value axes at all -> no secondary --
            ("c:chartSpace/c:chart/c:plotArea", False),
            # -- single value axis, no category axis (single val chart) -> no secondary --
            ("c:chartSpace/c:chart/c:plotArea/c:valAx", False),
            # -- single value axis, with category axis -> no secondary --
            ("c:chartSpace/c:chart/c:plotArea/(c:catAx,c:valAx)", False),
            # -- two value axes, no category axis (XY/scatter) -> no secondary --
            ("c:chartSpace/c:chart/c:plotArea/(c:valAx,c:valAx)", False),
            # -- two value axes with category axis -> has secondary --
            ("c:chartSpace/c:chart/c:plotArea/(c:catAx,c:valAx,c:valAx)", True),
            # -- two value axes with date axis -> has secondary --
            ("c:chartSpace/c:chart/c:plotArea/(c:dateAx,c:valAx,c:valAx)", True),
        ]
    )
    def has_sec_val_ax_fixture(self, request):
        chartSpace_cxml, expected_value = request.param
        chart = Chart(element(chartSpace_cxml), None)
        return chart, expected_value

    @pytest.fixture(
        params=[
            "c:chartSpace/c:chart/c:plotArea/(c:catAx,c:valAx,c:valAx)",
            "c:chartSpace/c:chart/c:plotArea/(c:dateAx,c:valAx,c:valAx)",
        ]
    )
    def sec_val_ax_fixture(self, request, ValueAxis_, value_axis_):
        chartSpace_cxml = request.param
        chartSpace = element(chartSpace_cxml)
        chart = Chart(chartSpace, None)
        sec_valAx = chartSpace.xpath(".//c:valAx")[1]
        return chart, ValueAxis_, sec_valAx, value_axis_

    @pytest.fixture(
        params=[
            # -- no value axes --
            "c:chartSpace/c:chart/c:plotArea",
            # -- only a primary value axis --
            "c:chartSpace/c:chart/c:plotArea/(c:catAx,c:valAx)",
            # -- XY chart (two valAx, no catAx) --
            "c:chartSpace/c:chart/c:plotArea/(c:valAx,c:valAx)",
        ]
    )
    def sec_val_ax_raise_fixture(self, request):
        chartSpace_cxml = request.param
        return Chart(element(chartSpace_cxml), None)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def CategoryAxis_(self, request, category_axis_):
        return class_mock(request, "pptx.chart.chart.CategoryAxis", return_value=category_axis_)

    @pytest.fixture
    def category_axis_(self, request):
        return instance_mock(request, CategoryAxis)

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, ChartData)

    @pytest.fixture
    def ChartTitle_(self, request, chart_title_):
        return class_mock(request, "pptx.chart.chart.ChartTitle", return_value=chart_title_)

    @pytest.fixture
    def chart_title_(self, request):
        return instance_mock(request, ChartTitle)

    @pytest.fixture
    def DateAxis_(self, request, date_axis_):
        return class_mock(request, "pptx.chart.chart.DateAxis", return_value=date_axis_)

    @pytest.fixture
    def date_axis_(self, request):
        return instance_mock(request, DateAxis)

    @pytest.fixture
    def Font_(self, request):
        return class_mock(request, "pptx.chart.chart.Font")

    @pytest.fixture
    def font_(self, request):
        return instance_mock(request, Font)

    @pytest.fixture
    def Legend_(self, request, legend_):
        return class_mock(request, "pptx.chart.chart.Legend", return_value=legend_)

    @pytest.fixture
    def legend_(self, request):
        return instance_mock(request, Legend)

    @pytest.fixture
    def PlotTypeInspector_(self, request):
        return class_mock(request, "pptx.chart.chart.PlotTypeInspector")

    @pytest.fixture
    def _Plots_(self, request, plots_):
        return class_mock(request, "pptx.chart.chart._Plots", return_value=plots_)

    @pytest.fixture
    def plot_(self, request):
        return instance_mock(request, _BasePlot)

    @pytest.fixture
    def plots_(self, request):
        return instance_mock(request, _Plots)

    @pytest.fixture
    def SeriesCollection_(self, request, series_collection_):
        return class_mock(
            request,
            "pptx.chart.chart.SeriesCollection",
            return_value=series_collection_,
        )

    @pytest.fixture
    def SeriesXmlRewriterFactory_(self, request, series_rewriter_):
        return function_mock(
            request,
            "pptx.chart.chart.SeriesXmlRewriterFactory",
            return_value=series_rewriter_,
            autospec=True,
        )

    @pytest.fixture
    def series_collection_(self, request):
        return instance_mock(request, SeriesCollection)

    @pytest.fixture
    def series_rewriter_(self, request):
        return instance_mock(request, _BaseSeriesXmlRewriter)

    @pytest.fixture
    def ValueAxis_(self, request, value_axis_):
        return class_mock(request, "pptx.chart.chart.ValueAxis", return_value=value_axis_)

    @pytest.fixture
    def value_axis_(self, request):
        return instance_mock(request, ValueAxis)

    @pytest.fixture
    def workbook_(self, request):
        return instance_mock(request, ChartWorkbook)

    @pytest.fixture
    def workbook_prop_(self, request, workbook_):
        return property_mock(request, Chart, "_workbook", return_value=workbook_)


class DescribeChartTitle(object):
    """Unit-test suite for `pptx.chart.chart.ChartTitle` objects."""

    def it_provides_access_to_its_format(self, format_fixture):
        chart_title, ChartFormat_, format_ = format_fixture
        format = chart_title.format
        ChartFormat_.assert_called_once_with(chart_title.element)
        assert format is format_

    def it_knows_whether_it_has_a_text_frame(self, has_tf_get_fixture):
        chart_title, expected_value = has_tf_get_fixture
        value = chart_title.has_text_frame
        assert value is expected_value

    def it_can_change_whether_it_has_a_text_frame(self, has_tf_set_fixture):
        chart_title, value, expected_xml = has_tf_set_fixture
        chart_title.has_text_frame = value
        assert chart_title._element.xml == expected_xml

    def it_provides_access_to_its_text_frame(self, text_frame_fixture):
        chart_title, TextFrame_, text_frame_ = text_frame_fixture
        text_frame = chart_title.text_frame
        TextFrame_.assert_called_once_with(chart_title._element.tx.rich, chart_title)
        assert text_frame is text_frame_

    @pytest.mark.parametrize(
        ("title_cxml", "expected_value"),
        [
            ("c:title", None),
            ("c:title/c:layout", None),
            ("c:title/c:layout/c:manualLayout", None),
            (
                "c:title/c:layout/c:manualLayout/(c:x{val=0.3},c:y{val=0.4})",
                None,
            ),
            (
                "c:title/c:layout/c:manualLayout/"
                "(c:xMode,c:x{val=0.25},c:y{val=0.5})",
                None,
            ),
            (
                "c:title/c:layout/c:manualLayout/"
                "(c:xMode,c:yMode,c:x{val=0.25},c:y{val=0.5})",
                (0.25, 0.5),
            ),
            (
                "c:title/c:layout/c:manualLayout/"
                "(c:xMode{val=edge},c:yMode,c:x{val=0.1},c:y{val=0.2})",
                None,
            ),
        ],
    )
    def it_knows_its_position(self, title_cxml, expected_value):
        chart_title = ChartTitle(element(title_cxml))
        assert chart_title.position == expected_value

    @pytest.mark.parametrize(
        ("title_cxml", "value", "expected_cxml"),
        [
            (
                "c:title",
                (0.3, 0.5),
                "c:title/c:layout/c:manualLayout/"
                "(c:xMode,c:yMode,c:x{val=0.3},c:y{val=0.5})",
            ),
            (
                "c:title/c:overlay{val=0}",
                (0.0, 0.0),
                "c:title/(c:layout/c:manualLayout/"
                "(c:xMode,c:yMode,c:x{val=0.0},c:y{val=0.0}),c:overlay{val=0})",
            ),
            (
                "c:title/c:layout/c:manualLayout/"
                "(c:xMode,c:yMode,c:x{val=0.1},c:y{val=0.2})",
                (0.7, 0.8),
                "c:title/c:layout/c:manualLayout/"
                "(c:xMode,c:yMode,c:x{val=0.7},c:y{val=0.8})",
            ),
            (
                "c:title/c:layout/c:manualLayout/"
                "(c:xMode,c:yMode,c:x{val=0.1},c:y{val=0.2})",
                None,
                "c:title",
            ),
            ("c:title", None, "c:title"),
        ],
    )
    def it_can_change_its_position(self, title_cxml, value, expected_cxml):
        chart_title = ChartTitle(element(title_cxml))
        chart_title.position = value
        assert chart_title._element.xml == xml(expected_cxml)

    def it_raises_on_invalid_position_assignment(self):
        chart_title = ChartTitle(element("c:title"))
        with pytest.raises(ValueError):
            chart_title.position = (0.5,)
        with pytest.raises(ValueError):
            chart_title.position = 0.5

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def format_fixture(self, request, ChartFormat_, format_):
        chart_title = ChartTitle(element("c:title"))
        return chart_title, ChartFormat_, format_

    @pytest.fixture(
        params=[
            ("c:title", False),
            ("c:title/c:tx", False),
            ("c:title/c:tx/c:strRef", False),
            ("c:title/c:tx/c:rich", True),
        ]
    )
    def has_tf_get_fixture(self, request):
        title_cxml, expected_value = request.param
        chart_title = ChartTitle(element(title_cxml))
        return chart_title, expected_value

    @pytest.fixture(
        params=[
            (
                "c:title{a:b=c}",
                True,
                "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr" ")",
            ),
            (
                "c:title{a:b=c}/c:tx",
                True,
                "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr" ")",
            ),
            (
                "c:title{a:b=c}/c:tx/c:strRef",
                True,
                "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr" ")",
            ),
            ("c:title/c:tx/c:rich", True, "c:title/c:tx/c:rich"),
            ("c:title", False, "c:title"),
            ("c:title/c:tx", False, "c:title"),
            ("c:title/c:tx/c:rich", False, "c:title"),
            ("c:title/c:tx/c:strRef", False, "c:title"),
        ]
    )
    def has_tf_set_fixture(self, request):
        title_cxml, value, expected_cxml = request.param
        chart_title = ChartTitle(element(title_cxml))
        expected_xml = xml(expected_cxml)
        return chart_title, value, expected_xml

    @pytest.fixture
    def text_frame_fixture(self, request, TextFrame_):
        chart_title = ChartTitle(element("c:title"))
        text_frame_ = TextFrame_.return_value
        return chart_title, TextFrame_, text_frame_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartFormat_(self, request, format_):
        return class_mock(request, "pptx.chart.chart.ChartFormat", return_value=format_)

    @pytest.fixture
    def format_(self, request):
        return instance_mock(request, ChartFormat)

    @pytest.fixture
    def TextFrame_(self, request):
        return class_mock(request, "pptx.chart.chart.TextFrame")


class DescribePlotArea(object):
    """Unit-test suite for `pptx.chart.chart.PlotArea` objects (issue #298)."""

    def it_provides_access_to_its_format(self, request):
        """`PlotArea.format` returns a `ChartFormat` on the `c:plotArea` element."""
        chart_format_ = instance_mock(request, ChartFormat)
        ChartFormat_ = class_mock(
            request, "pptx.chart.chart.ChartFormat", return_value=chart_format_
        )
        plotArea = element("c:plotArea")
        plot_area = PlotArea(plotArea)

        chart_format = plot_area.format

        ChartFormat_.assert_called_once_with(plotArea)
        assert chart_format is chart_format_

    def it_inserts_spPr_in_correct_schema_order_for_fill(self):
        """Regression for #298.

        The CT_PlotArea sequence places ``c:spPr`` between ``c:dTable`` and
        ``c:extLst``. Adding a ``c:spPr`` via ``PlotArea.format`` must not
        corrupt the tree when an ``c:extLst`` is already present.
        """
        plot_area = PlotArea(element("c:plotArea/(c:barChart,c:catAx,c:valAx,c:extLst)"))

        spPr = plot_area.format._element.get_or_add_spPr()

        assert plot_area._element.xml == xml(
            "c:plotArea/(c:barChart,c:catAx,c:valAx,c:spPr,c:extLst)"
        )
        assert spPr is plot_area._element.xpath("c:spPr")[0]


class Describe_Plots(object):
    """Unit-test suite for `pptx.chart.chart._Plots` objects."""

    def it_supports_indexed_access(self, getitem_fixture):
        plots, idx, PlotFactory_, plot_elm, chart_, plot_ = getitem_fixture
        plot = plots[idx]
        PlotFactory_.assert_called_once_with(plot_elm, chart_)
        assert plot is plot_

    def it_supports_len(self, len_fixture):
        plots, expected_len = len_fixture
        assert len(plots) == expected_len

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:plotArea/c:barChart", 0),
            ("c:plotArea/(c:radarChart,c:barChart)", 1),
        ]
    )
    def getitem_fixture(self, request, PlotFactory_, chart_, plot_):
        plotArea_cxml, idx = request.param
        plotArea = element(plotArea_cxml)
        plot_elm = plotArea[idx]
        plots = _Plots(plotArea, chart_)
        return plots, idx, PlotFactory_, plot_elm, chart_, plot_

    @pytest.fixture(
        params=[
            ("c:plotArea", 0),
            ("c:plotArea/c:barChart", 1),
            ("c:plotArea/(c:barChart,c:lineChart)", 2),
        ]
    )
    def len_fixture(self, request):
        plotArea_cxml, expected_len = request.param
        plots = _Plots(element(plotArea_cxml), None)
        return plots, expected_len

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def PlotFactory_(self, request, plot_):
        return function_mock(request, "pptx.chart.chart.PlotFactory", return_value=plot_)

    @pytest.fixture
    def plot_(self, request):
        return instance_mock(request, _BasePlot)


class _XlsxPartStub(object):
    """Fake `EmbeddedXlsxPart` exposing only the `blob` attribute used here."""

    def __init__(self, blob):
        self.blob = blob


def _build_update_cached_xlsx_blob(string_cells, number_cells, formula_cells=None):
    """Return a minimal xlsx blob containing the given cells.

    `string_cells` maps (sheet, row, col) -> str; values go through
    sharedStrings. `number_cells` maps (sheet, row, col) -> float.
    `formula_cells` (optional) maps (sheet, row, col) -> (formula, cached);
    each cell is written with a ``<f>formula</f><v>cached</v>`` body so the
    ``WorkbookReader`` flags it as a formula cell.
    """
    import io as _io
    import zipfile as _zipfile

    # --- collect shared strings and build a cell-address map ---
    shared = []
    sst_idx = {}
    for (_s, _r, _c), v in string_cells.items():
        if v not in sst_idx:
            sst_idx[v] = len(shared)
            shared.append(v)

    def _addr(row, col):
        letters = ""
        n = col
        while n:
            rem = (n - 1) % 26
            letters = chr(ord("A") + rem) + letters
            n = (n - 1) // 26
        return "%s%d" % (letters, row)

    # --- group by row for the sheet xml ---
    rows = {}
    for (_s, r, c), v in string_cells.items():
        rows.setdefault(r, []).append(
            '<c r="%s" t="s"><v>%d</v></c>' % (_addr(r, c), sst_idx[v])
        )
    for (_s, r, c), v in number_cells.items():
        rows.setdefault(r, []).append('<c r="%s"><v>%r</v></c>' % (_addr(r, c), v))
    for (_s, r, c), (formula, cached) in (formula_cells or {}).items():
        rows.setdefault(r, []).append(
            '<c r="%s"><f>%s</f><v>%r</v></c>' % (_addr(r, c), formula, cached)
        )

    row_xml = "".join(
        '<row r="%d">%s</row>' % (r, "".join(cells)) for r, cells in sorted(rows.items())
    )
    sheet_xml = (
        '<?xml version="1.0"?>'
        '<worksheet xmlns='
        '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>" + row_xml + "</sheetData></worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels_xml = (
        '<?xml version="1.0"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    buf = _io.BytesIO()
    with _zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        if shared:
            sst_xml = (
                '<?xml version="1.0"?>'
                '<sst xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                + "".join("<si><t>%s</t></si>" % s for s in shared)
                + "</sst>"
            )
            z.writestr("xl/sharedStrings.xml", sst_xml)
    return buf.getvalue()
