"""Gherkin step implementations for chart features."""

from __future__ import annotations

import hashlib
from itertools import islice

from behave import given, then, when
from helpers import count, test_pptx

from pptx import Presentation
from pptx.chart.chart import Legend
from pptx.chart.data import BubbleChartData, CategoryChartData, ChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_DISPLAY_BLANKS_AS
from pptx.parts.embeddedpackage import EmbeddedXlsxPart
from pptx.util import Inches

# given ===================================================


@given("a Chart object as chart")
def given_a_Chart_object_as_chart(context):
    slide = Presentation(test_pptx("shp-common-props")).slides[0]
    context.chart = slide.shapes[6].chart


@given("a chart having {a_or_no} title")
def given_a_chart_having_a_or_no_title(context, a_or_no):
    shape_idx = {"no": 0, "a": 1}[a_or_no]
    prs = Presentation(test_pptx("cht-chart-props"))
    context.chart = prs.slides[0].shapes[shape_idx].chart


@given("a chart {having_or_not} a legend")
def given_a_chart_having_or_not_a_legend(context, having_or_not):
    slide_idx = {"having": 0, "not having": 1}[having_or_not]
    prs = Presentation(test_pptx("cht-legend"))
    context.chart = prs.slides[slide_idx].shapes[0].chart


@given("a chart whose embedded workbook relationship is missing")
def given_a_chart_with_broken_externalData_rId(context):
    # --- regression scenario for issue #490: build a normal chart, then delete
    # --- the `r:Package` relationship from the chart-part's rels without clearing
    # --- the `c:externalData` element that references it. This mirrors the state
    # --- seen when a chart is pasted from a pre-2007 `.xls` source or when another
    # --- client strips the embedded-workbook rel while leaving the chart XML intact.
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["Old-A", "Old-B", "Old-C"]
    chart_data.add_series("Old-Series", (1.0, 2.0, 3.0))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    ).chart

    # --- drop the embedded-workbook rel from the chart part so the rId referenced
    # --- by c:externalData no longer resolves; simulates the corruption in #490. ---
    chart_part = chart.part
    stale_rId = chart_part._element.xlsx_part_rId
    assert stale_rId is not None
    del chart_part._rels._rels[stale_rId]
    # --- `c:externalData` still points at the now-stale rId; this is the broken state
    assert chart_part._element.xlsx_part_rId == stale_rId
    context.chart = chart


@given("a chart of size and type {spec}")
def given_a_chart_of_size_and_type_spec(context, spec):
    slide_idx = {
        "2x2 Clustered Bar": 0,
        "2x2 100% Stacked Bar": 1,
        "2x2 Clustered Column": 2,
        "4x3 Line": 3,
        "3x1 Pie": 4,
        "3x2 XY": 5,
        "3x2 Bubble": 6,
    }[spec]
    prs = Presentation(test_pptx("cht-replace-data"))
    chart = prs.slides[slide_idx].shapes[0].chart
    context.chart = chart
    context.xlsx_sha1 = hashlib.sha1(chart._workbook.xlsx_part.blob).hexdigest()


@given("a chart of type {chart_type}")
def given_a_chart_of_type_chart_type(context, chart_type):
    slide_idx, shape_idx = {
        "Area": (0, 0),
        "Stacked Area": (0, 1),
        "100% Stacked Area": (0, 2),
        "3-D Area": (0, 3),
        "3-D Stacked Area": (0, 4),
        "3-D 100% Stacked Area": (0, 5),
        "Clustered Bar": (1, 0),
        "Stacked Bar": (1, 1),
        "100% Stacked Bar": (1, 2),
        "Clustered Column": (1, 3),
        "Stacked Column": (1, 4),
        "100% Stacked Column": (1, 5),
        "Line": (2, 0),
        "Stacked Line": (2, 1),
        "100% Stacked Line": (2, 2),
        "Marked Line": (2, 3),
        "Stacked Marked Line": (2, 4),
        "100% Stacked Marked Line": (2, 5),
        "Pie": (3, 0),
        "Exploded Pie": (3, 1),
        "XY (Scatter)": (4, 0),
        "XY Lines": (4, 1),
        "XY Lines No Markers": (4, 2),
        "XY Smooth Lines": (4, 3),
        "XY Smooth No Markers": (4, 4),
        "Bubble": (5, 0),
        "3D-Bubble": (5, 1),
        "Radar": (6, 0),
        "Marked Radar": (6, 1),
        "Filled Radar": (6, 2),
        "Line (with date categories)": (7, 0),
    }[chart_type]
    prs = Presentation(test_pptx("cht-chart-type"))
    context.chart = prs.slides[slide_idx].shapes[shape_idx].chart


@given("a chart title")
def given_a_chart_title(context):
    prs = Presentation(test_pptx("cht-chart-props"))
    context.chart_title = prs.slides[0].shapes[1].chart.chart_title


@given("a chart with an explicitly-colored series")
def given_a_chart_with_an_explicitly_colored_series(context):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("Series 1", (1.1, 2.2, 3.3))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(8),
        Inches(5),
        chart_data,
    ).chart
    # -- paint the sole series with an explicit sRGB color to simulate a template chart
    # -- whose series has been "hard-coded" to a particular color (the cloning-the-bug
    # -- case behind issue #529). --
    fill = chart.series[0].format.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
    context.chart = chart


@given("a chart title having {a_or_no} text frame")
def given_a_chart_title_having_a_or_no_text_frame(context, a_or_no):
    prs = Presentation(test_pptx("cht-chart-props"))
    shape_idx = {"no": 0, "a": 1}[a_or_no]
    context.chart_title = prs.slides[1].shapes[shape_idx].chart.chart_title


# -- issue #338 combo-chart scenario steps ---------------------------------


@given("a column chart and new line-plot chart data")
def given_a_column_chart_and_new_line_plot_chart_data(context):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3"]
    chart_data.add_series("Revenue", (11.0, 22.0, 33.0))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(8),
        Inches(5),
        chart_data,
    ).chart
    context.chart = chart
    line_data = CategoryChartData()
    line_data.categories = ["Q1", "Q2", "Q3"]
    line_data.add_series("Trend", (10.0, 20.0, 30.0))
    context.line_data = line_data


@when("I call chart.add_plot(XL_CHART_TYPE.LINE, line_data)")
def when_I_call_add_plot_LINE(context):
    context.new_plot = context.chart.add_plot(
        XL_CHART_TYPE.LINE_MARKERS, context.line_data
    )


@then("chart.plots has length 2")
def then_chart_plots_has_length_2(context):
    assert len(context.chart.plots) == 2, (
        "expected 2 plots, got %d" % len(context.chart.plots)
    )


@then("the second plot is a LinePlot")
def then_the_second_plot_is_a_LinePlot(context):
    plot = context.chart.plots[1]
    assert type(plot).__name__ == "LinePlot", (
        "expected LinePlot, got %s" % type(plot).__name__
    )


@then("the line plot's axId elements match the column plot's")
def then_axId_match(context):
    from pptx.oxml.ns import qn

    xCharts = context.chart._chartSpace.plotArea.xCharts
    col_ids = [e.get("val") for e in xCharts[0].findall(qn("c:axId"))]
    line_ids = [e.get("val") for e in xCharts[1].findall(qn("c:axId"))]
    assert col_ids == line_ids and len(col_ids) == 2, (
        "axId mismatch: col=%r line=%r" % (col_ids, line_ids)
    )


@then("the line plot's series use inline numLit and strLit data")
def then_line_plot_uses_numLit_strLit(context):
    from pptx.oxml.ns import qn

    line_xChart = context.chart._chartSpace.plotArea.xCharts[1]
    # -- c:val uses numLit (literal values, no workbook range ref) --
    assert line_xChart.find(".//" + qn("c:val") + "/" + qn("c:numLit")) is not None
    assert line_xChart.find(".//" + qn("c:val") + "/" + qn("c:numRef")) is None
    # -- c:cat uses strLit (literal labels, no workbook range ref) --
    assert line_xChart.find(".//" + qn("c:cat") + "/" + qn("c:strLit")) is not None
    assert line_xChart.find(".//" + qn("c:cat") + "/" + qn("c:strRef")) is None


@given("a {axis_config}-axes chart")
def given_a_axis_config_axes_chart(context, axis_config):
    # -- build a basic column chart and optionally inject a secondary value axis --
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["East", "West", "Midwest"]
    chart_data.add_series("Primary", (1.1, 2.2, 3.3))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(6), Inches(4), chart_data
    ).chart

    if axis_config == "primary-secondary":
        _add_secondary_value_axis(chart)
    elif axis_config == "single-value":
        pass
    else:
        raise KeyError("unrecognized axes-config: %s" % axis_config)

    context.chart = chart


def _add_secondary_value_axis(chart):
    """Inject a secondary value axis into `chart` by appending XML.

    This mirrors the pattern PowerPoint writes for a category chart with a
    secondary value axis: a hidden secondary catAx (axPos=t) and a visible
    secondary valAx (axPos=r), each crossing the other via `c:crossAx`.
    """
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import _nsmap

    plotArea = chart._chartSpace.plotArea
    sec_val_id = 1000001
    sec_cat_id = 1000002
    ns = ' xmlns:c="%s"' % _nsmap["c"]

    sec_catAx_xml = (
        "<c:catAx%s>"
        '<c:axId val="%d"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="1"/>'
        '<c:axPos val="t"/>'
        '<c:crossAx val="%d"/>'
        "</c:catAx>"
    ) % (ns, sec_cat_id, sec_val_id)
    sec_valAx_xml = (
        "<c:valAx%s>"
        '<c:axId val="%d"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="r"/>'
        '<c:crossAx val="%d"/>'
        "</c:valAx>"
    ) % (ns, sec_val_id, sec_cat_id)

    plotArea.append(parse_xml(sec_catAx_xml))
    plotArea.append(parse_xml(sec_valAx_xml))
@given("a chart with an extended c14:style wrapped in mc:AlternateContent")
def given_a_chart_with_an_extended_c14_style(context):
    # -- `cht-chart-props.pptx` charts are authored by PowerPoint with the extended
    # -- `c14:style val="118"` wrapped inside `mc:AlternateContent` and a `c:style val="18"`
    # -- fallback. Before issue #516, `Chart.chart_style` returned `None` for this layout
    # -- because it only looked at the plain `c:style` child.
    prs = Presentation(test_pptx("cht-chart-props"))
    context.chart = prs.slides[0].shapes[0].chart


@given("a chart with no explicit chart style")
def given_a_chart_with_no_explicit_chart_style(context):
    # -- an `add_chart`-constructed column chart has no style child at all --
    from pptx.util import Inches as _In

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("S1", (1, 2, 3))
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, _In(1), _In(1), _In(4), _In(3), chart_data
    )
    context.chart = gf.chart


@given("a chart authored by add_chart (default orientation)")
def given_a_chart_authored_by_add_chart(context):
    # -- `add_chart` writes series-in-columns (default orientation) — the ---
    # -- state python-pptx produces when building a chart from            ---
    # -- CategoryChartData. Used to pin series_in_rows == False (#828).   ---
    from pptx.util import Inches as _In

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["East", "West", "Mid"]
    chart_data.add_series("Q1", (1, 2, 3))
    chart_data.add_series("Q2", (4, 5, 6))
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, _In(1), _In(1), _In(4), _In(3), chart_data
    )
    context.chart = gf.chart


@given("a chart whose XML has been flipped to row-series orientation")
def given_a_chart_flipped_to_row_series(context):
    # -- Simulates the XML PowerPoint writes after a user applies          ---
    # -- "Chart Design > Switch Row/Column". Each c:ser references a       ---
    # -- single-cell series name in column A (row >1), a c:cat range that  ---
    # -- runs horizontally across row 1, and a c:val range in the same     ---
    # -- row as its series name. Built on top of an add_chart-authored     ---
    # -- chart so everything else (part/rel plumbing) is real.             ---
    from pptx.oxml.ns import qn as _qn
    from pptx.util import Inches as _In

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["East", "West", "Mid"]
    chart_data.add_series("Q1", (1, 2, 3))
    chart_data.add_series("Q2", (4, 5, 6))
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, _In(1), _In(1), _In(4), _In(3), chart_data
    )
    chart = gf.chart

    # --- rewrite c:f references to simulate the switch-row/col layout ---
    sers = chart._chartSpace.plotArea.findall(".//" + _qn("c:ser"))
    # -- after switching, series live in row 2.., categories span across cols B.. --
    # -- Q1 -> row 2 (name in A2, values in B2:D2); Q2 -> row 3. --
    for ser_idx, ser in enumerate(sers):
        tx = ser.find(_qn("c:tx"))
        if tx is not None:
            strRef = tx.find(_qn("c:strRef"))
            if strRef is not None:
                strRef.find(_qn("c:f")).text = "Sheet1!$A$%d" % (ser_idx + 2)
        cat = ser.find(_qn("c:cat"))
        if cat is not None:
            strRef = cat.find(_qn("c:strRef"))
            if strRef is not None:
                strRef.find(_qn("c:f")).text = "Sheet1!$B$1:$D$1"
        val = ser.find(_qn("c:val"))
        if val is not None:
            numRef = val.find(_qn("c:numRef"))
            if numRef is not None:
                numRef.find(_qn("c:f")).text = "Sheet1!$B$%d:$D$%d" % (
                    ser_idx + 2, ser_idx + 2
                )
    context.chart = chart


# when ====================================================


@when("I add a Clustered bar chart with multi-level categories")
def when_I_add_a_clustered_bar_chart_with_multi_level_categories(context):
    chart_type = XL_CHART_TYPE.BAR_CLUSTERED
    chart_data = CategoryChartData()

    WEST = chart_data.add_category("WEST")
    WEST.add_sub_category("SF")
    WEST.add_sub_category("LA")
    EAST = chart_data.add_category("EAST")
    EAST.add_sub_category("NY")
    EAST.add_sub_category("NJ")

    chart_data.add_series("Series 1", (1, 2, None, 4))
    chart_data.add_series("Series 2", (5, None, 7, 8))

    context.chart = context.slide.shapes.add_chart(
        chart_type, Inches(1), Inches(1), Inches(8), Inches(5), chart_data
    ).chart


@when("I add a {kind} chart with {cats} categories and {sers} series")
def when_I_add_a_chart_with_categories_and_series(context, kind, cats, sers):
    chart_type = {
        "Area": XL_CHART_TYPE.AREA,
        "Stacked Area": XL_CHART_TYPE.AREA_STACKED,
        "100% Stacked Area": XL_CHART_TYPE.AREA_STACKED_100,
        "Clustered Bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "Stacked Bar": XL_CHART_TYPE.BAR_STACKED,
        "100% Stacked Bar": XL_CHART_TYPE.BAR_STACKED_100,
        "Clustered Column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "Stacked Column": XL_CHART_TYPE.COLUMN_STACKED,
        "100% Stacked Column": XL_CHART_TYPE.COLUMN_STACKED_100,
        "Doughnut": XL_CHART_TYPE.DOUGHNUT,
        "Exploded Doughnut": XL_CHART_TYPE.DOUGHNUT_EXPLODED,
        "Line": XL_CHART_TYPE.LINE,
        "Line with Markers": XL_CHART_TYPE.LINE_MARKERS,
        "Line Markers Stacked": XL_CHART_TYPE.LINE_MARKERS_STACKED,
        "100% Line Markers Stacked": XL_CHART_TYPE.LINE_MARKERS_STACKED_100,
        "Line Stacked": XL_CHART_TYPE.LINE_STACKED,
        "100% Line Stacked": XL_CHART_TYPE.LINE_STACKED_100,
        "Pie": XL_CHART_TYPE.PIE,
        "Exploded Pie": XL_CHART_TYPE.PIE_EXPLODED,
        "3D Pie": XL_CHART_TYPE.THREE_D_PIE,
        "Exploded 3D Pie": XL_CHART_TYPE.THREE_D_PIE_EXPLODED,
        "Radar": XL_CHART_TYPE.RADAR,
        "Filled Radar": XL_CHART_TYPE.RADAR_FILLED,
        "Radar with markers": XL_CHART_TYPE.RADAR_MARKERS,
    }[kind]
    category_count, series_count = int(cats), int(sers)
    category_source = ("Foo", "Bar", "Baz", "Boo", "Far", "Faz")
    series_value_source = count(1.1, 1.1)

    chart_data = CategoryChartData()
    chart_data.categories = category_source[:category_count]
    for idx in range(series_count):
        series_title = "Series %d" % (idx + 1)
        series_values = tuple(islice(series_value_source, category_count))
        chart_data.add_series(series_title, series_values)

    context.chart = context.slide.shapes.add_chart(
        chart_type, Inches(1), Inches(1), Inches(8), Inches(5), chart_data
    ).chart


@when("I add a {bubble_type} chart having 2 series of 3 points each")
def when_I_add_a_bubble_chart_having_2_series_of_3_pts(context, bubble_type):
    chart_type = getattr(XL_CHART_TYPE, bubble_type)
    data = (
        ("Series 1", ((-0.1, 0.5, 1.0), (16.2, 0.0, 2.0), (8.0, -0.2, 3.0))),
        ("Series 2", ((12.4, 0.8, 4.0), (-7.5, 0.5, 5.0), (5.1, -0.5, 6.0))),
    )

    chart_data = BubbleChartData()

    for series_data in data:
        series_label, points = series_data
        series = chart_data.add_series(series_label)
        for point in points:
            x, y, size = point
            series.add_data_point(x, y, size)

    context.chart = context.slide.shapes.add_chart(
        chart_type, Inches(1), Inches(1), Inches(8), Inches(5), chart_data
    ).chart


@when("I assign {value} to chart.has_legend")
def when_I_assign_value_to_chart_has_legend(context, value):
    new_value = {"True": True, "False": False}[value]
    context.chart.has_legend = new_value


@when("I assign {value} to chart.has_title")
def when_I_assign_value_to_chart_has_title(context, value):
    context.chart.has_title = {"True": True, "False": False}[value]


@when("I assign {value} to chart_title.has_text_frame")
def when_I_assign_value_to_chart_title_has_text_frame(context, value):
    context.chart_title.has_text_frame = {"True": True, "False": False}[value]


@when("I assign {value:d} to chart.chart_style")
def when_I_assign_value_to_chart_chart_style(context, value):
    context.chart.chart_style = value


@when("I assign XL_DISPLAY_BLANKS_AS.{member} to chart.display_blanks_as")
def when_I_assign_display_blanks_as(context, member):
    context.chart.display_blanks_as = getattr(XL_DISPLAY_BLANKS_AS, member)


@when("I replace its data with {cats} categories and {sers} series")
def when_I_replace_its_data_with_categories_and_series(context, cats, sers):
    category_count, series_count = int(cats), int(sers)
    category_source = ("Foo", "Bar", "Baz", "Boo", "Far", "Faz")
    series_value_source = count(1.1, 1.1)

    chart_data = ChartData()
    chart_data.categories = category_source[:category_count]
    for idx in range(series_count):
        series_title = "New Series %d" % (idx + 1)
        series_values = tuple(islice(series_value_source, category_count))
        chart_data.add_series(series_title, series_values)

    context.chart.replace_data(chart_data)


@when("I replace its data with 6 series that require 5 new cloned series")
def when_I_replace_its_data_with_6_series_requiring_clones(context):
    # -- a 1-series chart becomes 6 series; five new sers are cloned from the single existing
    # -- one, which carries an explicit red sRGB fill. After the fix for issue #529 the five
    # -- new sers should adopt theme-accent schemeClr values rather than cloning the red. --
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    for idx in range(6):
        chart_data.add_series("S%d" % (idx + 1), (1.0, 2.0, 3.0))
    context.chart.replace_data(chart_data)


@when("I replace its data with 3 series of 3 bubble points each")
def when_I_replace_its_data_with_3_series_of_three_bubble_pts_each(context):
    chart_data = BubbleChartData()
    for idx in range(3):
        series_title = "New Series %d" % (idx + 1)
        series = chart_data.add_series(series_title)
        for jdx in range(3):
            x, y, size = idx * 3 + jdx, idx * 2 + jdx, idx + jdx
            series.add_data_point(x, y, size)

    context.chart.replace_data(chart_data)


@when("I replace its data with 3 series of 3 points each")
def when_I_replace_its_data_with_3_series_of_three_points_each(context):
    chart_data = XyChartData()
    x = y = 0
    for idx in range(3):
        series_title = "New Series %d" % (idx + 1)
        series = chart_data.add_series(series_title)
        for jdx in range(3):
            x, y = idx * 3 + jdx, idx * 2 + jdx
            series.add_data_point(x, y)

    context.chart.replace_data(chart_data)


# then ====================================================


@then("replacing its data with {wrong_kind} raises ValueError")
def then_replacing_its_data_with_wrong_kind_raises_ValueError(context, wrong_kind):
    """Verify Chart.replace_data() rejects a mismatched ChartData type (issue #396).

    *wrong_kind* selects which deliberately-incompatible ChartData subclass to
    build; the resulting instance is minimally-populated so the validation
    check is what trips the error (no xlsx-writing happens).
    """
    if wrong_kind == "CategoryData":
        chart_data = CategoryChartData()
        chart_data.categories = ["A", "B"]
        chart_data.add_series("S1", (1.0, 2.0))
    elif wrong_kind == "XyData":
        chart_data = XyChartData()
        series = chart_data.add_series("S1")
        series.add_data_point(1, 2)
    elif wrong_kind == "BubbleData":
        chart_data = BubbleChartData()
        series = chart_data.add_series("S1")
        series.add_data_point(1, 2, 3)
    else:  # pragma: no cover - guard against feature-file typos
        raise ValueError("unknown wrong_kind '%s'" % wrong_kind)

    try:
        context.chart.replace_data(chart_data)
    except ValueError:
        return
    raise AssertionError(
        "Chart.replace_data() did not raise ValueError for wrong_kind=%r" % wrong_kind
    )


@then("chart.category_axis is a {cls_name} object")
def then_chart_category_axis_is_a_cls_name_object(context, cls_name):
    category_axis = context.chart.category_axis
    type_name = type(category_axis).__name__
    assert type_name == cls_name, "got %s" % type_name


@then("chart.chart_title is a ChartTitle object")
def then_chart_chart_title_is_a_ChartTitle_object(context):
    class_name = type(context.chart.chart_title).__name__
    assert class_name == "ChartTitle", "got %s" % class_name


@then("chart.chart_style is {value:d}")
def then_chart_chart_style_is_value(context, value):
    actual = context.chart.chart_style
    assert actual == value, "got %r" % actual


@then("chart.display_blanks_as is XL_DISPLAY_BLANKS_AS.{member}")
def then_chart_display_blanks_as_is(context, member):
    expected = getattr(XL_DISPLAY_BLANKS_AS, member)
    actual = context.chart.display_blanks_as
    assert actual == expected, "got %r" % actual


@then("chart.series_in_rows is {value}")
def then_chart_series_in_rows_is(context, value):
    expected = {"True": True, "False": False, "None": None}[value]
    actual = context.chart.series_in_rows
    assert actual is expected, "got %r" % actual


@then("chartSpace has an mc:AlternateContent/mc:Choice/c14:style val={val:d}")
def then_chartSpace_has_c14_style(context, val):
    cs = context.chart._chartSpace
    mc_ns = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
    c14_ns = "{http://schemas.microsoft.com/office/drawing/2007/8/2/chart}"
    for ac in cs.iterchildren(mc_ns + "AlternateContent"):
        for choice in ac.iterchildren(mc_ns + "Choice"):
            c14_style = choice.find(c14_ns + "style")
            if c14_style is not None and int(c14_style.get("val")) == val:
                return
    raise AssertionError(
        "mc:AlternateContent/mc:Choice/c14:style val=%d not found" % val
    )


@then("chartSpace has an mc:AlternateContent/mc:Fallback/c:style val={val:d}")
def then_chartSpace_has_fallback_cstyle(context, val):
    cs = context.chart._chartSpace
    mc_ns = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
    c_ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    for ac in cs.iterchildren(mc_ns + "AlternateContent"):
        fallback = ac.find(mc_ns + "Fallback")
        if fallback is None:
            continue
        style = fallback.find(c_ns + "style")
        if style is not None and int(style.get("val")) == val:
            return
    raise AssertionError(
        "mc:AlternateContent/mc:Fallback/c:style val=%d not found" % val
    )


@then("chart.chart_type is {enum_member}")
def then_chart_chart_type_is_value(context, enum_member):
    expected_value = getattr(XL_CHART_TYPE, enum_member)
    chart = context.chart
    assert chart.chart_type is expected_value, "got %s" % chart.chart_type


@then("chart.font is a Font object")
def then_chart_font_is_a_Font_object(context):
    actual = type(context.chart.font).__name__
    expected = "Font"
    assert actual == expected, "chart.font is a %s object" % actual


@then("chart.plot_area is a PlotArea object")
def then_chart_plot_area_is_a_PlotArea_object(context):
    actual = type(context.chart.plot_area).__name__
    assert actual == "PlotArea", "got %s" % actual


@then("chart.plot_area.format is a ChartFormat object")
def then_chart_plot_area_format_is_a_ChartFormat_object(context):
    actual = type(context.chart.plot_area.format).__name__
    assert actual == "ChartFormat", "got %s" % actual


@then("chart.plot_area.format.fill is a FillFormat object")
def then_chart_plot_area_format_fill_is_a_FillFormat_object(context):
    actual = type(context.chart.plot_area.format.fill).__name__
    assert actual == "FillFormat", "got %s" % actual


@then("chart.plot_area.format.line is a LineFormat object")
def then_chart_plot_area_format_line_is_a_LineFormat_object(context):
    actual = type(context.chart.plot_area.format.line).__name__
    assert actual == "LineFormat", "got %s" % actual


@then("chart.has_legend is {value}")
def then_chart_has_legend_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    chart = context.chart
    assert chart.has_legend is expected_value


@then("chart.has_title is {value}")
def then_chart_has_title_is_value(context, value):
    chart = context.chart
    actual_value = chart.has_title
    expected_value = {"True": True, "False": False}[value]
    assert actual_value is expected_value, "got %s" % actual_value


@then("chart.legend is a legend object")
def then_chart_legend_is_a_legend_object(context):
    chart = context.chart
    assert isinstance(chart.legend, Legend)


@then("chart.series is a SeriesCollection object")
def then_chart_series_is_a_SeriesCollection_object(context):
    type_name = type(context.chart.series).__name__
    assert type_name == "SeriesCollection", "got %s" % type_name


@then("chart.value_axis is a ValueAxis object")
def then_chart_value_axis_is_a_ValueAxis_object(context):
    value_axis = context.chart.value_axis
    assert type(value_axis).__name__ == "ValueAxis"


@then("chart.has_secondary_value_axis is {value}")
def then_chart_has_secondary_value_axis_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    actual_value = context.chart.has_secondary_value_axis
    assert actual_value is expected_value, "got %s" % actual_value


@then("chart.secondary_value_axis is a ValueAxis object")
def then_chart_secondary_value_axis_is_a_ValueAxis_object(context):
    secondary_value_axis = context.chart.secondary_value_axis
    assert type(secondary_value_axis).__name__ == "ValueAxis"


@then("accessing chart.secondary_value_axis raises ValueError")
def then_accessing_chart_secondary_value_axis_raises(context):
    try:
        context.chart.secondary_value_axis
    except ValueError:
        return
    raise AssertionError("ValueError not raised")


@then("chart.primary_value_axis is a ValueAxis object")
def then_chart_primary_value_axis_is_a_ValueAxis_object(context):
    primary_value_axis = context.chart.primary_value_axis
    assert type(primary_value_axis).__name__ == "ValueAxis"


@then("chart.primary_value_axis is the first c:valAx")
def then_chart_primary_value_axis_is_first_valAx(context):
    from pptx.oxml.ns import qn

    valAx_lst = context.chart._chartSpace.findall(".//" + qn("c:valAx"))
    assert valAx_lst, "chart has no c:valAx element"
    primary = context.chart.primary_value_axis
    assert primary._element is valAx_lst[0], (
        "primary_value_axis._element is not the first c:valAx"
    )


@then("chart.primary_value_axis is not chart.value_axis")
def then_chart_primary_value_axis_is_not_chart_value_axis(context):
    assert context.chart.primary_value_axis._element is not context.chart.value_axis._element, (
        "primary_value_axis and value_axis unexpectedly alias the same c:valAx"
    )


@then("chart.value_axis is chart.secondary_value_axis")
def then_chart_value_axis_is_chart_secondary_value_axis(context):
    assert (
        context.chart.value_axis._element is context.chart.secondary_value_axis._element
    ), "value_axis is not aliased to secondary_value_axis"


@then("chart_title.format is a ChartFormat object")
def then_chart_title_format_is_a_ChartFormat_object(context):
    class_name = type(context.chart_title.format).__name__
    assert class_name == "ChartFormat", "got %s" % class_name


@then("chart_title.format.fill is a FillFormat object")
def then_chart_title_format_fill_is_a_FillFormat_object(context):
    class_name = type(context.chart_title.format.fill).__name__
    assert class_name == "FillFormat", "got %s" % class_name


@then("chart_title.format.line is a LineFormat object")
def then_chart_title_format_line_is_a_LineFormat_object(context):
    class_name = type(context.chart_title.format.line).__name__
    assert class_name == "LineFormat", "got %s" % class_name


@then("chart_title.has_text_frame is {value}")
def then_chart_title_has_text_frame_is_value(context, value):
    actual_value = context.chart_title.has_text_frame
    expected_value = {"True": True, "False": False}[value]
    assert actual_value is expected_value, "got %s" % actual_value


@then("chart_title.text_frame is a TextFrame object")
def then_chart_title_text_frame_is_a_TextFrame_object(context):
    class_name = type(context.chart_title.text_frame).__name__
    assert class_name == "TextFrame", "got %s" % class_name


@then("each series has a new name")
def then_each_series_has_a_new_name(context):
    for series in context.chart.plots[0].series:
        assert series.name.startswith("New ")


@then("each cloned series uses a distinct theme-accent schemeClr")
def then_each_cloned_series_uses_distinct_accent(context):
    # -- the original ser (idx=0) keeps its explicit sRGB color; the five clones (idx 1..5)
    # -- should each carry a unique a:schemeClr val="accent{n}" reference, cycling 2..6 --
    sers = context.chart._chartSpace.xpath(".//c:ser")
    assert len(sers) == 6, "expected 6 series, got %d" % len(sers)

    # -- source ser preserves its original explicit sRGB red --
    srgb_vals = sers[0].xpath(".//a:srgbClr/@val")
    assert "FF0000" in srgb_vals, "source ser lost its explicit sRGB color: %r" % srgb_vals

    expected_accents = ["accent2", "accent3", "accent4", "accent5", "accent6"]
    for ser, expected in zip(sers[1:], expected_accents):
        scheme_vals = ser.xpath(".//a:schemeClr/@val")
        assert expected in scheme_vals, (
            "cloned ser missing schemeClr val='%s'; got schemeClr vals=%r, srgb=%r"
            % (expected, scheme_vals, ser.xpath(".//a:srgbClr/@val"))
        )
        # -- no explicit sRGB should survive on the clone --
        assert ser.xpath(".//a:srgbClr") == [], (
            "cloned ser still contains an explicit a:srgbClr: %r"
            % ser.xpath(".//a:srgbClr/@val")
        )


@then("each series has {count} values")
def then_each_series_has_count_values(context, count):
    expected_count = int(count)
    for series in context.chart.plots[0].series:
        actual_value_count = len(series.values)
        assert actual_value_count == expected_count


@then("len(chart.series) is {count}")
def then_len_chart_series_is_count(context, count):
    expected_count = int(count)
    assert len(context.chart.series) == expected_count


@then("the chart has an Excel data worksheet")
def then_the_chart_has_an_Excel_data_worksheet(context):
    xlsx_part = context.chart._workbook.xlsx_part
    assert isinstance(xlsx_part, EmbeddedXlsxPart)


@then("the chart has new chart data")
def then_the_chart_has_new_chart_data(context):
    orig_xlsx_sha1 = context.xlsx_sha1
    new_xlsx_sha1 = hashlib.sha1(context.chart._workbook.xlsx_part.blob).hexdigest()
    assert new_xlsx_sha1 != orig_xlsx_sha1


# --- steps for cht-update-cache.feature ---------------------------------


@given("a chart whose embedded workbook has been updated externally")
def given_a_chart_with_externally_updated_workbook(context):
    import io as _io
    import re as _re
    import zipfile as _zipfile

    # --- build a chart with known initial values/labels ---
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["Old-A", "Old-B", "Old-C"]
    chart_data.add_series("Old-Series", (1.0, 2.0, 3.0))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    ).chart

    # --- externally rewrite the embedded xlsx: replace numeric cells and
    #     shared strings without touching the sheet-level shared-string
    #     index lookups (so this only changes the data, not the structure) ---
    blob = chart._workbook.xlsx_part.blob
    with _zipfile.ZipFile(_io.BytesIO(blob)) as z:
        files = {n: z.read(n) for n in z.namelist()}

    sheet = files["xl/worksheets/sheet1.xml"].decode()
    # Only numeric cells: pattern `s="1"` and no `t="s"`, with numeric `<v>`.
    def _bump(m):
        addr = m.group("addr")
        val = m.group("val")
        try:
            new = float(val) * 10
        except ValueError:
            new = val
        return '<c r="%s" s="1"><v>%s</v></c>' % (addr, new)

    new_sheet = _re.sub(
        r'<c r="(?P<addr>[A-Z]+\d+)" s="1"><v>(?P<val>[^<]+)</v></c>',
        _bump,
        sheet,
    )

    ss = files["xl/sharedStrings.xml"].decode()
    new_ss = (
        ss.replace("<t>Old-A</t>", "<t>New-A</t>")
        .replace("<t>Old-B</t>", "<t>New-B</t>")
        .replace("<t>Old-C</t>", "<t>New-C</t>")
        .replace("<t>Old-Series</t>", "<t>New-Series</t>")
    )
    files["xl/worksheets/sheet1.xml"] = new_sheet.encode()
    files["xl/sharedStrings.xml"] = new_ss.encode()
    buf = _io.BytesIO()
    with _zipfile.ZipFile(buf, "w", _zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    chart._workbook.xlsx_part.blob = buf.getvalue()

    context.chart = chart


@when("I call chart.update_cached_values()")
def when_I_call_chart_update_cached_values(context):
    context.chart.update_cached_values()


@then("the cached series values match the embedded workbook")
def then_cached_series_values_match_the_embedded_workbook(context):
    chart = context.chart
    series = chart.series[0]
    # --- values should now be 10.0, 20.0, 30.0 ---
    assert tuple(series.values) == (10.0, 20.0, 30.0), tuple(series.values)


@then("the cached category labels match the embedded workbook")
def then_cached_category_labels_match_the_embedded_workbook(context):
    chart = context.chart
    labels = [c.label for c in chart.plots[0].categories]
    assert labels == ["New-A", "New-B", "New-C"], labels


@then("the cached series name matches the embedded workbook")
def then_cached_series_name_matches_the_embedded_workbook(context):
    chart = context.chart
    series = chart.series[0]
    assert series.name == "New-Series", series.name


# --- F5: embedded-workbook handler ----------------------------------

@given("a chart with an embedded workbook")
def given_a_chart_with_an_embedded_workbook(context):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("S1", (1.0, 2.0, 3.0))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    ).chart
    context.chart = chart


@then("chart.workbook is the bytes of the embedded xlsx")
def then_chart_workbook_is_bytes_of_embedded_xlsx(context):
    blob = context.chart.workbook
    assert isinstance(blob, bytes), type(blob)
    assert blob == context.chart._workbook.xlsx_part.blob


@then("chart.workbook reads the .xlsx zip magic")
def then_chart_workbook_reads_xlsx_zip_magic(context):
    # --- every .xlsx is a zip; zip files begin with 'PK\x03\x04' ---
    assert context.chart.workbook[:4] == b"PK\x03\x04"


@when("I assign new bytes to chart.workbook")
def when_I_assign_new_bytes_to_chart_workbook(context):
    # --- round-trip the original blob with a trivial rewrite: we
    # --- use the existing xlsx bytes so the chart remains valid,
    # --- then prepend a zip comment-safe byte (simulated by just
    # --- setting identical bytes). A simpler probe is to assign a
    # --- captured snapshot; the getter/setter semantics are the
    # --- test focus here. ---
    context._new_bytes = b"x-new-bytes"
    context.chart.workbook = context._new_bytes


@then("chart.workbook returns the new bytes")
def then_chart_workbook_returns_new_bytes(context):
    assert context.chart.workbook == context._new_bytes


@when('I call update_embedded_xlsx_cell(chart, "Sheet1", "B2", 42.0)')
def when_I_call_update_embedded_xlsx_cell(context):
    from pptx.chart.chart import update_embedded_xlsx_cell

    update_embedded_xlsx_cell(context.chart, "Sheet1", "B2", 42.0)


@then("the workbook's Sheet1!B2 cell reads 42.0")
def then_workbook_B2_reads_42(context):
    from pptx.chart.xlsx import WorkbookReader

    with WorkbookReader(context.chart.workbook) as reader:
        assert reader.cell_value("Sheet1", 2, 2) == 42.0


@then("the first series' first value reads 42.0")
def then_first_series_first_value_reads_42(context):
    assert tuple(context.chart.series[0].values)[0] == 42.0


@given("two charts in the same presentation")
def given_two_charts_in_same_presentation(context):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd_a = CategoryChartData()
    cd_a.categories = ["A", "B"]
    cd_a.add_series("S", (1.0, 2.0))
    chart_a = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(4),
        Inches(3),
        cd_a,
    ).chart
    cd_b = CategoryChartData()
    cd_b.categories = ["X"]
    cd_b.add_series("T", (9.0,))
    chart_b = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(5),
        Inches(4),
        Inches(3),
        cd_b,
    ).chart
    context.source_chart = chart_a
    context.target_chart = chart_b


@when("I call clone_embedded_xlsx(source_chart.part, target_chart.part)")
def when_I_call_clone_embedded_xlsx(context):
    from pptx.parts.embeddedpackage import clone_embedded_xlsx

    context._source_xlsx_part = context.source_chart.part.chart_workbook.xlsx_part
    clone_embedded_xlsx(context.source_chart.part, context.target_chart.part)


@then("target_chart.workbook equals source_chart.workbook")
def then_target_workbook_equals_source(context):
    assert context.target_chart.workbook == context.source_chart.workbook


@then("the target chart's xlsx part is not the source's xlsx part")
def then_target_xlsx_part_differs_from_source(context):
    assert (
        context.target_chart.part.chart_workbook.xlsx_part
        is not context._source_xlsx_part
    )


# --- steps for cht-clone-to.feature (#877 cross-slide chart copy) -------


@given("a chart with an embedded workbook on slide 1")
def given_a_chart_on_slide_1_with_a_second_slide(context):
    prs = Presentation()
    slide1 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    cd.add_series("S1", (10.0, 20.0, 30.0))
    gf = slide1.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1), Inches(1), Inches(5), Inches(3),
        cd,
    )
    context.prs = prs
    context.source_chart = gf.chart
    context.source_gf = gf
    context.slide_2 = slide2
    context._x = Inches(2)
    context._y = Inches(2)
    context._cx = Inches(4)
    context._cy = Inches(3)


@when("I call chart.clone_to(slide_2.shapes, x, y, cx, cy)")
def when_I_call_chart_clone_to_slide_2(context):
    context._src_xlsx_part = context.source_chart.part.chart_workbook.xlsx_part
    context._src_chart_part = context.source_chart.part
    new_gf = context.source_chart.clone_to(
        context.slide_2.shapes, context._x, context._y, context._cx, context._cy
    )
    context.new_gf = new_gf
    context.cloned_chart = new_gf.chart


@then("slide 2 has a chart shape at (x, y) sized (cx, cy)")
def then_slide_2_has_chart_shape(context):
    gf = context.new_gf
    assert gf.has_chart
    assert gf.left == context._x
    assert gf.top == context._y
    assert gf.width == context._cx
    assert gf.height == context._cy


@then("the cloned chart has a distinct chart part")
def then_cloned_chart_has_distinct_chart_part(context):
    assert context.cloned_chart.part is not context._src_chart_part


@then("the cloned chart has a distinct embedded xlsx part")
def then_cloned_chart_has_distinct_xlsx_part(context):
    src_xlsx = context._src_xlsx_part
    tgt_xlsx = context.cloned_chart.part.chart_workbook.xlsx_part
    assert src_xlsx is not None
    assert tgt_xlsx is not None
    assert src_xlsx is not tgt_xlsx


@then("the cloned chart's workbook bytes equal the source workbook bytes")
def then_cloned_workbook_bytes_match(context):
    assert context.cloned_chart.workbook == context.source_chart.workbook


@then("the round-tripped presentation still has two charts")
def then_round_tripped_prs_has_two_charts(context):
    import io as _io

    bio = _io.BytesIO()
    context.prs.save(bio)
    bio.seek(0)
    prs2 = Presentation(bio)
    slides = list(prs2.slides)
    assert len(slides) == 2
    # -- last shape of each slide is our chart graphicFrame --
    gf1 = slides[0].shapes[-1]
    gf2 = slides[1].shapes[-1]
    assert gf1.has_chart and gf2.has_chart
    # -- both charts have embedded workbooks --
    assert gf1.chart.workbook is not None
    assert gf2.chart.workbook is not None
    # -- and the charts live in distinct parts on reload --
    assert gf1.chart.part is not gf2.chart.part


@given("a chart in presentation A")
def given_a_chart_in_presentation_A(context):
    prs_a = Presentation()
    slide_a = prs_a.slides.add_slide(prs_a.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["X", "Y", "Z"]
    cd.add_series("T", (5.0, 6.0, 7.0))
    gf = slide_a.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(1), Inches(1), Inches(5), Inches(3),
        cd,
    )
    context.prs_a = prs_a
    context.source_chart = gf.chart


@given("an empty presentation B with one slide")
def given_an_empty_presentation_B(context):
    prs_b = Presentation()
    slide_b = prs_b.slides.add_slide(prs_b.slide_layouts[5])
    context.prs_b = prs_b
    context.slide_B = slide_b
    context._x = Inches(1)
    context._y = Inches(1)
    context._cx = Inches(4)
    context._cy = Inches(3)


@when("I call chart.clone_to(slide_B.shapes, x, y, cx, cy)")
def when_I_call_chart_clone_to_slide_B(context):
    new_gf = context.source_chart.clone_to(
        context.slide_B.shapes, context._x, context._y, context._cx, context._cy
    )
    context.new_gf = new_gf
    context.cloned_chart = new_gf.chart


@then("presentation B's slide has a chart")
def then_prs_B_slide_has_a_chart(context):
    assert context.new_gf.has_chart


@then("that chart's chart part belongs to presentation B's package")
def then_chart_part_belongs_to_prs_B(context):
    assert context.cloned_chart.part.package is context.prs_b.part.package


@then("saving presentation B alone round-trips the chart")
def then_prs_B_round_trips_the_chart(context):
    import io as _io

    bio = _io.BytesIO()
    context.prs_b.save(bio)
    bio.seek(0)
    prs_b2 = Presentation(bio)
    s = prs_b2.slides[0]
    gf = s.shapes[-1]
    assert gf.has_chart
    assert gf.chart.workbook is not None
    assert gf.chart.chart_type == XL_CHART_TYPE.LINE


# --- add_chart_from (CLO-8) ----------------------------------------

@when("I call slide_2.shapes.add_chart_from(source_chart, x, y, cx, cy)")
def when_I_call_slide_2_add_chart_from(context):
    context._src_xlsx_part = context.source_chart.part.chart_workbook.xlsx_part
    context._src_chart_part = context.source_chart.part
    new_gf = context.slide_2.shapes.add_chart_from(
        context.source_chart, context._x, context._y, context._cx, context._cy
    )
    context.new_gf = new_gf
    context.cloned_chart = new_gf.chart


@when("I call slide_2.shapes.add_chart_from(source_chart, x, y) with no size")
def when_I_call_slide_2_add_chart_from_no_size(context):
    context._src_xlsx_part = context.source_chart.part.chart_workbook.xlsx_part
    context._src_chart_part = context.source_chart.part
    new_gf = context.slide_2.shapes.add_chart_from(
        context.source_chart, context._x, context._y
    )
    context.new_gf = new_gf
    context.cloned_chart = new_gf.chart


@then("slide 2 has a chart shape at (x, y) sized 5 by 3 inches")
def then_slide_2_has_chart_default_size(context):
    gf = context.new_gf
    assert gf.has_chart
    assert gf.left == context._x
    assert gf.top == context._y
    assert gf.width == Inches(5)
    assert gf.height == Inches(3)


# --- Chart.clone_from (CLO-8 follow-up, in-place chart clone) -------

@given("two charts with embedded workbooks on two different slides")
def given_two_charts_on_two_slides(context):
    prs = Presentation()
    slide1 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])

    src_cd = CategoryChartData()
    src_cd.categories = ["A", "B", "C"]
    src_cd.add_series("Src", (10.0, 20.0, 30.0))
    src_gf = slide1.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(1), Inches(1), Inches(5), Inches(3),
        src_cd,
    )

    tgt_cd = CategoryChartData()
    tgt_cd.categories = ["X", "Y", "Z"]
    tgt_cd.add_series("Tgt", (1.0, 2.0, 3.0))
    tgt_gf = slide2.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(2), Inches(2), Inches(4), Inches(3),
        tgt_cd,
    )

    context.prs = prs
    context.source_chart = src_gf.chart
    context.target_chart = tgt_gf.chart
    context.target_gf = tgt_gf
    context._tgt_partname_before = tgt_gf.chart.part.partname
    context._tgt_part_before = tgt_gf.chart.part
    context._src_xlsx_part = src_gf.chart.part.chart_workbook.xlsx_part
    context._src_chart_type = XL_CHART_TYPE.LINE


@when("I call target_chart.clone_from(source_chart)")
def when_I_call_target_chart_clone_from(context):
    result = context.target_chart.clone_from(context.source_chart)
    context._clone_from_result = result


@then("the target chart's partname is preserved")
def then_target_partname_preserved(context):
    assert context.target_chart.part is context._tgt_part_before
    assert context.target_chart.part.partname == context._tgt_partname_before


@then("the target chart's chart-type matches the source chart")
def then_target_chart_type_matches_source(context):
    assert context.target_chart.chart_type == context._src_chart_type


@then("the target chart has a distinct embedded xlsx part from the source")
def then_target_has_distinct_xlsx_part(context):
    src_xlsx = context._src_xlsx_part
    tgt_xlsx = context.target_chart.part.chart_workbook.xlsx_part
    assert src_xlsx is not None
    assert tgt_xlsx is not None
    assert src_xlsx is not tgt_xlsx


@then("the target chart's workbook bytes equal the source workbook bytes")
def then_target_workbook_equals_source(context):
    assert context.target_chart.workbook == context.source_chart.workbook


@then("the slide's graphicFrame still references the target chart part")
def then_gf_still_references_target_chart_part(context):
    assert context.target_gf.chart.part is context._tgt_part_before


@then("the round-tripped presentation still has both charts")
def then_round_trip_still_has_both_charts(context):
    import io as _io

    bio = _io.BytesIO()
    context.prs.save(bio)
    bio.seek(0)
    prs2 = Presentation(bio)
    slides = list(prs2.slides)
    assert len(slides) == 2
    gf1 = slides[0].shapes[-1]
    gf2 = slides[1].shapes[-1]
    assert gf1.has_chart and gf2.has_chart
    # -- both charts have their own embedded workbooks --
    assert gf1.chart.workbook is not None
    assert gf2.chart.workbook is not None
    # -- workbook bytes match (source workbook was cloned into target) --
    assert gf1.chart.workbook == gf2.chart.workbook
    # -- but distinct parts on reload --
    assert gf1.chart.part is not gf2.chart.part


# --- replace_data_preserve_formulas (issue #239) --------------------

@given("a chart with an embedded workbook and a formula in the last value cell")
def given_chart_with_formula_cell_in_workbook(context):
    """Build a BAR chart then inject a formula into B4 (3rd value cell)."""
    import io
    import zipfile

    from lxml import etree

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("S1", (1.0, 2.0, 3.0))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1), Inches(1), Inches(6), Inches(4),
        chart_data,
    ).chart

    # -- rewrite the embedded workbook to mark B4 (row 4, col 2) as a formula --
    orig = chart.workbook
    with zipfile.ZipFile(io.BytesIO(orig)) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ws_path = "xl/worksheets/sheet1.xml"
    ws = etree.fromstring(files[ws_path])
    sheetData = ws.find(f"{{{SML}}}sheetData")
    for row in sheetData.findall(f"{{{SML}}}row"):
        if row.get("r") != "4":
            continue
        for c in row.findall(f"{{{SML}}}c"):
            if c.get("r") == "B4":
                # -- replace <v>3</v> with <f>B2+B3</f><v>3</v> --
                for child in list(c):
                    c.remove(child)
                f_elm = etree.SubElement(c, f"{{{SML}}}f")
                f_elm.text = "B2+B3"
                v_elm = etree.SubElement(c, f"{{{SML}}}v")
                v_elm.text = "3"
    files[ws_path] = etree.tostring(
        ws, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in files.items():
            zout.writestr(n, b)
    chart.workbook = out.getvalue()
    context.chart = chart


@when("I call chart.replace_data_preserve_formulas(new_data)")
def when_I_call_replace_data_preserve_formulas(context):
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    cd.add_series("S1", (10.0, 20.0, 30.0))
    context._written = context.chart.replace_data_preserve_formulas(cd)


@then("the first-series first-value cell holds the new value")
def then_first_value_cell_holds_new_value(context):
    from pptx.chart.xlsx import WorkbookReader

    with WorkbookReader(context.chart.workbook) as reader:
        # -- B2 is the first value cell (row 2, col 2) --
        assert reader.cell_value("Sheet1", 2, 2) == 10.0


# ---------------------------------------------------------------------------
# Chart user-shapes (issue #351)
# ---------------------------------------------------------------------------


@given("a newly-authored bar chart")
def given_a_newly_authored_bar_chart(context):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("S", (1.0, 2.0, 3.0))
    context.chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    ).chart


@given("a chart with two user-shape anchors attached")
def given_a_chart_with_two_user_shape_anchors(context):
    # -- build a baseline bar chart and attach a ChartDrawingPart seeded
    # -- with one cdr:relSizeAnchor and one cdr:absSizeAnchor. This mirrors
    # -- the on-disk structure PowerPoint writes when a user annotates a
    # -- chart via Insert > Shapes; python-pptx itself has no authoring
    # -- API yet (see docs/dev/analysis/chart-user-shapes.rst), so the
    # -- anchors are synthesised directly on the raw lxml element.
    from lxml import etree

    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.oxml.ns import qn
    from pptx.parts.chartdrawing import ChartDrawingPart

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("S", (1.0, 2.0, 3.0))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    ).chart

    drawing_part = ChartDrawingPart.new(chart.part.package)
    rel_anchor = etree.SubElement(drawing_part._element, qn("cdr:relSizeAnchor"))
    etree.SubElement(rel_anchor, qn("cdr:from"))
    etree.SubElement(rel_anchor, qn("cdr:to"))
    etree.SubElement(rel_anchor, qn("cdr:sp"))
    abs_anchor = etree.SubElement(drawing_part._element, qn("cdr:absSizeAnchor"))
    etree.SubElement(abs_anchor, qn("cdr:from"))
    etree.SubElement(abs_anchor, qn("cdr:ext"))
    etree.SubElement(abs_anchor, qn("cdr:cxnSp"))

    chart.part.relate_to(drawing_part, RT.CHART_USER_SHAPES)
    context.chart = chart


@then("chart.user_shapes is None")
def then_chart_user_shapes_is_None(context):
    assert context.chart.user_shapes is None, (
        "expected chart.user_shapes to be None, got %r" % context.chart.user_shapes
    )


@then("chart.has_user_shapes is {value}")
def then_chart_has_user_shapes_is(context, value):
    expected = {"True": True, "False": False}[value]
    actual = context.chart.has_user_shapes
    assert actual is expected, (
        "expected chart.has_user_shapes to be %r, got %r" % (expected, actual)
    )


@then("chart.user_shapes is a ChartDrawingPart object")
def then_chart_user_shapes_is_a_ChartDrawingPart_object(context):
    from pptx.parts.chartdrawing import ChartDrawingPart

    assert isinstance(context.chart.user_shapes, ChartDrawingPart), (
        "expected chart.user_shapes to be a ChartDrawingPart, got %r"
        % type(context.chart.user_shapes).__name__
    )


@then("chart.user_shapes.anchor_count is {count_:d}")
def then_chart_user_shapes_anchor_count_is(context, count_):
    actual = context.chart.user_shapes.anchor_count
    assert actual == count_, (
        "expected chart.user_shapes.anchor_count to be %d, got %d" % (count_, actual)
    )


@then("iterating chart.user_shapes yields {count_:d} anchor elements")
def then_iterating_chart_user_shapes_yields_n_anchors(context, count_):
    anchors = list(context.chart.user_shapes.iter_anchor_elements())
    assert len(anchors) == count_, (
        "expected %d anchor elements, got %d" % (count_, len(anchors))
    )
    context.anchors = anchors


@then("the anchor tags are cdr:relSizeAnchor and cdr:absSizeAnchor")
def then_anchor_tags_match(context):
    from pptx.oxml.ns import qn

    expected = [qn("cdr:relSizeAnchor"), qn("cdr:absSizeAnchor")]
    actual = [a.tag for a in context.anchors]
    assert actual == expected, "expected %r, got %r" % (expected, actual)


# -- issue #764: chart-customization convenience helpers --


@when('I call chart.set_title("{text}")')
def when_I_call_chart_set_title_text(context, text):
    context.chart.set_title(text)


@when("I call chart.set_title(None)")
def when_I_call_chart_set_title_None(context):
    context.chart.set_title(None)


@when('I call chart.set_axis_title("{axis}", "{text}")')
def when_I_call_chart_set_axis_title(context, axis, text):
    context.chart.set_axis_title(axis, text)


@then('chart.chart_title.text_frame.text is "{text}"')
def then_chart_title_text_is(context, text):
    actual = context.chart.chart_title.text_frame.text
    assert actual == text, "expected %r, got %r" % (text, actual)


@then("chart.value_axis.has_title is True")
def then_chart_value_axis_has_title_is_True(context):
    assert context.chart.value_axis.has_title is True


@then('chart.value_axis.axis_title.text_frame.text is "{text}"')
def then_chart_value_axis_title_text_is(context, text):
    actual = context.chart.value_axis.axis_title.text_frame.text
    assert actual == text, "expected %r, got %r" % (text, actual)


@then("the formula cell still contains its original formula")
def then_formula_cell_unchanged(context):
    import io
    import zipfile

    from lxml import etree

    SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(io.BytesIO(context.chart.workbook)) as z:
        ws = etree.fromstring(z.read("xl/worksheets/sheet1.xml"))
    for c in ws.iter(f"{{{SML}}}c"):
        if c.get("r") == "B4":
            f_elm = c.find(f"{{{SML}}}f")
            assert f_elm is not None, "formula element missing"
            assert f_elm.text == "B2+B3", "formula text was modified: %r" % f_elm.text
            break
    else:
        raise AssertionError("B4 cell not found in worksheet")
