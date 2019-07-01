# encoding: utf-8

"""
Test suite for pptx.chart.plot module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.category import Categories
from pptx.chart.chart import Chart
from pptx.chart.plot import (
    _BasePlot,
    AreaPlot,
    Area3DPlot,
    BarPlot,
    BubblePlot,
    DataLabels,
    DoughnutPlot,
    LinePlot,
    PiePlot,
    PlotFactory,
    PlotTypeInspector,
    RadarPlot,
    XyPlot,
)
from pptx.chart.series import SeriesCollection
from pptx.enum.chart import XL_CHART_TYPE as XL

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class Describe_BasePlot(object):
    def it_knows_which_chart_it_belongs_to(self, chart_fixture):
        plot, expected_value = chart_fixture
        assert plot.chart == expected_value

    def it_knows_whether_it_has_data_labels(self, has_data_labels_get_fixture):
        plot, expected_value = has_data_labels_get_fixture
        assert plot.has_data_labels == expected_value

    def it_can_change_whether_it_has_data_labels(self, has_data_labels_set_fixture):
        plot, new_value, expected_xml = has_data_labels_set_fixture
        plot.has_data_labels = new_value
        assert plot._element.xml == expected_xml

    def it_knows_whether_it_varies_color_by_category(
        self, vary_by_categories_get_fixture
    ):
        plot, expected_value = vary_by_categories_get_fixture
        assert plot.vary_by_categories == expected_value

    def it_can_change_whether_it_varies_color_by_category(
        self, vary_by_categories_set_fixture
    ):
        plot, new_value, expected_xml = vary_by_categories_set_fixture
        plot.vary_by_categories = new_value
        assert plot._element.xml == expected_xml

    def it_provides_access_to_its_categories(self, categories_fixture):
        plot, categories_, Categories_, xChart = categories_fixture
        categories = plot.categories
        Categories_.assert_called_once_with(xChart)
        assert categories is categories_

    def it_provides_access_to_the_data_labels(self, data_labels_fixture):
        plot, data_labels_, DataLabels_, dLbls = data_labels_fixture
        data_labels = plot.data_labels
        DataLabels_.assert_called_once_with(dLbls)
        assert data_labels is data_labels_

    def it_provides_access_to_its_series(self, series_fixture):
        plot, series_, SeriesCollection_, xChart = series_fixture
        series = plot.series
        SeriesCollection_.assert_called_once_with(xChart)
        assert series is series_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def categories_fixture(self, categories_, Categories_):
        xChart = element("c:barChart")
        plot = _BasePlot(xChart, None)
        return plot, categories_, Categories_, xChart

    @pytest.fixture
    def chart_fixture(self, chart_):
        plot = _BasePlot(None, chart_)
        expected_value = chart_
        return plot, expected_value

    @pytest.fixture
    def data_labels_fixture(self, DataLabels_, data_labels_):
        barChart = element("c:barChart/c:dLbls")
        dLbls = barChart[0]
        plot = _BasePlot(barChart, None)
        return plot, data_labels_, DataLabels_, dLbls

    @pytest.fixture(
        params=[
            ("c:barChart", False),
            ("c:barChart/c:dLbls", True),
            ("c:lineChart", False),
            ("c:lineChart/c:dLbls", True),
            ("c:pieChart", False),
            ("c:pieChart/c:dLbls", True),
        ]
    )
    def has_data_labels_get_fixture(self, request):
        xChart_cxml, expected_value = request.param
        plot = _BasePlot(element(xChart_cxml), None)
        return plot, expected_value

    @pytest.fixture(
        params=[
            ("c:barChart", True, "c:barChart/c:dLbls/+"),
            ("c:barChart/c:dLbls", True, "c:barChart/c:dLbls"),
            ("c:barChart", False, "c:barChart"),
            ("c:barChart/c:dLbls", False, "c:barChart"),
            ("c:bubbleChart", True, "c:bubbleChart/c:dLbls/+"),
            ("c:bubbleChart/c:dLbls", False, "c:bubbleChart"),
            ("c:lineChart", True, "c:lineChart/c:dLbls/+"),
            ("c:lineChart/c:dLbls", False, "c:lineChart"),
            ("c:pieChart", True, "c:pieChart/c:dLbls/+"),
            ("c:pieChart/c:dLbls", False, "c:pieChart"),
        ]
    )
    def has_data_labels_set_fixture(self, request):
        xChart_cxml, new_value, expected_cxml = request.param
        # apply extended suffix to replace trailing '+' where present
        if expected_cxml.endswith("+"):
            expected_cxml = expected_cxml[:-1] + (
                "(c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val="
                "0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSi"
                "ze{val=0},c:showLeaderLines{val=1})"
            )
        plot = PlotFactory(element(xChart_cxml), None)
        expected_xml = xml(expected_cxml)
        return plot, new_value, expected_xml

    @pytest.fixture
    def series_fixture(self, SeriesCollection_, series_collection_):
        xChart = element("c:barChart")
        plot = _BasePlot(xChart, None)
        return plot, series_collection_, SeriesCollection_, xChart

    @pytest.fixture(
        params=[
            ("c:barChart", True),
            ("c:lineChart/c:varyColors", True),
            ("c:pieChart/c:varyColors{val=0}", False),
        ]
    )
    def vary_by_categories_get_fixture(self, request):
        xChart_cxml, expected_value = request.param
        plot = _BasePlot(element(xChart_cxml), None)
        return plot, expected_value

    @pytest.fixture(
        params=[
            ("c:barChart", False, "c:barChart/c:varyColors{val=0}"),
            ("c:barChart", True, "c:barChart/c:varyColors"),
            ("c:lineChart/c:varyColors{val=0}", True, "c:lineChart/c:varyColors"),
            ("c:pieChart/c:varyColors{val=1}", False, "c:pieChart/c:varyColors{val=0}"),
        ]
    )
    def vary_by_categories_set_fixture(self, request):
        xChart_cxml, new_value, expected_xChart_cxml = request.param
        plot = _BasePlot(element(xChart_cxml), None)
        expected_xml = xml(expected_xChart_cxml)
        return plot, new_value, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Categories_(self, request, categories_):
        return class_mock(
            request, "pptx.chart.plot.Categories", return_value=categories_
        )

    @pytest.fixture
    def categories_(self, request):
        return instance_mock(request, Categories)

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def DataLabels_(self, request, data_labels_):
        return class_mock(
            request, "pptx.chart.plot.DataLabels", return_value=data_labels_
        )

    @pytest.fixture
    def data_labels_(self, request):
        return instance_mock(request, DataLabels)

    @pytest.fixture
    def SeriesCollection_(self, request, series_collection_):
        return class_mock(
            request, "pptx.chart.plot.SeriesCollection", return_value=series_collection_
        )

    @pytest.fixture
    def series_collection_(self, request):
        return instance_mock(request, SeriesCollection)


class DescribeBarPlot(object):
    def it_knows_its_gap_width(self, gap_width_get_fixture):
        bar_plot, expected_value = gap_width_get_fixture
        assert bar_plot.gap_width == expected_value

    def it_can_change_its_gap_width(self, gap_width_set_fixture):
        bar_plot, new_value, expected_xml = gap_width_set_fixture
        bar_plot.gap_width = new_value
        assert bar_plot._element.xml == expected_xml

    def it_knows_how_much_it_overlaps_the_adjacent_bar(self, overlap_get_fixture):
        bar_plot, expected_value = overlap_get_fixture
        assert bar_plot.overlap == expected_value

    def it_can_change_how_much_it_overlaps_the_adjacent_bar(self, overlap_set_fixture):
        bar_plot, new_value, expected_xml = overlap_set_fixture
        bar_plot.overlap = new_value
        assert bar_plot._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:barChart", 150),
            ("c:barChart/c:gapWidth", 150),
            ("c:barChart/c:gapWidth{val=175}", 175),
            ("c:barChart/c:gapWidth{val=042%}", 42),
        ]
    )
    def gap_width_get_fixture(self, request):
        barChart_cxml, expected_value = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        return bar_plot, expected_value

    @pytest.fixture(
        params=[
            ("c:barChart", 42, "c:barChart/c:gapWidth{val=42}"),
            ("c:barChart", 150, "c:barChart/c:gapWidth"),
            ("c:barChart/c:gapWidth{val=200%}", 300, "c:barChart/c:gapWidth{val=300}"),
        ]
    )
    def gap_width_set_fixture(self, request):
        barChart_cxml, new_value, expected_barChart_cxml = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        expected_xml = xml(expected_barChart_cxml)
        return bar_plot, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:barChart", 0),
            ("c:barChart/c:overlap", 0),
            ("c:barChart/c:overlap{val=42}", 42),
            ("c:barChart/c:overlap{val=-42}", -42),
            ("c:barChart/c:overlap{val=-042%}", -42),
        ]
    )
    def overlap_get_fixture(self, request):
        barChart_cxml, expected_value = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        return bar_plot, expected_value

    @pytest.fixture(
        params=[
            ("c:barChart", 42, "c:barChart/c:overlap{val=42}"),
            ("c:barChart/c:overlap{val=42}", 24, "c:barChart/c:overlap{val=24}"),
            ("c:barChart/c:overlap{val=42}", -5, "c:barChart/c:overlap{val=-5}"),
            ("c:barChart/c:overlap{val=42}", 0, "c:barChart"),
        ]
    )
    def overlap_set_fixture(self, request):
        barChart_cxml, new_value, expected_barChart_cxml = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        expected_xml = xml(expected_barChart_cxml)
        return bar_plot, new_value, expected_xml


class DescribeBubblePlot(object):
    def it_knows_its_bubble_scale(self, bubble_scale_get_fixture):
        bubble_plot, expected_value = bubble_scale_get_fixture
        assert bubble_plot.bubble_scale == expected_value

    def it_can_change_its_bubble_scale(self, bubble_scale_set_fixture):
        bubble_plot, new_value, expected_xml = bubble_scale_set_fixture
        bubble_plot.bubble_scale = new_value
        assert bubble_plot._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:bubbleChart", 100),
            ("c:bubbleChart/c:bubbleScale", 100),
            ("c:bubbleChart/c:bubbleScale{val=175}", 175),
            ("c:bubbleChart/c:bubbleScale{val=070%}", 70),
        ]
    )
    def bubble_scale_get_fixture(self, request):
        bubbleChart_cxml, expected_value = request.param
        bubble_plot = BubblePlot(element(bubbleChart_cxml), None)
        return bubble_plot, expected_value

    @pytest.fixture(
        params=[
            ("c:bubbleChart", 42, "c:bubbleChart/c:bubbleScale{val=42}"),
            (
                "c:bubbleChart/c:bubbleScale{val=042%}",
                150,
                "c:bubbleChart/c:bubbleScale{val=150}",
            ),
            ("c:bubbleChart/c:bubbleScale{val=150}", None, "c:bubbleChart"),
            (
                "c:bubbleChart/c:bubbleScale{val=150}",
                100,
                "c:bubbleChart/c:bubbleScale",
            ),
            ("c:bubbleChart", None, "c:bubbleChart"),
            ("c:bubbleChart", 100, "c:bubbleChart/c:bubbleScale"),
        ]
    )
    def bubble_scale_set_fixture(self, request):
        bubbleChart_cxml, new_value, expected_cxml = request.param
        bubble_plot = BubblePlot(element(bubbleChart_cxml), None)
        expected_xml = xml(expected_cxml)
        return bubble_plot, new_value, expected_xml


class DescribePlotFactory(object):
    def it_contructs_a_plot_object_from_a_plot_element(self, call_fixture):
        xChart, chart_, PlotClass_, plot_ = call_fixture
        plot = PlotFactory(xChart, chart_)
        PlotClass_.assert_called_once_with(xChart, chart_)
        assert plot is plot_

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:areaChart", AreaPlot),
            ("c:area3DChart", Area3DPlot),
            ("c:barChart", BarPlot),
            ("c:bubbleChart", BubblePlot),
            ("c:doughnutChart", DoughnutPlot),
            ("c:lineChart", LinePlot),
            ("c:pieChart", PiePlot),
            ("c:radarChart", RadarPlot),
            ("c:scatterChart", XyPlot),
        ]
    )
    def call_fixture(self, request, chart_):
        xChart_cxml, PlotCls = request.param
        plot_ = instance_mock(request, PlotCls, name="plot_")
        class_spec = "pptx.chart.plot.%s" % PlotCls.__name__
        PlotClass_ = class_mock(request, class_spec, return_value=plot_)
        xChart = element(xChart_cxml)
        return xChart, chart_, PlotClass_, plot_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)


class DescribePlotTypeInspector(object):
    def it_can_determine_the_chart_type_of_a_plot(self, chart_type_fixture):
        plot, expected_chart_type = chart_type_fixture
        chart_type = PlotTypeInspector.chart_type(plot)
        assert chart_type is expected_chart_type

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:areaChart", XL.AREA),
            ("c:areaChart/c:grouping", XL.AREA),
            ("c:areaChart/c:grouping{val=standard}", XL.AREA),
            ("c:areaChart/c:grouping{val=stacked}", XL.AREA_STACKED),
            ("c:areaChart/c:grouping{val=percentStacked}", XL.AREA_STACKED_100),
            ("c:area3DChart", XL.THREE_D_AREA),
            ("c:area3DChart/c:grouping", XL.THREE_D_AREA),
            ("c:area3DChart/c:grouping{val=standard}", XL.THREE_D_AREA),
            ("c:area3DChart/c:grouping{val=stacked}", XL.THREE_D_AREA_STACKED),
            (
                "c:area3DChart/c:grouping{val=percentStacked}",
                XL.THREE_D_AREA_STACKED_100,
            ),
            ("c:barChart/c:barDir", XL.COLUMN_CLUSTERED),
            ("c:barChart/c:barDir{val=col}", XL.COLUMN_CLUSTERED),
            ("c:barChart/(c:barDir{val=col},c:grouping)", XL.COLUMN_CLUSTERED),
            (
                "c:barChart/(c:barDir{val=col},c:grouping{val=clustered})",
                XL.COLUMN_CLUSTERED,
            ),
            (
                "c:barChart/(c:barDir{val=col},c:grouping{val=stacked})",
                XL.COLUMN_STACKED,
            ),
            (
                "c:barChart/(c:barDir{val=col},c:grouping{val=percentStacked})",
                XL.COLUMN_STACKED_100,
            ),
            ("c:barChart/c:barDir{val=bar}", XL.BAR_CLUSTERED),
            ("c:barChart/(c:barDir{val=bar},c:grouping)", XL.BAR_CLUSTERED),
            (
                "c:barChart/(c:barDir{val=bar},c:grouping{val=clustered})",
                XL.BAR_CLUSTERED,
            ),
            ("c:barChart/(c:barDir{val=bar},c:grouping{val=stacked})", XL.BAR_STACKED),
            (
                "c:barChart/(c:barDir{val=bar},c:grouping{val=percentStacked})",
                XL.BAR_STACKED_100,
            ),
            ("c:doughnutChart", XL.DOUGHNUT),
            ("c:doughnutChart/c:ser/c:explosion{val=25}", XL.DOUGHNUT_EXPLODED),
            ("c:lineChart", XL.LINE_MARKERS),
            ("c:lineChart/c:grouping", XL.LINE_MARKERS),
            ("c:lineChart/c:grouping{val=standard}", XL.LINE_MARKERS),
            ("c:lineChart/c:grouping{val=stacked}", XL.LINE_MARKERS_STACKED),
            ("c:lineChart/c:grouping{val=percentStacked}", XL.LINE_MARKERS_STACKED_100),
            ("c:lineChart/c:ser/c:marker/c:symbol{val=none}", XL.LINE),
            (
                "c:lineChart/(c:grouping{val=stacked},c:ser/c:marker/c:symbol{val=n"
                "one})",
                XL.LINE_STACKED,
            ),
            (
                "c:lineChart/(c:grouping{val=percentStacked},c:ser/c:marker/c:symbo"
                "l{val=none})",
                XL.LINE_STACKED_100,
            ),
            ("c:pieChart", XL.PIE),
            ("c:pieChart/c:ser/c:explosion{val=25}", XL.PIE_EXPLODED),
            ("c:scatterChart/c:scatterStyle", XL.XY_SCATTER),
            (
                "c:scatterChart/(c:scatterStyle{val=lineMarker},c:ser/c:spPr/a:ln/a"
                ":noFill)",
                XL.XY_SCATTER,
            ),
            ("c:scatterChart/c:scatterStyle{val=lineMarker}", XL.XY_SCATTER_LINES),
            (
                "c:scatterChart/(c:scatterStyle{val=lineMarker},c:ser/c:marker/c:sy"
                "mbol{val=none})",
                XL.XY_SCATTER_LINES_NO_MARKERS,
            ),
            (
                "c:scatterChart/(c:scatterStyle{val=lineMarker},c:ser/c:marker/c:sy"
                "mbol{val=diamond})",
                XL.XY_SCATTER_LINES,
            ),
            ("c:scatterChart/c:scatterStyle{val=smoothMarker}", XL.XY_SCATTER_SMOOTH),
            (
                "c:scatterChart/(c:scatterStyle{val=smoothMarker},c:ser/c:marker/c:"
                "symbol{val=none})",
                XL.XY_SCATTER_SMOOTH_NO_MARKERS,
            ),
            ("c:bubbleChart", XL.BUBBLE),
            ("c:bubbleChart/c:ser", XL.BUBBLE),
            ("c:bubbleChart/c:ser/c:bubble3D", XL.BUBBLE_THREE_D_EFFECT),
            ("c:bubbleChart/c:ser/c:bubble3D{val=0}", XL.BUBBLE),
            ("c:bubbleChart/c:ser/c:bubble3D{val=1}", XL.BUBBLE_THREE_D_EFFECT),
            ("c:radarChart/c:radarStyle", XL.RADAR),
            ("c:radarChart/c:radarStyle{val=marker}", XL.RADAR_MARKERS),
            ("c:radarChart/c:radarStyle{val=filled}", XL.RADAR_FILLED),
            (
                "c:radarChart/(c:radarStyle{val=marker},c:ser/c:marker/c:symbol{val"
                "=none})",
                XL.RADAR,
            ),
        ]
    )
    def chart_type_fixture(self, request):
        xChart_cxml, expected_chart_type = request.param
        plot = PlotFactory(element(xChart_cxml), None)
        return plot, expected_chart_type
