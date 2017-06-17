# encoding: utf-8

"""
Test suite for pptx.chart.series module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.marker import Marker
from pptx.chart.point import BubblePoints, CategoryPoints, XyPoints
from pptx.chart.series import (
    AreaSeries, BarSeries, _BaseCategorySeries, _BaseSeries, BubbleSeries,
    LineSeries, _MarkerMixin, PieSeries, RadarSeries, SeriesCollection,
    _SeriesFactory, XySeries
)
from pptx.dml.chtfmt import ChartFormat

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, function_mock, instance_mock


class Describe_BaseSeries(object):

    def it_knows_its_name(self, name_fixture):
        series, expected_value = name_fixture
        assert series.name == expected_value

    def it_knows_its_position_in_the_series_sequence(self, index_fixture):
        series, expected_value = index_fixture
        assert series.index == expected_value

    def it_provides_access_to_its_format(self, format_fixture):
        series, ChartFormat_, ser, format_ = format_fixture
        format = series.format
        ChartFormat_.assert_called_once_with(ser)
        assert format is format_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def format_fixture(self, ChartFormat_, chart_format_):
        ser = element('c:ser')
        series = _BaseSeries(ser)
        return series, ChartFormat_, ser, chart_format_

    @pytest.fixture
    def index_fixture(self):
        ser_cxml, expected_value = 'c:ser/c:idx{val=42}', 42
        series = _BaseSeries(element(ser_cxml))
        return series, expected_value

    @pytest.fixture(params=[
        ('c:ser',                                           ''),
        ('c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v"foobar"', 'foobar'),
    ])
    def name_fixture(self, request):
        ser_cxml, expected_value = request.param
        series = _BaseSeries(element(ser_cxml))
        return series, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartFormat_(self, request, chart_format_):
        return class_mock(
            request, 'pptx.chart.series.ChartFormat',
            return_value=chart_format_
        )

    @pytest.fixture
    def chart_format_(self, request):
        return instance_mock(request, ChartFormat)


class Describe_BaseCategorySeries(object):

    def it_is_a_BaseSeries_subclass(self, subclass_fixture):
        base_category_series = subclass_fixture
        assert isinstance(base_category_series, _BaseSeries)

    def it_provides_access_to_its_points(self, points_fixture):
        series, CategoryPoints_, ser, points_ = points_fixture
        points = series.points
        CategoryPoints_.assert_called_once_with(ser)
        assert points is points_

    def it_knows_its_values(self, values_get_fixture):
        series, expected_value = values_get_fixture
        assert series.values == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def points_fixture(self, CategoryPoints_, points_):
        ser = element('c:ser')
        series = _BaseCategorySeries(ser)
        return series, CategoryPoints_, ser, points_

    @pytest.fixture
    def subclass_fixture(self):
        return _BaseCategorySeries(None)

    @pytest.fixture(params=[
        ('c:ser',                                            ()),
        ('c:ser/c:val/c:numRef',                             ()),
        ('c:ser/c:val/c:numLit',                             ()),
        ('c:ser/c:val/c:numRef/c:numCache',                  ()),
        ('c:ser/c:val/c:numRef/c:numCache/c:ptCount{val=0}', ()),
        ('c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"'
         '1.1")',
         (1.1,)),
        ('c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=0}/c:v"'
         '1.1",c:pt{idx=2}/c:v"3.3")',
         (1.1, None, 3.3)),
        ('c:ser/c:val/c:numLit/(c:ptCount{val=3},c:pt{idx=0}/c:v"1.1",c:pt{i'
         'dx=2}/c:v"3.3")',
         (1.1, None, 3.3)),
        ('c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=2}/c:v"'
         '3.3",c:pt{idx=0}/c:v"1.1")',
         (1.1, None, 3.3)),
    ])
    def values_get_fixture(self, request):
        ser_cxml, expected_value = request.param
        series = _BaseCategorySeries(element(ser_cxml))
        return series, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def CategoryPoints_(self, request, points_):
        return class_mock(
            request, 'pptx.chart.series.CategoryPoints', return_value=points_
        )

    @pytest.fixture
    def points_(self, request):
        return instance_mock(request, CategoryPoints)


class Describe_MarkerMixin(object):

    def it_provides_access_to_the_series_marker(self, marker_fixture):
        series, Marker_, ser, marker_ = marker_fixture
        marker = series.marker
        Marker_.assert_called_once_with(ser)
        assert marker is marker_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def marker_fixture(self, Marker_, marker_):
        ser = element('c:ser')
        series = LineSeries(ser)
        return series, Marker_, ser, marker_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Marker_(self, request, marker_):
        return class_mock(
            request, 'pptx.chart.series.Marker', return_value=marker_
        )

    @pytest.fixture
    def marker_(self, request):
        return instance_mock(request, Marker)


class DescribeAreaSeries(object):

    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        area_series = subclass_fixture
        assert isinstance(area_series, _BaseCategorySeries)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def subclass_fixture(self):
        return AreaSeries(None)


class DescribeBarSeries(object):

    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        bar_series = subclass_fixture
        assert isinstance(bar_series, _BaseCategorySeries)

    def it_knows_whether_it_should_invert_if_negative(
            self, invert_if_negative_get_fixture):
        bar_series, expected_value = invert_if_negative_get_fixture
        assert bar_series.invert_if_negative == expected_value

    def it_can_change_whether_it_inverts_if_negative(
            self, invert_if_negative_set_fixture):
        bar_series, new_value, expected_xml = invert_if_negative_set_fixture
        bar_series.invert_if_negative = new_value
        assert bar_series._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:ser',                           True),
        ('c:ser/c:invertIfNegative',        True),
        ('c:ser/c:invertIfNegative{val=1}', True),
        ('c:ser/c:invertIfNegative{val=0}', False),
    ])
    def invert_if_negative_get_fixture(self, request):
        ser_cxml, expected_value = request.param
        bar_series = BarSeries(element(ser_cxml))
        return bar_series, expected_value

    @pytest.fixture(params=[
        ('c:ser/c:order', True,
         'c:ser/(c:order,c:invertIfNegative{val=1})'),
        ('c:ser/c:order', False,
         'c:ser/(c:order,c:invertIfNegative{val=0})'),
        ('c:ser/(c:spPr,c:invertIfNegative{val=0})', True,
         'c:ser/(c:spPr,c:invertIfNegative{val=1})'),
        ('c:ser/(c:tx,c:invertIfNegative{val=1})', False,
         'c:ser/(c:tx,c:invertIfNegative{val=0})'),
    ])
    def invert_if_negative_set_fixture(self, request):
        ser_cxml, new_value, expected_ser_cxml = request.param
        bar_series = BarSeries(element(ser_cxml))
        expected_xml = xml(expected_ser_cxml)
        return bar_series, new_value, expected_xml

    @pytest.fixture
    def subclass_fixture(self):
        return BarSeries(None)


class Describe_BubbleSeries(object):

    def it_provides_access_to_its_points(self, points_fixture):
        series, BubblePoints_, ser, points_ = points_fixture
        points = series.points
        BubblePoints_.assert_called_once_with(ser)
        assert points is points_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def points_fixture(self, BubblePoints_, points_):
        ser = element('c:ser')
        series = BubbleSeries(ser)
        return series, BubblePoints_, ser, points_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BubblePoints_(self, request, points_):
        return class_mock(
            request, 'pptx.chart.series.BubblePoints', return_value=points_
        )

    @pytest.fixture
    def points_(self, request):
        return instance_mock(request, BubblePoints)


class DescribeLineSeries(object):

    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _BaseCategorySeries)

    def it_uses__MarkerMixin(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _MarkerMixin)

    def it_knows_whether_it_should_use_curve_smoothing(
            self, smooth_get_fixture):
        series, expected_value = smooth_get_fixture
        assert series.smooth == expected_value

    def it_can_change_whether_it_uses_curve_smoothing(
            self, smooth_set_fixture):
        series, new_value, expected_xml = smooth_set_fixture
        series.smooth = new_value
        assert series._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:ser',                 True),
        ('c:ser/c:smooth',        True),
        ('c:ser/c:smooth{val=1}', True),
        ('c:ser/c:smooth{val=0}', False),
    ])
    def smooth_get_fixture(self, request):
        ser_cxml, expected_value = request.param
        series = LineSeries(element(ser_cxml))
        return series, expected_value

    @pytest.fixture(params=[
        ('c:ser',                 True,  'c:ser/c:smooth'),
        ('c:ser/c:smooth',        False, 'c:ser/c:smooth{val=0}'),
        ('c:ser/c:smooth{val=0}', True,  'c:ser/c:smooth'),
    ])
    def smooth_set_fixture(self, request):
        ser_cxml, new_value, expected_ser_cxml = request.param
        series = LineSeries(element(ser_cxml))
        expected_xml = xml(expected_ser_cxml)
        return series, new_value, expected_xml

    @pytest.fixture
    def subclass_fixture(self):
        return LineSeries(None)


class DescribePieSeries(object):

    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        pie_series = subclass_fixture
        assert isinstance(pie_series, _BaseCategorySeries)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def subclass_fixture(self):
        return PieSeries(None)


class DescribeRadarSeries(object):

    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        radar_series = subclass_fixture
        assert isinstance(radar_series, _BaseCategorySeries)

    def it_uses__MarkerMixin(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _MarkerMixin)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def subclass_fixture(self):
        return RadarSeries(None)


class Describe_XySeries(object):

    def it_uses__MarkerMixin(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _MarkerMixin)

    def it_provides_access_to_its_points(self, points_fixture):
        series, XyPoints_, ser, points_ = points_fixture
        points = series.points
        XyPoints_.assert_called_once_with(ser)
        assert points is points_

    def it_knows_its_values(self, values_get_fixture):
        series, expected_values = values_get_fixture
        assert series.values == expected_values

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def points_fixture(self, XyPoints_, points_):
        ser = element('c:ser')
        series = XySeries(ser)
        return series, XyPoints_, ser, points_

    @pytest.fixture
    def subclass_fixture(self):
        return XySeries(None)

    @pytest.fixture(params=[
        ('c:ser', ()),
        ('c:ser/c:yVal/c:numRef', ()),
        ('c:ser/c:val/c:numRef/c:numCache', ()),
        ('c:ser/c:yVal/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v'
         '"1.1")', (1.1,)),
        ('c:ser/c:yVal/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=0}/c:v'
         '"1.1",c:pt{idx=2}/c:v"3.3")', (1.1, None, 3.3)),
        ('c:ser/c:val/c:numLit', ()),
        ('c:ser/c:yVal/c:numLit/(c:ptCount{val=3},c:pt{idx=0}/c:v"1.1",c:pt{'
         'idx=2}/c:v"3.3")', (1.1, None, 3.3)),
    ])
    def values_get_fixture(self, request):
        ser_cxml, expected_values = request.param
        series = XySeries(element(ser_cxml))
        return series, expected_values

    # fixture components ---------------------------------------------

    @pytest.fixture
    def XyPoints_(self, request, points_):
        return class_mock(
            request, 'pptx.chart.series.XyPoints', return_value=points_
        )

    @pytest.fixture
    def points_(self, request):
        return instance_mock(request, XyPoints)


class DescribeSeriesCollection(object):

    def it_supports_indexed_access(self, getitem_fixture):
        series_collection, index, _SeriesFactory_, ser, series_ = (
            getitem_fixture
        )
        series = series_collection[index]
        _SeriesFactory_.assert_called_once_with(ser)
        assert series is series_

    def it_supports_len(self, len_fixture):
        series_collection, expected_len = len_fixture
        assert len(series_collection) == expected_len

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:barChart/c:ser/c:order{val=42}', 0, 0),
        ('c:barChart/(c:ser/c:order{val=9},c:ser/c:order{val=6},c:ser/c:orde'
         'r{val=3})', 2, 0),
        ('c:plotArea/(c:barChart/(c:ser/(c:idx{val=1},c:order{val=3}),c:ser/'
         '(c:idx{val=0},c:order{val=0})),c:lineChart/c:ser/(c:idx{val=2},c:o'
         'rder{val=1}))', 1, 0),
    ])
    def getitem_fixture(self, request, _SeriesFactory_, series_):
        cxml, index, _offset = request.param
        parent_elm = element(cxml)
        ser = parent_elm.xpath('.//c:ser')[_offset]
        series_collection = SeriesCollection(parent_elm)
        return series_collection, index, _SeriesFactory_, ser, series_

    @pytest.fixture(params=[
        ('c:barChart', 0),
        ('c:barChart/c:ser/c:order{val=4}', 1),
        ('c:barChart/(c:ser/c:order{val=4},c:ser/c:order{val=1},c:ser/c:orde'
         'r{val=6})', 3),
        ('c:plotArea/c:barChart', 0),
        ('c:plotArea/c:barChart/c:ser/c:order{val=4}', 1),
        ('c:plotArea/c:barChart/(c:ser/c:order{val=4},c:ser/c:order{val=1},c'
         ':ser/c:order{val=6})', 3),
    ])
    def len_fixture(self, request):
        cxml, expected_len = request.param
        series_collection = SeriesCollection(element(cxml))
        return series_collection, expected_len

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _SeriesFactory_(self, request, series_):
        return function_mock(
            request, 'pptx.chart.series._SeriesFactory', return_value=series_
        )

    @pytest.fixture
    def series_(self, request):
        return instance_mock(request, _BaseSeries)


class Describe_SeriesFactory(object):

    def it_contructs_a_series_object_from_a_plot_element(self, call_fixture):
        ser, SeriesCls_, series_ = call_fixture
        series = _SeriesFactory(ser)
        SeriesCls_.assert_called_once_with(ser)
        assert series is series_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:areaChart/c:ser',     'AreaSeries'),
        ('c:barChart/c:ser',      'BarSeries'),
        ('c:bubbleChart/c:ser',   'BubbleSeries'),
        ('c:doughnutChart/c:ser', 'PieSeries'),
        ('c:lineChart/c:ser',     'LineSeries'),
        ('c:pieChart/c:ser',      'PieSeries'),
        ('c:radarChart/c:ser',    'RadarSeries'),
        ('c:scatterChart/c:ser',  'XySeries'),
    ])
    def call_fixture(self, request):
        xChart_cxml, cls_name = request.param
        ser = element(xChart_cxml).ser_lst[0]
        SeriesCls_ = class_mock(request, 'pptx.chart.series.%s' % cls_name)
        series_ = SeriesCls_.return_value
        return ser, SeriesCls_, series_
