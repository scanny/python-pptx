# encoding: utf-8

"""
Test suite for pptx.chart.series module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.series import (
    BarSeries, _BaseSeries, LineSeries, PieSeries, SeriesCollection,
    _SeriesFactory
)
from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, function_mock, instance_mock


class Describe_BaseSeries(object):

    def it_knows_its_name(self, name_fixture):
        series, expected_value = name_fixture
        assert series.name == expected_value

    def it_knows_its_position_in_the_series_sequence(self, index_fixture):
        series, expected_value = index_fixture
        assert series.index == expected_value

    def it_knows_its_values(self, values_get_fixture):
        series, expected_value = values_get_fixture
        assert series.values == expected_value

    # fixtures -------------------------------------------------------

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

    @pytest.fixture(params=[
        ('c:ser', ()),
        ('c:ser/c:val/c:numRef', ()),
        ('c:ser/c:val/c:numRef/c:numCache', ()),
        ('c:ser/c:val/c:numRef/c:numCache/(c:pt{idx=1}/c:v"2.3",c:pt{idx=0}/'
         'c:v"1.2",c:pt{idx=2}/c:v"3.4")',
         (1.2, 2.3, 3.4)),
        ('c:ser/c:val/c:numLit', ()),
        ('c:ser/c:val/c:numLit/(c:pt{idx=2}/c:v"6.7",c:pt{idx=0}/'
         'c:v"4.5",c:pt{idx=1}/c:v"5.6")',
         (4.5, 5.6, 6.7)),
    ])
    def values_get_fixture(self, request):
        ser_cxml, expected_value = request.param
        series = _BaseSeries(element(ser_cxml))
        return series, expected_value


class DescribeBarSeries(object):

    def it_provides_access_to_the_series_fill_format(self, fill_fixture):
        bar_series, FillFormat_, spPr, fill_ = fill_fixture
        fill = bar_series.fill
        FillFormat_.from_fill_parent.assert_called_once_with(spPr)
        assert fill is fill_

    def it_provides_access_to_the_series_line_format(self, line_fixture):
        bar_series, LineFormat_, line_ = line_fixture
        line = bar_series.line
        LineFormat_.assert_called_once_with(bar_series)
        assert line is line_

    def it_can_get_the_width_of_its_line(self, line_width_get_fixture):
        """
        Tests integration between BarSeries and LineFormat, the latter
        requiring an ln() method on the former to do this job.
        """
        bar_series, expected_width = line_width_get_fixture
        assert bar_series.line.width == expected_width

    def it_can_change_the_width_of_its_line(self, line_width_set_fixture):
        """
        Tests integration between BarSeries and LineFormat, the latter
        requiring a get_or_add_ln() method on the former to do this job.
        """
        bar_series, width, expected_xml = line_width_set_fixture
        bar_series.line.width = width
        assert bar_series._element.xml == expected_xml

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

    @pytest.fixture
    def fill_fixture(self, FillFormat_, fill_):
        ser = element('c:ser/c:spPr')
        spPr = ser.spPr
        bar_series = BarSeries(ser)
        return bar_series, FillFormat_, spPr, fill_

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
         'c:ser/(c:order,c:invertIfNegative)'),
        ('c:ser/c:order', False,
         'c:ser/(c:order,c:invertIfNegative{val=0})'),
        ('c:ser/(c:spPr,c:invertIfNegative{val=0})', True,
         'c:ser/(c:spPr,c:invertIfNegative)'),
        ('c:ser/(c:tx,c:invertIfNegative{val=1})', False,
         'c:ser/(c:tx,c:invertIfNegative{val=0})'),
    ])
    def invert_if_negative_set_fixture(self, request):
        ser_cxml, new_value, expected_ser_cxml = request.param
        bar_series = BarSeries(element(ser_cxml))
        expected_xml = xml(expected_ser_cxml)
        return bar_series, new_value, expected_xml

    @pytest.fixture
    def line_fixture(self, LineFormat_, line_):
        bar_series = BarSeries(element('c:ser'))
        return bar_series, LineFormat_, line_

    @pytest.fixture(params=[
        ('c:ser', 0),
        ('c:ser/c:spPr', 0),
        ('c:ser/c:spPr/a:ln{w=12700}', 12700),
    ])
    def line_width_get_fixture(self, request):
        ser_cxml, expected_width = request.param
        bar_series = BarSeries(element(ser_cxml))
        return bar_series, expected_width

    @pytest.fixture
    def line_width_set_fixture(self):
        ser_cxml = 'c:ser{a:foo=bar}/c:order'
        bar_series = BarSeries(element(ser_cxml))
        width = 12700  # 1 point
        expected_xml = xml('c:ser{a:foo=bar}/(c:order,c:spPr/a:ln{w=12700})')
        return bar_series, width, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def FillFormat_(self, request, fill_):
        FillFormat_ = class_mock(request, 'pptx.chart.series.FillFormat')
        FillFormat_.from_fill_parent.return_value = fill_
        return FillFormat_

    @pytest.fixture
    def fill_(self, request):
        return instance_mock(request, FillFormat)

    @pytest.fixture
    def LineFormat_(self, request, line_):
        return class_mock(
            request, 'pptx.chart.series.LineFormat', return_value=line_
        )

    @pytest.fixture
    def line_(self, request):
        return instance_mock(request, LineFormat)


class DescribeLineSeries(object):

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


class DescribeSeriesCollection(object):

    def it_supports_indexed_access(self, getitem_fixture):
        series_collection, idx, _SeriesFactory_, ser, series_ = (
            getitem_fixture
        )
        series = series_collection[idx]
        _SeriesFactory_.assert_called_once_with(ser)
        assert series is series_

    def it_supports_len(self, len_fixture):
        series_collection, expected_len = len_fixture
        assert len(series_collection) == expected_len

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:barChart/c:ser',               0),
        ('c:barChart/(c:ser,c:ser,c:ser)', 2),
        ('c:chartSpace/(c:barChart/(c:ser/c:idx{val=0},c:ser/c:idx{val=1}),c'
         ':lineChart/c:ser/c:idx{val=2})', 2),
    ])
    def getitem_fixture(self, request, _SeriesFactory_, series_):
        cxml, idx = request.param
        parent_elm = element(cxml)
        ser = parent_elm.xpath('.//c:ser')[idx]
        series_collection = SeriesCollection(parent_elm)
        return series_collection, idx, _SeriesFactory_, ser, series_

    @pytest.fixture(params=[
        ('c:barChart',                     0),
        ('c:barChart/c:ser',               1),
        ('c:barChart/(c:ser,c:ser)',       2),
        ('c:barChart/(c:idx,c:tx,c:ser)',  1),
        ('c:chartSpace/(c:barChart/(c:ser/c:idx{val=0},c:ser/c:idx{val=1}),c'
         ':lineChart/c:ser/c:idx{val=2})', 3),
    ])
    def len_fixture(self, request):
        xChart_cxml, expected_len = request.param
        series_collection = SeriesCollection(element(xChart_cxml))
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
        'barChart',
        'lineChart',
        'pieChart'
    ])
    def call_fixture(
            self, request, BarSeries_, bar_series_, LineSeries_,
            line_series_, PieSeries_, pie_series_):
        xChart_cxml, SeriesCls_, series_ = {
            'barChart':  ('c:barChart/c:ser',  BarSeries_,  bar_series_),
            'lineChart': ('c:lineChart/c:ser', LineSeries_, line_series_),
            'pieChart':  ('c:pieChart/c:ser',  PieSeries_,  pie_series_),
        }[request.param]
        ser = element(xChart_cxml).ser_lst[0]
        return ser, SeriesCls_, series_

    # fixture components -----------------------------------

    @pytest.fixture
    def BarSeries_(self, request, bar_series_):
        return class_mock(
            request, 'pptx.chart.series.BarSeries',
            return_value=bar_series_
        )

    @pytest.fixture
    def bar_series_(self, request):
        return instance_mock(request, BarSeries)

    @pytest.fixture
    def LineSeries_(self, request, line_series_):
        return class_mock(
            request, 'pptx.chart.series.LineSeries',
            return_value=line_series_
        )

    @pytest.fixture
    def line_series_(self, request):
        return instance_mock(request, LineSeries)

    @pytest.fixture
    def PieSeries_(self, request, pie_series_):
        return class_mock(
            request, 'pptx.chart.series.PieSeries',
            return_value=pie_series_
        )

    @pytest.fixture
    def pie_series_(self, request):
        return instance_mock(request, PieSeries)
