# encoding: utf-8

"""
Test suite for pptx.chart.series module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.series import BarSeries, LineSeries, PieSeries, SeriesFactory
from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


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


class DescribeSeriesFactory(object):

    def it_contructs_a_series_object_from_a_plot_element(self, call_fixture):
        xChart, ser, SeriesCls_, series_ = call_fixture
        series = SeriesFactory(xChart, ser)
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
        xChart = element(xChart_cxml)
        ser = xChart[0]
        return xChart, ser, SeriesCls_, series_

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
