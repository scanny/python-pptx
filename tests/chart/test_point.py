# encoding: utf-8

"""
Unit test suite for the pptx.chart.point module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.chart.datalabel import DataLabel
from pptx.chart.marker import Marker
from pptx.chart.point import BubblePoints, CategoryPoints, Point, XyPoints
from pptx.dml.chtfmt import ChartFormat

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class Describe_BasePoints(object):

    def it_supports_indexed_access(self, getitem_fixture):
        points, idx, Point_, ser, point_ = getitem_fixture
        point = points[idx]
        Point_.assert_called_once_with(ser, idx)
        assert point is point_

    def it_raises_on_indexed_access_out_of_range(self):
        points = XyPoints(element(
            'c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:num'
            'Ref/c:numCache/c:ptCount{val=3})'
        ))
        with pytest.raises(IndexError):
            points[-1]
        with pytest.raises(IndexError):
            points[3]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getitem_fixture(self, request, Point_, point_):
        ser = element(
            'c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:num'
            'Ref/c:numCache/c:ptCount{val=3})'
        )
        points = XyPoints(ser)
        idx = 2
        return points, idx, Point_, ser, point_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Point_(self, request, point_):
        return class_mock(
            request, 'pptx.chart.point.Point', return_value=point_
        )

    @pytest.fixture
    def point_(self, request):
        return instance_mock(request, Point)


class DescribeBubblePoints(object):

    def it_supports_len(self, len_fixture):
        points, expected_value = len_fixture
        assert len(points) == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:ser', 0),
        ('c:ser/c:bubbleSize/c:numRef/c:numCache/c:ptCount{val=3}', 0),
        ('c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=1},c:yVal/c:numRef'
         '/c:numCache/c:ptCount{val=2},c:bubbleSize/c:numRef/c:numCache/c:pt'
         'Count{val=3})', 1),
        ('c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef'
         '/c:numCache/c:ptCount{val=3},c:bubbleSize/c:numRef/c:numCache/c:pt'
         'Count{val=3})', 3),
    ])
    def len_fixture(self, request):
        ser_cxml, expected_value = request.param
        points = BubblePoints(element(ser_cxml))
        return points, expected_value


class DescribeCategoryPoints(object):

    def it_supports_len(self, len_fixture):
        points, expected_value = len_fixture
        assert len(points) == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:ser', 0),
        ('c:ser/c:cat/c:numRef/c:numCache/c:ptCount{val=42}', 42),
        ('c:ser/c:cat/c:numRef/c:numCache/c:ptCount{val=24}', 24),
    ])
    def len_fixture(self, request):
        ser_cxml, expected_value = request.param
        points = CategoryPoints(element(ser_cxml))
        return points, expected_value


class DescribePoint(object):

    def it_provides_access_to_its_data_label(self, data_label_fixture):
        point, DataLabel_, ser, idx, data_label_ = data_label_fixture
        data_label = point.data_label
        DataLabel_.assert_called_once_with(ser, idx)
        assert data_label is data_label_

    def it_provides_access_to_its_format(self, format_fixture):
        point, ChartFormat_, ser, chart_format_, expected_xml = (
            format_fixture
        )
        chart_format = point.format
        ChartFormat_.assert_called_once_with(
            ser.xpath('c:dPt[c:idx/@val="42"]')[0]
        )
        assert chart_format is chart_format_
        assert point._element.xml == expected_xml

    def it_provides_access_to_its_marker(self, marker_fixture):
        point, Marker_, dPt, marker_ = marker_fixture
        marker = point.marker
        Marker_.assert_called_once_with(dPt)
        assert marker is marker_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def data_label_fixture(self, DataLabel_, data_label_):
        ser, idx = element('c:ser'), 42
        point = Point(ser, idx)
        return point, DataLabel_, ser, idx, data_label_

    @pytest.fixture(params=[
        ('c:ser',
         'c:ser/c:dPt/c:idx{val=42}'),
        ('c:ser/c:dPt/c:idx{val=42}',
         'c:ser/c:dPt/c:idx{val=42}'),
        ('c:ser/c:dPt/c:idx{val=45}',
         'c:ser/(c:dPt/c:idx{val=45},c:dPt/c:idx{val=42})'),
    ])
    def format_fixture(self, request, ChartFormat_, chart_format_):
        ser_cxml, expected_cxml = request.param
        ser = element(ser_cxml)
        point = Point(ser, 42)
        expected_xml = xml(expected_cxml)
        return point, ChartFormat_, ser, chart_format_, expected_xml

    @pytest.fixture
    def marker_fixture(self, Marker_, marker_):
        ser = element('c:ser/c:dPt/c:idx{val=42}')
        point = Point(ser, 42)
        dPt = ser[0]
        return point, Marker_, dPt, marker_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartFormat_(self, request, chart_format_):
        return class_mock(
            request, 'pptx.chart.point.ChartFormat',
            return_value=chart_format_
        )

    @pytest.fixture
    def chart_format_(self, request):
        return instance_mock(request, ChartFormat)

    @pytest.fixture
    def DataLabel_(self, request, data_label_):
        return class_mock(
            request, 'pptx.chart.point.DataLabel', return_value=data_label_
        )

    @pytest.fixture
    def data_label_(self, request):
        return instance_mock(request, DataLabel)

    @pytest.fixture
    def Marker_(self, request, marker_):
        return class_mock(
            request, 'pptx.chart.point.Marker', return_value=marker_
        )

    @pytest.fixture
    def marker_(self, request):
        return instance_mock(request, Marker)


class DescribeXyPoints(object):

    def it_supports_len(self, len_fixture):
        points, expected_value = len_fixture
        assert len(points) == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:ser', 0),
        ('c:ser/c:xVal', 0),
        ('c:ser/c:xVal/c:numRef/c:numCache/c:ptCount{val=3}', 0),
        ('c:ser/c:yVal/c:numRef/c:numCache/c:ptCount{val=3}', 0),
        ('c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=1},c:yVal/c:numRef'
         '/c:numCache/c:ptCount{val=3})', 1),
        ('c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef'
         '/c:numCache/c:ptCount{val=1})', 1),
        ('c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef'
         '/c:numCache/c:ptCount{val=3})', 3),
    ])
    def len_fixture(self, request):
        ser_cxml, expected_value = request.param
        points = XyPoints(element(ser_cxml))
        return points, expected_value
