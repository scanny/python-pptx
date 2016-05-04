# encoding: utf-8

"""
Unit test suite for the pptx.chart.point module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.chart.datalabel import DataLabel
from pptx.chart.point import BubblePoints, Point, XyPoints

from ..unitutil.cxml import element
from ..unitutil.mock import function_mock, instance_mock


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
        return function_mock(
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


class DescribePoint(object):

    def it_provides_access_to_its_data_label(self, data_label_fixture):
        point, DataLabel_, ser, idx, data_label_ = data_label_fixture
        data_label = point.data_label
        DataLabel_.assert_called_once_with(ser, idx)
        assert data_label is data_label_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def data_label_fixture(self, DataLabel_, data_label_):
        ser, idx = element('c:ser'), 42
        point = Point(ser, idx)
        return point, DataLabel_, ser, idx, data_label_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def DataLabel_(self, request, data_label_):
        return function_mock(
            request, 'pptx.chart.point.DataLabel', return_value=data_label_
        )

    @pytest.fixture
    def data_label_(self, request):
        return instance_mock(request, DataLabel)


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
