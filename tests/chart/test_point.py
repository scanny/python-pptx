# encoding: utf-8

"""
Unit test suite for the pptx.chart.point module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.chart.point import BubblePoints, XyPoints

from ..unitutil.cxml import element


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
