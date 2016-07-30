# encoding: utf-8

"""
Unit test suite for the pptx.chart.marker module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.chart.marker import Marker
from pptx.dml.chtfmt import ChartFormat

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribeMarker(object):

    def it_provides_access_to_its_format(self, format_fixture):
        marker, ChartFormat_, _element, chart_format_, expected_xml = (
            format_fixture
        )
        chart_format = marker.format
        ChartFormat_.assert_called_once_with(_element.xpath('c:marker')[0])
        assert chart_format is chart_format_
        assert marker._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:ser',          'c:ser/c:marker'),
        ('c:ser/c:marker', 'c:ser/c:marker'),
        ('c:dPt',          'c:dPt/c:marker'),
        ('c:dPt/c:marker', 'c:dPt/c:marker'),
    ])
    def format_fixture(self, request, ChartFormat_, chart_format_):
        cxml, expected_cxml = request.param
        _element = element(cxml)
        marker = Marker(_element)
        expected_xml = xml(expected_cxml)
        return marker, ChartFormat_, _element, chart_format_, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartFormat_(self, request, chart_format_):
        return class_mock(
            request, 'pptx.chart.marker.ChartFormat',
            return_value=chart_format_
        )

    @pytest.fixture
    def chart_format_(self, request):
        return instance_mock(request, ChartFormat)
