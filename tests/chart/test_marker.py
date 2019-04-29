# encoding: utf-8

"""
Unit test suite for the pptx.chart.marker module.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.chart.marker import Marker
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_MARKER_STYLE

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribeMarker(object):
    def it_knows_its_size(self, size_get_fixture):
        marker, expected_value = size_get_fixture
        assert marker.size == expected_value

    def it_can_change_its_size(self, size_set_fixture):
        marker, value, expected_xml = size_set_fixture
        marker.size = value
        assert marker._element.xml == expected_xml

    def it_knows_its_style(self, style_get_fixture):
        marker, expected_value = style_get_fixture
        assert marker.style == expected_value

    def it_can_change_its_style(self, style_set_fixture):
        marker, value, expected_xml = style_set_fixture
        marker.style = value
        assert marker._element.xml == expected_xml

    def it_provides_access_to_its_format(self, format_fixture):
        marker, ChartFormat_, _element, chart_format_, expected_xml = format_fixture
        chart_format = marker.format
        ChartFormat_.assert_called_once_with(_element.xpath("c:marker")[0])
        assert chart_format is chart_format_
        assert marker._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:ser", "c:ser/c:marker"),
            ("c:ser/c:marker", "c:ser/c:marker"),
            ("c:dPt", "c:dPt/c:marker"),
            ("c:dPt/c:marker", "c:dPt/c:marker"),
        ]
    )
    def format_fixture(self, request, ChartFormat_, chart_format_):
        cxml, expected_cxml = request.param
        _element = element(cxml)
        marker = Marker(_element)
        expected_xml = xml(expected_cxml)
        return marker, ChartFormat_, _element, chart_format_, expected_xml

    @pytest.fixture(
        params=[
            ("c:ser", None),
            ("c:ser/c:marker", None),
            ("c:ser/c:marker/c:size{val=24}", 24),
            ("c:dPt", None),
            ("c:dPt/c:marker", None),
            ("c:dPt/c:marker/c:size{val=36}", 36),
        ]
    )
    def size_get_fixture(self, request):
        cxml, expected_value = request.param
        marker = Marker(element(cxml))
        return marker, expected_value

    @pytest.fixture(
        params=[
            ("c:ser", 42, "c:ser/c:marker/c:size{val=42}"),
            ("c:ser/c:marker", 42, "c:ser/c:marker/c:size{val=42}"),
            ("c:ser/c:marker/c:size{val=42}", 24, "c:ser/c:marker/c:size{val=24}"),
            ("c:ser/c:marker/c:size{val=24}", None, "c:ser/c:marker"),
            ("c:ser", None, "c:ser/c:marker"),
        ]
    )
    def size_set_fixture(self, request):
        cxml, value, expected_cxml = request.param
        marker = Marker(element(cxml))
        expected_xml = xml(expected_cxml)
        return marker, value, expected_xml

    @pytest.fixture(
        params=[
            ("c:ser", None),
            ("c:ser/c:marker", None),
            ("c:ser/c:marker/c:symbol{val=circle}", XL_MARKER_STYLE.CIRCLE),
            ("c:dPt", None),
            ("c:dPt/c:marker", None),
            ("c:dPt/c:marker/c:symbol{val=square}", XL_MARKER_STYLE.SQUARE),
        ]
    )
    def style_get_fixture(self, request):
        cxml, expected_value = request.param
        marker = Marker(element(cxml))
        return marker, expected_value

    @pytest.fixture(
        params=[
            ("c:ser", XL_MARKER_STYLE.CIRCLE, "c:ser/c:marker/c:symbol{val=circle}"),
            (
                "c:ser/c:marker",
                XL_MARKER_STYLE.SQUARE,
                "c:ser/c:marker/c:symbol{val=square}",
            ),
            (
                "c:ser/c:marker/c:symbol{val=square}",
                XL_MARKER_STYLE.AUTOMATIC,
                "c:ser/c:marker/c:symbol{val=auto}",
            ),
            ("c:ser/c:marker/c:symbol{val=auto}", None, "c:ser/c:marker"),
            ("c:ser", None, "c:ser/c:marker"),
        ]
    )
    def style_set_fixture(self, request):
        cxml, value, expected_cxml = request.param
        marker = Marker(element(cxml))
        expected_xml = xml(expected_cxml)
        return marker, value, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartFormat_(self, request, chart_format_):
        return class_mock(
            request, "pptx.chart.marker.ChartFormat", return_value=chart_format_
        )

    @pytest.fixture
    def chart_format_(self, request):
        return instance_mock(request, ChartFormat)
