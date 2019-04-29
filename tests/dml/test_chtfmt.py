# encoding: utf-8

"""
Unit test suite for the pptx.dml.chtfmt module.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.dml.chtfmt import ChartFormat
from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribeChartFormat(object):
    def it_provides_access_to_its_fill(self, fill_fixture):
        chart_format, FillFormat_, fill_, expected_xml = fill_fixture
        fill = chart_format.fill
        FillFormat_.from_fill_parent.assert_called_once_with(
            chart_format._element.xpath("c:spPr")[0]
        )
        assert fill is fill_
        assert chart_format._element.xml == expected_xml

    def it_provides_access_to_its_line(self, line_fixture):
        chart_format, LineFormat_, line_, expected_xml = line_fixture
        line = chart_format.line
        LineFormat_.assert_called_once_with(chart_format._element.xpath("c:spPr")[0])
        assert line is line_
        assert chart_format._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:catAx", "c:catAx/c:spPr"),
            ("c:catAx/c:spPr", "c:catAx/c:spPr"),
            ("c:dPt", "c:dPt/c:spPr"),
            ("c:dPt/c:spPr", "c:dPt/c:spPr"),
            ("c:majorGridlines", "c:majorGridlines/c:spPr"),
            ("c:majorGridlines/c:spPr", "c:majorGridlines/c:spPr"),
            ("c:valAx", "c:valAx/c:spPr"),
            ("c:valAx/c:spPr", "c:valAx/c:spPr"),
        ]
    )
    def fill_fixture(self, request, FillFormat_, fill_):
        dPt_cxml, expected_cxml = request.param
        chart_format = ChartFormat(element(dPt_cxml))
        FillFormat_.from_fill_parent.return_value = fill_
        expected_xml = xml(expected_cxml)
        return chart_format, FillFormat_, fill_, expected_xml

    @pytest.fixture(
        params=[
            ("c:catAx", "c:catAx/c:spPr"),
            ("c:catAx/c:spPr", "c:catAx/c:spPr"),
            ("c:dPt", "c:dPt/c:spPr"),
            ("c:dPt/c:spPr", "c:dPt/c:spPr"),
            ("c:majorGridlines", "c:majorGridlines/c:spPr"),
            ("c:majorGridlines/c:spPr", "c:majorGridlines/c:spPr"),
            ("c:valAx", "c:valAx/c:spPr"),
            ("c:valAx/c:spPr", "c:valAx/c:spPr"),
        ]
    )
    def line_fixture(self, request, LineFormat_, line_):
        cxml, expected_cxml = request.param
        chart_format = ChartFormat(element(cxml))
        expected_xml = xml(expected_cxml)
        return chart_format, LineFormat_, line_, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def FillFormat_(self, request):
        return class_mock(request, "pptx.dml.chtfmt.FillFormat")

    @pytest.fixture
    def fill_(self, request):
        return instance_mock(request, FillFormat)

    @pytest.fixture
    def LineFormat_(self, request, line_):
        return class_mock(request, "pptx.dml.chtfmt.LineFormat", return_value=line_)

    @pytest.fixture
    def line_(self, request):
        return instance_mock(request, LineFormat)
