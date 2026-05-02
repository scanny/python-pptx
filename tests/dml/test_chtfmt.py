"""Unit-test suite for `pptx.dml.chtfmt` module."""

from __future__ import annotations

import pytest

from pptx.dml.chtfmt import ChartFormat
from pptx.dml.effect import ShadowFormat
from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat
from pptx.util import Emu

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

    def it_provides_access_to_its_shadow(self):
        chart_format = ChartFormat(element("c:catAx"))
        shadow = chart_format.shadow
        assert isinstance(shadow, ShadowFormat)
        # -- shadow is anchored on a newly-materialized c:spPr --
        assert chart_format._element.xpath("c:spPr")

    def it_memoizes_its_shadow(self):
        chart_format = ChartFormat(element("c:catAx"))
        assert chart_format.shadow is chart_format.shadow

    def it_exposes_the_full_ShadowFormat_api_through_its_shadow(self):
        """Issue #130: ChartFormat.shadow exposes full blur/distance/direction/color."""
        chart_format = ChartFormat(element("c:catAx"))
        shadow = chart_format.shadow

        # -- before any writes, all four properties report the unconfigured state --
        assert shadow.inherit is True
        assert shadow.blur_radius is None
        assert shadow.distance is None
        assert shadow.direction is None

        # -- each attribute can be written independently --
        shadow.blur_radius = Emu(50800)
        shadow.distance = Emu(38100)
        shadow.direction = 90.0

        assert shadow.blur_radius == Emu(50800)
        assert shadow.distance == Emu(38100)
        assert shadow.direction == 90.0
        # -- first access to `.color` materializes a default black srgbClr --
        from pptx.dml.color import ColorFormat

        assert isinstance(shadow.color, ColorFormat)

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
