# encoding: utf-8

"""
Test suite for pptx.chart.legend module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.legend import Legend
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.text.text import Font

from ..unitutil.cxml import element, xml


class DescribeLegend(object):
    def it_provides_access_to_its_font(self, font_fixture):
        legend, expected_xml = font_fixture
        font = legend.font
        assert legend._element.xml == expected_xml
        assert isinstance(font, Font)
        assert font._element == legend._element.xpath(".//a:defRPr")[0]

    def it_knows_its_horizontal_offset(self, horz_offset_get_fixture):
        legend, expected_value = horz_offset_get_fixture
        assert legend.horz_offset == expected_value

    def it_can_change_its_horizontal_offset(self, horz_offset_set_fixture):
        legend, new_value, expected_xml = horz_offset_set_fixture
        legend.horz_offset = new_value
        assert legend._element.xml == expected_xml

    def it_knows_whether_it_should_overlap_the_chart(
        self, include_in_layout_get_fixture
    ):
        legend, expected_value = include_in_layout_get_fixture
        assert legend.include_in_layout == expected_value

    def it_can_change_whether_it_overlaps_the_chart(
        self, include_in_layout_set_fixture
    ):
        legend, new_value, expected_xml = include_in_layout_set_fixture
        legend.include_in_layout = new_value
        assert legend._element.xml == expected_xml

    def it_knows_its_position_with_respect_to_the_chart(self, position_get_fixture):
        legend, expected_value = position_get_fixture
        assert legend.position == expected_value

    def it_can_change_its_position(self, position_set_fixture):
        legend, new_value, expected_xml = position_set_fixture
        legend.position = new_value
        assert legend._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (
                "c:legend{a:b=c}",
                "c:legend{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)",
            )
        ]
    )
    def font_fixture(self, request):
        legend_cxml, expected_cxml = request.param
        legend = Legend(element(legend_cxml))
        expected_xml = xml(expected_cxml)
        return legend, expected_xml

    @pytest.fixture(
        params=[
            ("c:legend", 0.0),
            ("c:legend/c:layout", 0.0),
            ("c:legend/c:layout/c:manualLayout", 0.0),
            ("c:legend/c:layout/c:manualLayout/c:xMode", 0.0),
            ("c:legend/c:layout/c:manualLayout/c:xMode{val=edge}", 0.0),
            ("c:legend/c:layout/c:manualLayout/c:xMode{val=factor}", 0.0),
            ("c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.42})", 0.42),
            (
                "c:legend/c:layout/c:manualLayout/(c:xMode{val=factor},c:x{val=0.42"
                "})",
                0.42,
            ),
        ]
    )
    def horz_offset_get_fixture(self, request):
        legend_cxml, expected_value = request.param
        legend = Legend(element(legend_cxml))
        return legend, expected_value

    @pytest.fixture(
        params=[
            (
                "c:legend",
                0.42,
                "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.42})",
            ),
            (
                "c:legend/c:layout",
                0.42,
                "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.42})",
            ),
            (
                "c:legend/c:layout/c:manualLayout",
                -0.42,
                "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=-0.42})",
            ),
            (
                "c:legend/c:layout/c:manualLayout/c:xMode{val=edge}",
                -0.1,
                "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=-0.1})",
            ),
            (
                "c:legend/c:layout/c:manualLayout/c:xMode{val=factor}",
                0.2,
                "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.2})",
            ),
            ("c:legend/c:layout/c:manualLayout", 0, "c:legend/c:layout"),
        ]
    )
    def horz_offset_set_fixture(self, request):
        legend_cxml, new_value, expected_legend_cxml = request.param
        legend = Legend(element(legend_cxml))
        expected_xml = xml(expected_legend_cxml)
        return legend, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:legend", True),
            ("c:legend/c:overlay", True),
            ("c:legend/c:overlay{val=1}", True),
            ("c:legend/c:overlay{val=0}", False),
        ]
    )
    def include_in_layout_get_fixture(self, request):
        legend_cxml, expected_value = request.param
        legend = Legend(element(legend_cxml))
        return legend, expected_value

    @pytest.fixture(
        params=[
            ("c:legend", True, "c:legend/c:overlay{val=1}"),
            ("c:legend", False, "c:legend/c:overlay{val=0}"),
            ("c:legend/c:overlay{val=0}", True, "c:legend/c:overlay{val=1}"),
            ("c:legend/c:overlay{val=1}", False, "c:legend/c:overlay{val=0}"),
            ("c:legend/c:overlay{val=1}", None, "c:legend"),
        ]
    )
    def include_in_layout_set_fixture(self, request):
        legend_cxml, new_value, expected_legend_cxml = request.param
        legend = Legend(element(legend_cxml))
        expected_xml = xml(expected_legend_cxml)
        return legend, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:legend", "RIGHT"),
            ("c:legend/c:legendPos", "RIGHT"),
            ("c:legend/c:legendPos{val=r}", "RIGHT"),
            ("c:legend/c:legendPos{val=b}", "BOTTOM"),
        ]
    )
    def position_get_fixture(self, request):
        legend_cxml, expected_enum_member = request.param
        legend = Legend(element(legend_cxml))
        expected_value = getattr(XL_LEGEND_POSITION, expected_enum_member)
        return legend, expected_value

    @pytest.fixture(
        params=[
            ("c:legend/c:legendPos{val=r}", "BOTTOM", "c:legend/c:legendPos{val=b}"),
            ("c:legend/c:legendPos{val=b}", "RIGHT", "c:legend/c:legendPos"),
            ("c:legend", "TOP", "c:legend/c:legendPos{val=t}"),
            ("c:legend/c:legendPos", "CORNER", "c:legend/c:legendPos{val=tr}"),
        ]
    )
    def position_set_fixture(self, request):
        legend_cxml, new_enum_member, expected_legend_cxml = request.param
        legend = Legend(element(legend_cxml))
        new_value = getattr(XL_LEGEND_POSITION, new_enum_member)
        expected_xml = xml(expected_legend_cxml)
        return legend, new_value, expected_xml
