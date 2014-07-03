# encoding: utf-8

"""
Test suite for pptx.graphfrm module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.shapes.graphfrm import GraphicFrame
from pptx.spec import GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_TABLE

from ..unitutil.cxml import element


class DescribeGraphicFrame(object):

    def it_knows_if_it_contains_a_chart(self, has_chart_fixture):
        graphic_frame, expected_value = has_chart_fixture
        assert graphic_frame.has_chart is expected_value

    def it_knows_if_it_contains_a_table(self, has_table_fixture):
        graphic_frame, expected_value = has_table_fixture
        assert graphic_frame.has_table is expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (GRAPHIC_DATA_URI_CHART, True),
        (GRAPHIC_DATA_URI_TABLE, False),
    ])
    def has_chart_fixture(self, request):
        uri, expected_value = request.param
        graphicFrame = element(
            'p:graphicFrame/a:graphic/a:graphicData{uri=%s}' % uri
        )
        graphic_frame = GraphicFrame(graphicFrame, None)
        return graphic_frame, expected_value

    @pytest.fixture(params=[
        (GRAPHIC_DATA_URI_CHART, False),
        (GRAPHIC_DATA_URI_TABLE, True),
    ])
    def has_table_fixture(self, request):
        uri, expected_value = request.param
        graphicFrame = element(
            'p:graphicFrame/a:graphic/a:graphicData{uri=%s}' % uri
        )
        graphic_frame = GraphicFrame(graphicFrame, None)
        return graphic_frame, expected_value
