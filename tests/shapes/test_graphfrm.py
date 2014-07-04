# encoding: utf-8

"""
Test suite for pptx.graphfrm module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.parts.slide import _SlideShapeTree
from pptx.parts.chart import ChartPart
from pptx.shapes.graphfrm import GraphicFrame
from pptx.spec import GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_TABLE

from ..unitutil.cxml import element
from ..unitutil.mock import instance_mock


class DescribeGraphicFrame(object):

    def it_knows_if_it_contains_a_chart(self, has_chart_fixture):
        graphic_frame, expected_value = has_chart_fixture
        assert graphic_frame.has_chart is expected_value

    def it_knows_if_it_contains_a_table(self, has_table_fixture):
        graphic_frame, expected_value = has_table_fixture
        assert graphic_frame.has_table is expected_value

    def it_knows_its_shape_type(self, type_fixture):
        graphic_frame, expected_type = type_fixture
        assert graphic_frame.shape_type is expected_type

    def it_can_retrieve_a_linked_chart_part(self, chart_part_fixture):
        graphic_frame, chart_part_ = chart_part_fixture
        chart_part = graphic_frame.chart_part
        assert chart_part is chart_part_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def chart_part_fixture(self, parent_, chart_part_):
        graphicFrame_cxml = (
            'p:graphicFrame/a:graphic/a:graphicData/c:chart{r:id=rId42}'
        )
        graphic_frame = GraphicFrame(element(graphicFrame_cxml), parent_)
        return graphic_frame, chart_part_

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

    @pytest.fixture(params=[
        (GRAPHIC_DATA_URI_CHART, MSO_SHAPE_TYPE.CHART),
        (GRAPHIC_DATA_URI_TABLE, MSO_SHAPE_TYPE.TABLE),
        ('foobar',               None),
    ])
    def type_fixture(self, request):
        uri, expected_value = request.param
        graphicFrame = element(
            'p:graphicFrame/a:graphic/a:graphicData{uri=%s}' % uri
        )
        graphic_frame = GraphicFrame(graphicFrame, None)
        return graphic_frame, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_part_(self, request):
        return instance_mock(request, ChartPart)

    @pytest.fixture
    def parent_(self, request, chart_part_):
        parent_ = instance_mock(request, _SlideShapeTree)
        parent_.part.related_parts = {'rId42': chart_part_}
        return parent_
