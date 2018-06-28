# encoding: utf-8

"""
Test suite for pptx.shapes.graphfrm module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.chart import Chart
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from pptx.parts.chart import ChartPart
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.shapetree import SlideShapes
from pptx.spec import GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_TABLE

from ..unitutil.cxml import element
from ..unitutil.mock import instance_mock, property_mock


class DescribeGraphicFrame(object):

    def it_knows_if_it_contains_a_chart(self, has_chart_fixture):
        graphic_frame, expected_value = has_chart_fixture
        assert graphic_frame.has_chart is expected_value

    def it_knows_if_it_contains_a_table(self, has_table_fixture):
        graphic_frame, expected_value = has_table_fixture
        assert graphic_frame.has_table is expected_value

    def it_raises_on_shadow(self):
        graphic_frame = GraphicFrame(None, None)
        with pytest.raises(NotImplementedError):
            graphic_frame.shadow

    def it_knows_its_shape_type(self, type_fixture):
        graphic_frame, expected_type = type_fixture
        assert graphic_frame.shape_type is expected_type

    def it_can_retrieve_a_linked_chart_part(self, chart_part_fixture):
        graphic_frame, chart_part_ = chart_part_fixture
        chart_part = graphic_frame.chart_part
        assert chart_part is chart_part_

    def it_provides_access_to_the_chart_it_contains(self, chart_fixture):
        graphic_frame, chart_ = chart_fixture
        chart = graphic_frame.chart
        assert chart is chart_

    def it_raises_on_chart_if_there_isnt_one(self, chart_raise_fixture):
        graphic_frame = chart_raise_fixture
        with pytest.raises(ValueError):
            graphic_frame.chart

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def chart_fixture(self, request, chart_part_, graphicFrame_, chart_):
        property_mock(
            request, GraphicFrame, 'chart_part', return_value=chart_part_
        )
        graphic_frame = GraphicFrame(graphicFrame_, None)
        return graphic_frame, chart_

    @pytest.fixture
    def chart_raise_fixture(self, request, graphicFrame_):
        graphicFrame_.has_chart = False
        graphic_frame = GraphicFrame(graphicFrame_, None)
        return graphic_frame

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
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def chart_part_(self, request, chart_):
        return instance_mock(request, ChartPart, chart=chart_)

    @pytest.fixture
    def graphicFrame_(self, request):
        return instance_mock(
            request, CT_GraphicalObjectFrame, has_chart=True
        )

    @pytest.fixture
    def parent_(self, request, chart_part_):
        parent_ = instance_mock(request, SlideShapes)
        parent_.part.related_parts = {'rId42': chart_part_}
        return parent_
