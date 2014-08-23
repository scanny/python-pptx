# encoding: utf-8

"""
Test suite for pptx.oxml.graphfrm module.
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame

from ...unitutil.cxml import element, xml


CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'


class DescribeCT_GraphicalObjectFrame(object):

    def it_knows_whether_it_contains_a_chart(self, has_chart_fixture):
        graphicFrame, expected_value = has_chart_fixture
        assert graphicFrame.has_chart is expected_value

    def it_knows_whether_it_contains_a_table(self, has_table_fixture):
        graphicFrame, expected_value = has_table_fixture
        assert graphicFrame.has_table is expected_value

    def it_can_construct_a_new_graphicFrame(self, new_graphicFrame_fixture):
        id_, name, x, y, cx, cy, expected_xml = new_graphicFrame_fixture
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, x, y, cx, cy
        )
        assert graphicFrame.xml == expected_xml

    def it_can_construct_a_new_chart_graphicFrame(
            self, new_chart_graphicFrame_fixture):
        id_, name, rId, x, y, cx, cy, expected_xml = (
            new_chart_graphicFrame_fixture
        )
        graphicFrame = CT_GraphicalObjectFrame.new_chart_graphicFrame(
            id_, name, rId, x, y, cx, cy
        )
        assert graphicFrame.xml == expected_xml

    def it_can_construct_a_new_table_graphicFrame(
            self, new_table_graphicFrame_fixture):
        id_, name, rows, cols, x, y, cx, cy, expected_xml = (
            new_table_graphicFrame_fixture
        )
        graphicFrame = CT_GraphicalObjectFrame.new_table_graphicFrame(
            id_, name, rows, cols, x, y, cx, cy
        )
        assert graphicFrame.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (CHART_URI, True), (TABLE_URI, False), ('Foobar', False)
    ])
    def has_chart_fixture(self, request):
        uri, expected_value = request.param
        graphicFrame_cxml = (
            'p:graphicFrame/a:graphic/a:graphicData{uri=%s}' % uri
        )
        graphicFrame = element(graphicFrame_cxml)
        return graphicFrame, expected_value

    @pytest.fixture(params=[
        (CHART_URI, False), (TABLE_URI, True), ('Foobar', False)
    ])
    def has_table_fixture(self, request):
        uri, expected_value = request.param
        graphicFrame_cxml = (
            'p:graphicFrame/a:graphic/a:graphicData{uri=%s}' % uri
        )
        graphicFrame = element(graphicFrame_cxml)
        return graphicFrame, expected_value

    @pytest.fixture
    def new_chart_graphicFrame_fixture(self):
        id_, name, rId, x, y, cx, cy = 42, 'foobar', 'rId6', 1, 2, 3, 4
        xml_tmpl = xml(
            'p:graphicFrame/(p:nvGraphicFramePr/(p:cNvPr{id=42,name=foobar},'
            'p:cNvGraphicFramePr/a:graphicFrameLocks{noGrp=1},p:nvPr),p:xfrm'
            '/(a:off{x=1,y=2},a:ext{cx=3,cy=4}),a:graphic/a:graphicData{uri='
            '%s}"%%s")'
            % CHART_URI
        )
        expected_xml = xml_tmpl % (
            '\n      ' + xml('c:chart{r:id=rId6}') + '    '
        )
        return id_, name, rId, x, y, cx, cy, expected_xml

    @pytest.fixture
    def new_graphicFrame_fixture(self):
        id_, name, x, y, cx, cy = 42, 'foobar', 1, 2, 3, 4
        expected_xml = xml(
            'p:graphicFrame/(p:nvGraphicFramePr/(p:cNvPr{id=42,name=foobar},'
            'p:cNvGraphicFramePr/a:graphicFrameLocks{noGrp=1},p:nvPr),p:xfrm'
            '/(a:off{x=1,y=2},a:ext{cx=3,cy=4}),a:graphic/a:graphicData)'
        )
        return id_, name, x, y, cx, cy, expected_xml

    @pytest.fixture
    def new_table_graphicFrame_fixture(self):
        id_, name, rows, cols, x, y, cx, cy = 42, 'foobar', 1, 1, 1, 2, 3, 4
        expected_xml = xml(
            'p:graphicFrame/(p:nvGraphicFramePr/(p:cNvPr{id=42,name=foobar},'
            'p:cNvGraphicFramePr/a:graphicFrameLocks{noGrp=1},p:nvPr),p:xfrm'
            '/(a:off{x=1,y=2},a:ext{cx=3,cy=4}),a:graphic/a:graphicData{uri='
            '%s}/a:tbl/(a:tblPr{firstRow=1,bandRow=1}/a:tableStyleId"{5C2254'
            '4A-7EE6-4342-B048-85BDC9FD1C3A}",a:tblGrid/a:gridCol{w=3},a:tr{'
            'h=4}/a:tc/(a:txBody/(a:bodyPr,a:lstStyle,a:p),a:tcPr)))'
            % TABLE_URI
        )
        return id_, name, rows, cols, x, y, cx, cy, expected_xml
