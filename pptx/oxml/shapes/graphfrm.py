# encoding: utf-8

"""
lxml custom element class for CT_GraphicalObjectFrame XML element.
"""

from __future__ import absolute_import

from .. import parse_xml
from ..chart.chart import CT_Chart
from ..ns import nsdecls
from .shared import BaseShapeElement
from ..simpletypes import XsdString
from ...spec import GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_TABLE
from ..table import CT_Table
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne
)


class CT_GraphicalObject(BaseOxmlElement):
    """
    ``<a:graphic>`` element, which is the container for the reference to or
    definition of the framed graphical object (table, chart, etc.).
    """
    graphicData = OneAndOnlyOne('a:graphicData')

    @property
    def chart(self):
        """
        The ``<c:chart>`` grandchild element, or |None| if not present.
        """
        return self.graphicData.chart


class CT_GraphicalObjectData(BaseShapeElement):
    """
    ``<p:graphicData>`` element, the direct container for a table, a chart,
    or another graphical object.
    """
    chart = ZeroOrOne('c:chart')
    tbl = ZeroOrOne('a:tbl')
    uri = RequiredAttribute('uri', XsdString)


class CT_GraphicalObjectFrame(BaseShapeElement):
    """
    ``<p:graphicFrame>`` element, which is a container for a table, a chart,
    or another graphical object.
    """
    nvGraphicFramePr = OneAndOnlyOne('p:nvGraphicFramePr')
    xfrm = OneAndOnlyOne('p:xfrm')
    graphic = OneAndOnlyOne('a:graphic')

    @property
    def chart(self):
        """
        The ``<c:chart>`` great-grandchild element, or |None| if not present.
        """
        return self.graphic.chart

    @property
    def chart_rId(self):
        """
        The ``rId`` attribute of the ``<c:chart>`` great-grandchild element,
        or |None| if not present.
        """
        chart = self.chart
        if chart is None:
            return None
        return chart.rId

    def get_or_add_xfrm(self):
        """
        Return the required ``<p:xfrm>`` child element. Overrides version on
        BaseShapeElement.
        """
        return self.xfrm

    @property
    def has_chart(self):
        """
        True if graphicFrame contains a chart, False otherwise.
        """
        return self.graphic.graphicData.uri == GRAPHIC_DATA_URI_CHART

    @property
    def has_table(self):
        """
        True if graphicFrame contains a table, False otherwise.
        """
        return self.graphic.graphicData.uri == GRAPHIC_DATA_URI_TABLE

    @classmethod
    def new_chart_graphicFrame(cls, id_, name, rId, x, y, cx, cy):
        """
        Return a ``<p:graphicFrame>`` element tree populated with a chart
        element.
        """
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, x, y, cx, cy
        )
        graphicData = graphicFrame.graphic.graphicData
        graphicData.uri = GRAPHIC_DATA_URI_CHART
        graphicData.append(CT_Chart.new_chart(rId))
        return graphicFrame

    @classmethod
    def new_graphicFrame(cls, id_, name, x, y, cx, cy):
        """
        Return a new ``<p:graphicFrame>`` element tree suitable for
        containing a table or chart. Note that a graphicFrame element is not
        a valid shape until it contains a graphical object such as a table.
        """
        xml = cls._graphicFrame_tmpl() % (id_, name, x, y, cx, cy)
        graphicFrame = parse_xml(xml)
        return graphicFrame

    @classmethod
    def new_table_graphicFrame(cls, id_, name, rows, cols, x, y, cx, cy):
        """
        Return a ``<p:graphicFrame>`` element tree populated with a table
        element.
        """
        graphicFrame = cls.new_graphicFrame(id_, name, x, y, cx, cy)
        graphicFrame.graphic.graphicData.uri = GRAPHIC_DATA_URI_TABLE
        graphicFrame.graphic.graphicData.append(
            CT_Table.new_tbl(rows, cols, cx, cy)
        )
        return graphicFrame

    @classmethod
    def _graphicFrame_tmpl(cls):
        return (
            '<p:graphicFrame %s>\n'
            '  <p:nvGraphicFramePr>\n'
            '    <p:cNvPr id="%s" name="%s"/>\n'
            '    <p:cNvGraphicFramePr>\n'
            '      <a:graphicFrameLocks noGrp="1"/>\n'
            '    </p:cNvGraphicFramePr>\n'
            '    <p:nvPr/>\n'
            '  </p:nvGraphicFramePr>\n'
            '  <p:xfrm>\n'
            '    <a:off x="%s" y="%s"/>\n'
            '    <a:ext cx="%s" cy="%s"/>\n'
            '  </p:xfrm>\n'
            '  <a:graphic>\n'
            '    <a:graphicData/>\n'
            '  </a:graphic>\n'
            '</p:graphicFrame>' %
            (nsdecls('a', 'p'), '%d', '%s', '%d', '%d', '%d', '%d')
        )


class CT_GraphicalObjectFrameNonVisual(BaseOxmlElement):
    """
    ``<p:nvGraphicFramePr>`` element, container for the non-visual properties
    of a graphic frame, such as name, id, etc.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
    nvPr = OneAndOnlyOne('p:nvPr')
