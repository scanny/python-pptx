# encoding: utf-8

"""
lxml custom element class for CT_GraphicalObjectFrame XML element.
"""

from __future__ import absolute_import

from .. import parse_xml
from ..ns import nsdecls, qn
from .shared import BaseShapeElement
from .table import CT_Table
from ..xmlchemy import BaseOxmlElement, OneAndOnlyOne, ZeroOrOne


class CT_GraphicalObject(BaseOxmlElement):
    """
    ``<a:graphic>`` element, which is the container for the reference to or
    definition of the framed graphical object.
    """
    graphicData = OneAndOnlyOne('a:graphicData')


class CT_GraphicalObjectData(BaseShapeElement):
    """
    ``<p:graphicData>`` element, the direct container for a table, a chart,
    or another graphical object.
    """
    tbl = ZeroOrOne('a:tbl')


class CT_GraphicalObjectFrame(BaseShapeElement):
    """
    ``<p:graphicFrame>`` element, which is a container for a table, a chart,
    or another graphical object.
    """
    nvGraphicFramePr = OneAndOnlyOne('p:nvGraphicFramePr')
    graphic = OneAndOnlyOne('a:graphic')

    DATATYPE_TABLE = 'http://schemas.openxmlformats.org/drawingml/2006/table'

    _graphicFrame_tmpl = (
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

    def get_or_add_xfrm(self):
        """
        Return the required ``<p:xfrm>`` child element. Overrides version on
        BaseShapeElement.
        """
        return self.xfrm

    @property
    def has_table(self):
        """
        True if graphicFrame contains a table, False otherwise.
        """
        datatype = self.graphic.graphicData.get('uri')
        if datatype == CT_GraphicalObjectFrame.DATATYPE_TABLE:
            return True
        return False

    @staticmethod
    def new_graphicFrame(id_, name, left, top, width, height):
        """
        Return a new ``<p:graphicFrame>`` element tree suitable for containing
        a table or chart. Note that a graphicFrame element is not a valid
        shape until it contains a graphical object such as a table.
        """
        xml = CT_GraphicalObjectFrame._graphicFrame_tmpl % (
            id_, name, left, top, width, height)
        graphicFrame = parse_xml(xml)
        return graphicFrame

    @staticmethod
    def new_table(id_, name, rows, cols, left, top, width, height):
        """
        Return a ``<p:graphicFrame>`` element tree populated with a table
        element.
        """
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, left, top, width, height)

        # set type of contained graphic to table
        graphicData = graphicFrame.graphic.graphicData
        graphicData.set('uri', CT_GraphicalObjectFrame.DATATYPE_TABLE)

        # add tbl element tree
        tbl = CT_Table.new_tbl(rows, cols, width, height)
        graphicData.append(tbl)

        return graphicFrame

    @property
    def xfrm(self):
        """
        The required ``<p:xfrm>`` child element
        """
        return self.find(qn('p:xfrm'))


class CT_GraphicalObjectFrameNonVisual(BaseOxmlElement):
    """
    ``<p:nvGraphicFramePr>`` element, container for the non-visual properties
    of a graphic frame, such as name, id, etc.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
