# encoding: utf-8

"""
lxml custom element class for CT_GraphicalObjectFrame XML element.
"""

from __future__ import absolute_import

from .. import parse_xml
from ..ns import nsdecls
from .shared import BaseShapeElement
from ..simpletypes import XsdString
from .table import CT_Table
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne
)


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
    uri = RequiredAttribute('uri', XsdString)


class CT_GraphicalObjectFrame(BaseShapeElement):
    """
    ``<p:graphicFrame>`` element, which is a container for a table, a chart,
    or another graphical object.
    """
    nvGraphicFramePr = OneAndOnlyOne('p:nvGraphicFramePr')
    xfrm = OneAndOnlyOne('p:xfrm')
    graphic = OneAndOnlyOne('a:graphic')

    DATATYPE_TABLE = 'http://schemas.openxmlformats.org/drawingml/2006/table'

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
        datatype = self.graphic.graphicData.uri
        if datatype == CT_GraphicalObjectFrame.DATATYPE_TABLE:
            return True
        return False

    @classmethod
    def new_graphicFrame(cls, id_, name, left, top, width, height):
        """
        Return a new ``<p:graphicFrame>`` element tree suitable for
        containing a table or chart. Note that a graphicFrame element is not
        a valid shape until it contains a graphical object such as a table.
        """
        xml = cls._graphicFrame_tmpl() % (id_, name, left, top, width, height)
        graphicFrame = parse_xml(xml)
        return graphicFrame

    @classmethod
    def new_table(cls, id_, name, rows, cols, left, top, width, height):
        """
        Return a ``<p:graphicFrame>`` element tree populated with a table
        element.
        """
        graphicFrame = cls.new_graphicFrame(
            id_, name, left, top, width, height
        )
        # set type of contained graphic to table
        graphicData = graphicFrame.graphic.graphicData
        graphicData.uri = cls.DATATYPE_TABLE

        # add tbl element tree
        tbl = CT_Table.new_tbl(rows, cols, width, height)
        graphicData.append(tbl)

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
