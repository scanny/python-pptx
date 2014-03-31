# encoding: utf-8

"""
lxml custom element class for CT_GraphicalObjectFrame XML element.
"""

from __future__ import absolute_import

from lxml import objectify

from .. import parse_xml_bytes
from ..ns import nsdecls, qn
from .shared import BaseShapeElement
from .table import CT_Table


class CT_GraphicalObjectFrame(BaseShapeElement):
    """
    ``<p:graphicFrame>`` element, which is a container for a table, a chart,
    or another graphical object.
    """
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

    @property
    def has_table(self):
        """True if graphicFrame contains a table, False otherwise"""
        datatype = self[qn('a:graphic')].graphicData.get('uri')
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
        graphicFrame = parse_xml_bytes(xml)

        objectify.deannotate(graphicFrame, cleanup_namespaces=True)
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
        graphicData = graphicFrame[qn('a:graphic')].graphicData
        graphicData.set('uri', CT_GraphicalObjectFrame.DATATYPE_TABLE)

        # add tbl element tree
        tbl = CT_Table.new_tbl(rows, cols, width, height)
        graphicData.append(tbl)

        objectify.deannotate(graphicFrame, cleanup_namespaces=True)
        return graphicFrame

    @property
    def xfrm(self):
        """
        The required ``<p:xfrm>`` child element
        """
        return self.find(qn('p:xfrm'))
