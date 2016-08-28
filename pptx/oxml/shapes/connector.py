# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from .shared import BaseShapeElement
from ..xmlchemy import BaseOxmlElement, OneAndOnlyOne


class CT_Connector(BaseShapeElement):
    """
    A line/connector shape ``<p:cxnSp>`` element
    """
    _tag_seq = (
        'p:nvCxnSpPr', 'p:spPr', 'p:style', 'p:extLst'
    )
    spPr = OneAndOnlyOne('p:spPr')
    del _tag_seq


class CT_ConnectorNonVisual(BaseOxmlElement):
    """
    ``<p:nvCxnSpPr>`` element, container for the non-visual properties of
    a connector, such as name, id, etc.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
