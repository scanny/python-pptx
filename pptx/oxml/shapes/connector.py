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
    spPr = OneAndOnlyOne('p:spPr')


class CT_ConnectorNonVisual(BaseOxmlElement):
    """
    ``<p:nvCxnSpPr>`` element, container for the non-visual properties of
    a connector, such as name, id, etc.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
