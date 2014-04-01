# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from ..ns import qn
from .shared import BaseShapeElement


class CT_Connector(BaseShapeElement):
    """
    A line/connector shape ``<p:cxnSp>`` element
    """
    @property
    def spPr(self):
        return self.find(qn('p:spPr'))
