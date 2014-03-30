# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from .shared import BaseShapeElement


class CT_Connector(BaseShapeElement):
    """
    A line/connector shape ``<p:cxnSp>`` element
    """
