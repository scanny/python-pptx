# encoding: utf-8

"""
lxml custom element classes for shape tree-related XML elements.
"""

from __future__ import absolute_import

from pptx.oxml.shapes.shared import BaseShapeElement


class CT_GroupShape(BaseShapeElement):
    """
    Used for the shape tree (``<p:spTree>``) element as well as the group
    shape (``<p:grpSp>``) element.
    """
