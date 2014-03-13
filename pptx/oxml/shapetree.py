# encoding: utf-8

"""
lxml custom element classes for shape tree-related XML elements.
"""

from __future__ import absolute_import

from pptx.oxml.core import BaseOxmlElement


class CT_GroupShape(BaseOxmlElement):
    """
    Used for the shape tree (``<p:spTree>``) element as well as the group
    shape (``<p:grpSp>``) element.
    """
