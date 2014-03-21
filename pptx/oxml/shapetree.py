# encoding: utf-8

"""
lxml custom element classes for shape tree-related XML elements.
"""

from __future__ import absolute_import

from pptx.oxml.ns import qn
from pptx.oxml.shapes.shared import BaseShapeElement


class CT_GroupShape(BaseShapeElement):
    """
    Used for the shape tree (``<p:spTree>``) element as well as the group
    shape (``<p:grpSp>``) element.
    """

    _shape_tags = (
        qn('p:sp'), qn('p:grpSp'), qn('p:graphicFrame'), qn('p:cxnSp'),
        qn('p:pic'), qn('p:contentPart')
    )

    def iter_shape_elms(self):
        """
        Generate each child of this ``<p:spTree>`` element that corresponds
        to a shape, in the sequence they appear in the XML.
        """
        for elm in self.iterchildren():
            if elm.tag in self._shape_tags:
                yield elm
