# encoding: utf-8

"""
lxml custom element classes for shape tree-related XML elements.
"""

from __future__ import absolute_import

from .ns import qn
from .picture import CT_Picture
from .shapes.shared import BaseShapeElement


class CT_GroupShape(BaseShapeElement):
    """
    Used for the shape tree (``<p:spTree>``) element as well as the group
    shape (``<p:grpSp>``) element.
    """

    _shape_tags = (
        qn('p:sp'), qn('p:grpSp'), qn('p:graphicFrame'), qn('p:cxnSp'),
        qn('p:pic'), qn('p:contentPart')
    )

    def add_pic(self, id, name, desc, rId, x, y, cx, cy):
        """
        Append a ``<p:pic>`` shape to the group/shapetree having properties
        as specified in call.
        """
        pic = CT_Picture.new_pic(id, name, desc, rId, x, y, cx, cy)
        self.insert_element_before(pic, 'p:extLst')
        return pic

    def iter_shape_elms(self):
        """
        Generate each child of this ``<p:spTree>`` element that corresponds
        to a shape, in the sequence they appear in the XML.
        """
        for elm in self.iterchildren():
            if elm.tag in self._shape_tags:
                yield elm
