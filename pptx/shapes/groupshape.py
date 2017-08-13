# encoding: utf-8

"""
lxml custom element classes for shape tree-related XML elements.
"""

from __future__ import absolute_import

from .autoshape import CT_Shape
from .connector import CT_Connector
from ...enum.shapes import MSO_CONNECTOR_TYPE
from .graphfrm import CT_GraphicalObjectFrame
from ..ns import qn
from .picture import CT_Picture
from .shared import BaseShapeElement
from ..xmlchemy import BaseOxmlElement, OneAndOnlyOne, ZeroOrOne


class CT_GroupShape(BaseShapeElement):
    """
    Used for the shape tree (``<p:spTree>``) element as well as the group
    shape (``<p:grpSp>``) element.
    """
    nvGrpSpPr = OneAndOnlyOne('p:nvGrpSpPr')
    grpSpPr = OneAndOnlyOne('p:grpSpPr')

    _shape_tags = (
        qn('p:sp'), qn('p:grpSp'), qn('p:graphicFrame'), qn('p:cxnSp'),
        qn('p:pic'), qn('p:contentPart')
    )

    def add_autoshape(self, id_, name, prst, x, y, cx, cy):
        """
        Append a new ``<p:sp>`` shape to the group/shapetree having the
        properties specified in call.
        """
        sp = CT_Shape.new_autoshape_sp(id_, name, prst, x, y, cx, cy)
        print(sp)
        self.insert_element_before(sp, 'p:extLst')
        return sp

    def add_cloned_shape(self, ids_, names, element, grouped, x=None, y=None, cx=None, cy=None):
        """
        Append a new ``<p:sp>`` shape whose element exactly mathces *element* to the group/shapetree.
        """
        sp = self.update_infos(ids_, names, element, grouped, x, y, cx, cy)
        self.insert_element_before(sp, 'p:extLst')
        return sp

    def add_cxnSp(self, id_, name, type_member, x, y, cx, cy, flipH, flipV):
        """
        Append a new ``<p:cxnSp>`` shape to the group/shapetree having the
        properties specified in call.
        """
        prst = MSO_CONNECTOR_TYPE.to_xml(type_member)
        cxnSp = CT_Connector.new_cxnSp(
            id_, name, prst, x, y, cx, cy, flipH, flipV
        )
        self.insert_element_before(cxnSp, 'p:extLst')
        return cxnSp

    def add_pic(self, id_, name, desc, rId, x, y, cx, cy):
        """
        Append a ``<p:pic>`` shape to the group/shapetree having properties
        as specified in call.
        """
        pic = CT_Picture.new_pic(id_, name, desc, rId, x, y, cx, cy)
        self.insert_element_before(pic, 'p:extLst')
        return pic

    def add_placeholder(self, id_, name, ph_type, orient, sz, idx):
        """
        Append a newly-created placeholder ``<p:sp>`` shape having the
        specified placeholder properties.
        """
        sp = CT_Shape.new_placeholder_sp(
            id_, name, ph_type, orient, sz, idx
        )
        self.insert_element_before(sp, 'p:extLst')
        return sp

    def add_table(self, id_, name, rows, cols, x, y, cx, cy):
        """
        Append a ``<p:graphicFrame>`` shape containing a table as specified
        in call.
        """
        graphicFrame = CT_GraphicalObjectFrame.new_table_graphicFrame(
            id_, name, rows, cols, x, y, cx, cy
        )
        self.insert_element_before(graphicFrame, 'p:extLst')
        return graphicFrame

    def add_textbox(self, id_, name, x, y, cx, cy):
        """
        Append a newly-created textbox ``<p:sp>`` shape having the specified
        position and size.
        """
        sp = CT_Shape.new_textbox_sp(id_, name, x, y, cx, cy)
        self.insert_element_before(sp, 'p:extLst')
        return sp

    def get_or_add_xfrm(self):
        """
        Return the ``<a:xfrm>`` grandchild element, newly-added if not
        present.
        """
        return self.grpSpPr.get_or_add_xfrm()

    def iter_ph_elms(self):
        """
        Generate each placeholder shape child element in document order.
        """
        for e in self.iter_shape_elms():
            if e.has_ph_elm:
                yield e

    def iter_shape_elms(self):
        """
        Generate each child of this ``<p:spTree>`` element that corresponds
        to a shape, in the sequence they appear in the XML.
        """
        for elm in self.iterchildren():
            if elm.tag in self._shape_tags:
                yield elm

    def update_infos(self, ids_, names, element, grouped, x=None, y=None, cx=None, cy=None):
        """
        Update informations (including *id_*, *name*, position and sizes) from given (grouped) shape element.
        """
        if grouped:
            prop_ = element.xpath('./p:nvGrpSpPr/p:cNvPr')[0]
            prop_.id = ids_[0]
            prop_.name = names[0]

            pos_ = element.xpath('./p:grpSpPr/a:xfrm/a:off')[0]
            if (x is not None) and (y is not None):
                (pos_.x, pos_.y) = x, y
            else:
                (pos_.x, pos_.y) = pos_.x+500, pos_.y+500

            if (cx is not None) and (cy is not None):
                pos_ = element.xpath('./p:grpSpPr/a:xfrm/a:ext')[0]
                (pos_.cx, pos_.cy) = cx, cy

            for _, e in enumerate(element.xpath('.//p:nvSpPr/p:cNvPr[@id]')):
                e.id = ids_[_+1]
                e.name = names[_+1]

        else: # ungrouped
            prop_ = element.xpath('./p:nvSpPr/p:cNvPr')[0]
            (prop_.id, prop_.name) = ids_[0], names[0]

            pos_ = element.xpath('./p:spPr/a:xfrm/a:off')[0]
            if (x is not None) and (y is not None):
                (pos_.x, pos_.y) = x, y
            else:
                (pos_.x, pos_.y) = pos_.x+500, pos_.y+500

            if (cx is not None) and (cy is not None):
                pos_ = element.xpath('./p:spPr/a:xfrm/a:ext')[0]
                (pos_.cx, pos_.cy) = cx, cy

        return element

    @property
    def xfrm(self):
        """
        The ``<a:xfrm>`` grandchild element or |None| if not found
        """
        return self.grpSpPr.xfrm


class CT_GroupShapeNonVisual(BaseShapeElement):
    """
    ``<p:nvGrpSpPr>`` element.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')


class CT_GroupShapeProperties(BaseOxmlElement):
    """
    The ``<p:grpSpPr>`` element
    """
    xfrm = ZeroOrOne('a:xfrm', successors=(
        'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill',
        'a:grpFill', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:extLst'
    ))
