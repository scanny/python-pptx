# encoding: utf-8

"""
Common shape-related oxml objects
"""

from __future__ import absolute_import

from ...enum.shapes import PP_PLACEHOLDER
from ..ns import _nsmap, qn
from ..simpletypes import (
    ST_Coordinate, ST_DrawingElementId, ST_LineWidth, ST_PositiveCoordinate,
    XsdString, XsdUnsignedInt
)
from ...util import Emu
from ..xmlchemy import (
    BaseOxmlElement, Choice, OptionalAttribute, OxmlElement,
    RequiredAttribute, ZeroOrOne, ZeroOrOneChoice
)


class BaseShapeElement(BaseOxmlElement):
    """
    Provides common behavior for shape element classes like CT_Shape,
    CT_Picture, etc.
    """
    @property
    def cx(self):
        return self._get_xfrm_attr('cx')

    @cx.setter
    def cx(self, value):
        self._set_xfrm_attr('cx', value)

    @property
    def cy(self):
        return self._get_xfrm_attr('cy')

    @cy.setter
    def cy(self, value):
        self._set_xfrm_attr('cy', value)

    def get_or_add_xfrm(self):
        """
        Return the ``<a:xfrm>`` grandchild element, newly-added if not
        present. This version works for ``<p:sp>``, ``<p:cxnSp>``, and
        ``<p:pic>`` elements, others will need to override.
        """
        return self.spPr.get_or_add_xfrm()

    @property
    def has_ph_elm(self):
        """
        True if this shape element has a ``<p:ph>`` descendant, indicating it
        is a placeholder shape. False otherwise.
        """
        return self.ph is not None

    @property
    def ph(self):
        """
        The ``<p:ph>`` descendant element if there is one, None otherwise.
        """
        xpath = './*[1]/p:nvPr/p:ph'
        ph_elms = self.xpath(xpath, namespaces=_nsmap)
        if len(ph_elms) == 0:
            return None
        return ph_elms[0]

    @property
    def ph_idx(self):
        """
        Integer value of placeholder idx attribute. Raises |ValueError| if
        shape is not a placeholder.
        """
        ph = self.ph
        if ph is None:
            raise ValueError("not a placeholder shape")
        return ph.idx

    @property
    def ph_orient(self):
        """
        Placeholder orientation, e.g. 'vert'. Raises |ValueError| if shape is
        not a placeholder.
        """
        ph = self.ph
        if ph is None:
            raise ValueError("not a placeholder shape")
        return ph.orient

    @property
    def ph_sz(self):
        """
        Placeholder size, e.g. ST_PlaceholderSize.HALF, None if shape has no
        ``<p:ph>`` descendant.
        """
        ph = self.ph
        if ph is None:
            return None
        return ph.get('sz', ST_PlaceholderSize.FULL)

    @property
    def ph_type(self):
        """
        Placeholder type, e.g. ST_PlaceholderType.TITLE ('title'), none if
        shape has no ``<p:ph>`` descendant.
        """
        ph = self.ph
        if ph is None:
            return None
        return ph.type

    @property
    def shape_id(self):
        """
        Integer id of this shape
        """
        return self._nvXxPr.cNvPr.id

    @property
    def shape_name(self):
        """
        Name of this shape
        """
        return self._nvXxPr.cNvPr.name

    @property
    def txBody(self):
        """
        Child ``<p:txBody>`` element, None if not present
        """
        return self.find(qn('p:txBody'))

    @property
    def x(self):
        return self._get_xfrm_attr('x')

    @x.setter
    def x(self, value):
        self._set_xfrm_attr('x', value)

    @property
    def xfrm(self):
        """
        The ``<a:xfrm>`` grandchild element or |None| if not found. This
        version works for ``<p:sp>``, ``<p:cxnSp>``, and ``<p:pic>``
        elements, others will need to override.
        """
        return self.spPr.xfrm

    @property
    def y(self):
        return self._get_xfrm_attr('y')

    @y.setter
    def y(self, value):
        self._set_xfrm_attr('y', value)

    @property
    def _nvXxPr(self):
        """
        Required non-visual shape properties element for this shape. Actual
        name depends on the shape type, e.g. ``<p:nvPicPr>`` for picture
        shape.
        """
        return self.xpath('./*[1]', namespaces=_nsmap)[0]

    def _get_xfrm_attr(self, name):
        xfrm = self.xfrm
        if xfrm is None:
            return None
        return getattr(xfrm, name)

    def _set_xfrm_attr(self, name, value):
        xfrm = self.get_or_add_xfrm()
        setattr(xfrm, name, value)


class CT_ApplicationNonVisualDrawingProps(BaseOxmlElement):
    """
    ``<p:nvPr>`` element
    """
    ph = ZeroOrOne('p:ph', successors=(
        'a:audioCd', 'a:wavAudioFile', 'a:audioFile', 'a:videoFile',
        'a:quickTimeFile', 'p:custDataLst', 'p:extLst'
    ))


class CT_LineProperties(BaseOxmlElement):
    """
    Custom element class for <a:ln> element
    """
    eg_lineFillProperties = ZeroOrOneChoice(
        (Choice('a:noFill'), Choice('a:solidFill'), Choice('a:gradFill'),
         Choice('a:pattFill')),
        successors=(
            'a:prstDash', 'a:custDash', 'a:round', 'a:bevel', 'a:miter',
            'a:headEnd', 'a:tailEnd', 'a:extLst'
        )
    )
    w = OptionalAttribute('w', ST_LineWidth, default=Emu(0))

    @property
    def eg_fillProperties(self):
        """
        Required to fulfill the interface used by dml.fill.
        """
        return self.eg_lineFillProperties


class CT_NonVisualDrawingProps(BaseOxmlElement):
    """
    ``<p:cNvPr>`` custom element class.
    """
    id = RequiredAttribute('id', ST_DrawingElementId)
    name = RequiredAttribute('name', XsdString)


class CT_Placeholder(BaseOxmlElement):
    """
    ``<p:ph>`` custom element class.
    """
    @property
    def idx(self):
        idx_str = self.get('idx')
        if idx_str is None:
            return 0
        return int(idx_str)

    @idx.setter
    def idx(self, value):
        XsdUnsignedInt.validate(value)
        if value == 0 or value is None:
            if 'idx' in self.attrib:
                del self.attrib['idx']
            return
        str_value = str(value)
        self.set('idx', str_value)

    @property
    def orient(self):
        orient_str = self.get('orient')
        if orient_str is None:
            return ST_Direction.HORZ
        return orient_str

    @orient.setter
    def orient(self, value):
        if value not in (ST_Direction.HORZ, ST_Direction.VERT, None):
            raise ValueError("invalid setting for orient attribute")
        if value in (ST_Direction.HORZ, None):
            if 'orient' in self.attrib:
                del self.attrib['orient']
            return
        str_value = str(value)
        self.set('orient', str_value)

    @property
    def type(self):
        type_str = self.get('type')
        if type_str is None:
            return PP_PLACEHOLDER.OBJECT
        return PP_PLACEHOLDER.from_xml(type_str)

    @type.setter
    def type(self, value):
        if value is None or value is PP_PLACEHOLDER.OBJECT:
            if 'type' in self.attrib:
                del self.attrib['type']
            return
        str_value = PP_PLACEHOLDER.to_xml(value)
        self.set('type', str_value)


class CT_Point2D(BaseOxmlElement):
    """
    Custom element class for <a:off> element.
    """
    x = RequiredAttribute('x', ST_Coordinate)
    y = RequiredAttribute('y', ST_Coordinate)


class CT_PositiveSize2D(BaseOxmlElement):
    """
    Custom element class for <a:ext> element.
    """
    cx = RequiredAttribute('cx', ST_PositiveCoordinate)
    cy = RequiredAttribute('cy', ST_PositiveCoordinate)


class CT_ShapeProperties(BaseOxmlElement):
    """
    Custom element class for <p:spPr> element. Shared by ``<p:sp>``,
    ``<p:pic>``, and ``<p:cxnSp>`` elements as well as a few more obscure
    ones.
    """
    xfrm = ZeroOrOne('a:xfrm', successors=(
        'a:custGeom', 'a:prstGeom', 'a:ln', 'a:effectLst', 'a:effectDag',
        'a:scene3d', 'a:sp3d', 'a:extLst'
    ))
    eg_fillProperties = ZeroOrOneChoice(
        (Choice('a:noFill'), Choice('a:solidFill'), Choice('a:gradFill'),
         Choice('a:blipFill'), Choice('a:pattFill'), Choice('a:grpFill')),
        successors=(
            'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d',
            'a:extLst'
        )
    )
    ln = ZeroOrOne('a:ln', successors=(
        'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'
    ))

    @property
    def cx(self):
        """
        Shape width as an instance of Emu, or None if not present.
        """
        cx_str_lst = self.xpath('./a:xfrm/a:ext/@cx', namespaces=_nsmap)
        if not cx_str_lst:
            return None
        return Emu(cx_str_lst[0])

    @property
    def cy(self):
        """
        Shape height as an instance of Emu, or None if not present.
        """
        cy_str_lst = self.xpath('./a:xfrm/a:ext/@cy', namespaces=_nsmap)
        if not cy_str_lst:
            return None
        return Emu(cy_str_lst[0])

    @property
    def prstGeom(self):
        """
        The <a:prstGeom> child element, or None if not present.
        """
        return self.find(qn('a:prstGeom'))

    @property
    def x(self):
        """
        The offset of the left edge of the shape from the left edge of the
        slide, as an instance of Emu. Corresponds to the value of the
        `./xfrm/off/@x` attribute. None if not present.
        """
        x_str_lst = self.xpath('./a:xfrm/a:off/@x', namespaces=_nsmap)
        if not x_str_lst:
            return None
        return Emu(x_str_lst[0])

    @property
    def y(self):
        """
        The offset of the top of the shape from the top of the slide, as an
        instance of Emu. None if not present.
        """
        y_str_lst = self.xpath('./a:xfrm/a:off/@y', namespaces=_nsmap)
        if not y_str_lst:
            return None
        return Emu(y_str_lst[0])


class CT_Transform2D(BaseOxmlElement):
    """
    Custom element class for <a:xfrm> element.
    """
    off = ZeroOrOne('a:off', successors=('a:ext',))
    ext = ZeroOrOne('a:ext', successors=())

    @property
    def x(self):
        off = self.off
        if off is None:
            return None
        return off.x

    @x.setter
    def x(self, value):
        off = self.get_or_add_off()
        off.x = value

    @property
    def y(self):
        off = self.off
        if off is None:
            return None
        return off.y

    @y.setter
    def y(self, value):
        off = self.get_or_add_off()
        off.y = value

    @property
    def cx(self):
        ext = self.ext
        if ext is None:
            return None
        return ext.cx

    @cx.setter
    def cx(self, value):
        ext = self.get_or_add_ext()
        ext.cx = value

    @property
    def cy(self):
        ext = self.ext
        if ext is None:
            return None
        return ext.cy

    @cy.setter
    def cy(self, value):
        ext = self.get_or_add_ext()
        ext.cy = value

    def _new_ext(self):
        ext = OxmlElement('a:ext')
        ext.cx = 0
        ext.cy = 0
        return ext

    def _new_off(self):
        off = OxmlElement('a:off')
        off.x = 0
        off.y = 0
        return off


class ST_Direction(object):
    """
    Valid values for <p:ph orient=""> attribute
    """
    HORZ = 'horz'
    VERT = 'vert'


class ST_PlaceholderSize(object):
    """
    Valid values for <p:ph> sz (size) attribute
    """
    FULL = 'full'
    HALF = 'half'
    QUARTER = 'quarter'
