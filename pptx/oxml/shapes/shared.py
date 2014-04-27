# encoding: utf-8

"""
Common shape-related oxml objects
"""

from __future__ import absolute_import

from ..dml.fill import EG_FillProperties
from ..dml.line import (
    EG_LineDashProperties, EG_LineFillProperties, EG_LineJoinProperties
)
from ..ns import _nsmap, qn
from ..shared import BaseOxmlElement, ChildTagnames, Element
from ...util import Emu


class BaseShapeElement(BaseOxmlElement):
    """
    Provides common behavior for shape element classes like CT_Shape,
    CT_Picture, etc.
    """
    def __getattr__(self, name):
        # common code for position and size attributes
        if name in ('x', 'y', 'cx', 'cy'):
            xfrm = self.xfrm
            if xfrm is None:
                return None
            return getattr(xfrm, name)
        else:
            return super(BaseShapeElement, self).__getattr__(name)

    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('x', 'y', 'cx', 'cy'):
            xfrm = self.get_or_add_xfrm()
            setattr(xfrm, name, value)
        else:
            super(BaseShapeElement, self).__setattr__(name, value)

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
        Integer value of placeholder idx attribute, None if shape has no
        ``<p:ph>`` descendant.
        """
        ph = self.ph
        if ph is None:
            return None
        return int(ph.get('idx', 0))

    @property
    def ph_orient(self):
        """
        Placeholder orientation, e.g. ST_Direction.VERT ('vert'), None if
        shape has no ``<p:ph>`` descendant.
        """
        ph = self.ph
        if ph is None:
            return None
        return ph.get('orient', ST_Direction.HORZ)

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
        return ph.get('type', ST_PlaceholderType.OBJ)

    @property
    def shape_id(self):
        """
        Integer id of this shape
        """
        return int(self._nvXxPr.cNvPr.get('id'))

    @property
    def shape_name(self):
        """
        Name of this shape
        """
        return self._nvXxPr.cNvPr.get('name')

    @property
    def txBody(self):
        """
        Child ``<p:txBody>`` element, None if not present
        """
        return self.find(qn('p:txBody'))

    @property
    def xfrm(self):
        """
        The ``<a:xfrm>`` grandchild element or |None| if not found. This
        version works for ``<p:sp>``, ``<p:cxnSp>``, and ``<p:pic>``
        elements, others will need to override.
        """
        return self.spPr.xfrm

    @property
    def _nvXxPr(self):
        """
        Non-visual shape properties element for this shape. Actual name
        depends on the shape type, e.g. ``<p:nvPicPr>`` for picture shape.
        """
        return self.xpath('./*[1]', namespaces=_nsmap)[0]


class Fillable(BaseOxmlElement):
    """
    Provides common behavior for property elements that can contain one of
    the fill properties elements like ``<a:solidFill>``. Subclassed by
    CT_ShapeProperties and CT_LineProperties, perhaps others in the future.
    """
    def get_or_change_to_noFill(self):
        """
        Return the <a:noFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        if self.noFill is not None:
            return self.noFill
        self.remove_fill_element()
        return self._add_noFill()

    def get_or_change_to_solidFill(self):
        """
        Return the <a:solidFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        if self.solidFill is not None:
            return self.solidFill
        self.remove_fill_element()
        return self._add_solidFill()

    @property
    def noFill(self):
        """
        The <a:noFill> child element, or None if not present.
        """
        return self.find(qn('a:noFill'))

    @property
    def solidFill(self):
        """
        The <a:solidFill> child element, or None if not present.
        """
        return self.find(qn('a:solidFill'))

    def _add_noFill(self):
        """
        Return a newly added <a:noFill> child element, assuming no other fill
        EG_FillProperties element is present.
        """
        noFill = Element('a:noFill')
        successor_tagnames = self.child_tagnames_after('a:noFill')
        self.insert_element_before(noFill, *successor_tagnames)
        return noFill

    def _add_solidFill(self):
        """
        Return a newly added <a:solidFill> child element.
        """
        solidFill = Element('a:solidFill')
        successor_tagnames = self.child_tagnames_after('a:solidFill')
        self.insert_element_before(solidFill, *successor_tagnames)
        return solidFill


class EG_EffectProperties(object):

    __member_names__ = ('a:effectLst', 'a:effectDag')


class EG_Geometry(object):

    __member_names__ = ('a:custGeom', 'a:prstGeom')


class CT_LineProperties(Fillable):
    """
    Custom element class for <a:ln> element
    """

    child_tagnames = ChildTagnames.from_nested_sequence(
        EG_LineFillProperties.__member_names__,
        EG_LineDashProperties.__member_names__,
        EG_LineJoinProperties.__member_names__,
        'a:headEnd', 'a:tailEnd', 'a:extLst',
    )

    @property
    def fill_element(self):
        """
        Return the child representing the EG_FillProperties element group
        member in this element, or |None| if no such child is present.
        """
        return self.first_child_found_in(
            *EG_LineFillProperties.__member_names__
        )

    def remove_fill_element(self):
        """
        Remove the fill child element, e.g ``<a:solidFill>`` if present.
        """
        self.remove_if_present(*EG_LineFillProperties.__member_names__)

    @property
    def w(self):
        """
        Integer value of optional ``w`` attribute, or |None| if not present
        """
        w_str = self.get('w')
        if w_str is None:
            return None
        return Emu(int(w_str))


class CT_Point2D(BaseOxmlElement):
    """
    Custom element class for <a:off> element.
    """
    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('x', 'y'):
            self.set(name, str(value))
        else:
            super(CT_Point2D, self).__setattr__(name, value)

    @property
    def x(self):
        """
        Integer value of required ``x`` attribute.
        """
        x_str = self.get('x')
        return int(x_str)

    @property
    def y(self):
        """
        Integer value of required ``y`` attribute.
        """
        y_str = self.get('y')
        return int(y_str)


class CT_PositiveSize2D(BaseOxmlElement):
    """
    Custom element class for <a:ext> element.
    """
    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('cx', 'cy'):
            self.set(name, str(value))
        else:
            super(CT_PositiveSize2D, self).__setattr__(name, value)

    @property
    def cx(self):
        """
        Integer value of required ``cx`` attribute.
        """
        cx_str = self.get('cx')
        return int(cx_str)

    @property
    def cy(self):
        """
        Integer value of required ``cy`` attribute.
        """
        cy_str = self.get('cy')
        return int(cy_str)


class CT_ShapeProperties(Fillable):
    """
    Custom element class for <p:spPr> element. Shared by ``<p:sp>``,
    ``<p:pic>``, and ``<p:cxnSp>`` elements as well as a few more obscure
    ones.
    """

    child_tagnames = ChildTagnames.from_nested_sequence(
        'a:xfrm',
        EG_Geometry.__member_names__,
        EG_FillProperties.__member_names__,
        'a:ln',
        EG_EffectProperties.__member_names__,
        'a:scene3d', 'a:sp3d', 'a:extLst',
    )

    @property
    def cx(self):
        """
        Shape width, or None if not present.
        """
        cx_str_lst = self.xpath('./a:xfrm/a:ext/@cx', namespaces=_nsmap)
        if not cx_str_lst:
            return None
        return int(cx_str_lst[0])

    @property
    def cy(self):
        """
        Shape height, or None if not present.
        """
        cy_str_lst = self.xpath('./a:xfrm/a:ext/@cy', namespaces=_nsmap)
        if not cy_str_lst:
            return None
        return int(cy_str_lst[0])

    @property
    def fill_element(self):
        """
        Return the child representing the EG_FillProperties element group
        member in this element, or |None| if no such child is present.
        """
        return self.first_child_found_in(
            *EG_FillProperties.__member_names__
        )

    def get_or_add_ln(self):
        """
        Return the <a:ln> child element, newly added if not present.
        """
        ln = self.ln
        if ln is None:
            ln = self._add_ln()
        return ln

    def get_or_add_xfrm(self):
        """
        Return the <a:xfrm> child element, newly added if not already
        present.
        """
        xfrm = self.xfrm
        if xfrm is None:
            xfrm = self._add_xfrm()
        return xfrm

    @property
    def ln(self):
        """
        The <a:ln> child element, or None if not present.
        """
        return self.find(qn('a:ln'))

    def remove_fill_element(self):
        """
        Remove the fill child element, e.g ``<a:solidFill>`` if present.
        """
        self.remove_if_present(*EG_FillProperties.__member_names__)

    @property
    def x(self):
        """
        The integer value of `./xfrm/off/@x` attribute, or None if not
        present.
        """
        x_str_lst = self.xpath('./a:xfrm/a:off/@x', namespaces=_nsmap)
        if not x_str_lst:
            return None
        return int(x_str_lst[0])

    @property
    def xfrm(self):
        """
        The <a:xfrm> child element, or None if not present.
        """
        return self.find(qn('a:xfrm'))

    @property
    def y(self):
        """
        The top of the shape, or None if not present.
        """
        y_str_lst = self.xpath('./a:xfrm/a:off/@y', namespaces=_nsmap)
        if not y_str_lst:
            return None
        return int(y_str_lst[0])

    def _add_ln(self):
        """
        Return a newly added <a:ln> child element. It is the caller's
        responsibility to ensure one is not already present.
        """
        ln = Element('a:ln')
        successor_tagnames = self.child_tagnames_after('a:ln')
        self.insert_element_before(ln, *successor_tagnames)
        return ln

    def _add_xfrm(self):
        """
        Return a newly added <a:xfrm> child element.
        """
        xfrm = Element('a:xfrm')
        successor_tagnames = self.child_tagnames_after('a:xfrm')
        self.insert_element_before(xfrm, *successor_tagnames)
        return xfrm


class CT_Transform2D(BaseOxmlElement):
    """
    Custom element class for <a:xfrm> element.
    """
    def __getattr__(self, name):
        # common code for position and size attributes
        if name in ('x', 'y'):
            off = self.off
            if off is None:
                return None
            return getattr(off, name)
        elif name in ('cx', 'cy'):
            ext = self.ext
            if ext is None:
                return None
            return getattr(ext, name)
        else:
            return super(CT_Transform2D, self).__getattr__(name)

    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('x', 'y'):
            off = self.get_or_add_off()
            setattr(off, name, value)
        elif name in ('cx', 'cy'):
            ext = self.get_or_add_ext()
            setattr(ext, name, value)
        else:
            super(CT_Transform2D, self).__setattr__(name, value)

    @property
    def ext(self):
        """
        The <a:ext> child element, or None if not present.
        """
        return self.find(qn('a:ext'))

    def get_or_add_ext(self):
        """
        Return the <a:ext> child element, newly added if not already
        present.
        """
        ext = self.ext
        if ext is None:
            ext = Element('a:ext')
            ext.set('cx', '0')
            ext.set('cy', '0')
            self.append(ext)
        return ext

    def get_or_add_off(self):
        """
        Return the <a:off> child element, newly added if not already
        present.
        """
        off = self.off
        if off is None:
            off = Element('a:off')
            off.set('x', '0')
            off.set('y', '0')
            self.insert(0, off)
        return off

    @property
    def off(self):
        """
        The <a:off> child element, or None if not present.
        """
        return self.find(qn('a:off'))


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


class ST_PlaceholderType(object):

    BODY = 'body'
    CHART = 'chart'
    CLIP_ART = 'clipArt'
    CTR_TITLE = 'ctrTitle'
    DGM = 'dgm'
    DT = 'dt'
    FTR = 'ftr'
    HDR = 'hdr'
    MEDIA = 'media'
    OBJ = 'obj'
    PIC = 'pic'
    SLD_IMG = 'sldImg'
    SLD_NUM = 'sldNum'
    SUB_TITLE = 'subTitle'
    TBL = 'tbl'
    TITLE = 'title'
