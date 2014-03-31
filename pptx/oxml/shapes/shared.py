# encoding: utf-8

"""
Common shape-related oxml objects
"""

from __future__ import absolute_import

from ..ns import _nsmap, qn
from ..shared import BaseOxmlElement, Element
from ...spec import PH_ORIENT_HORZ, PH_SZ_FULL, PH_TYPE_OBJ


class BaseShapeElement(BaseOxmlElement):
    """
    Provides common behavior for shape element classes like CT_Shape,
    CT_Picture, etc.
    """
    def __getattr__(self, name):
        # common code for position and size attributes
        if name in ('x', 'y'):
            xfrm = self.xfrm
            if xfrm is None:
                return None
            return getattr(xfrm, name)
        else:
            return super(BaseShapeElement, self).__getattr__(name)

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
        Placeholder orientation, e.g. PH_ORIENT_VERT ('vert'), None if shape
        has no ``<p:ph>`` descendant.
        """
        ph = self.ph
        if ph is None:
            return None
        return ph.get('orient', PH_ORIENT_HORZ)

    @property
    def ph_sz(self):
        """
        Placeholder size, e.g. PH_SZ_HALF, None if shape has no ``<p:ph>``
        descendant.
        """
        ph = self.ph
        if ph is None:
            return None
        return ph.get('sz', PH_SZ_FULL)

    @property
    def ph_type(self):
        """
        Placeholder type, e.g. PH_TYPE_TITLE ('title'), none if shape has no
        ``<p:ph>`` descendant.
        """
        ph = self.ph
        if ph is None:
            return None
        return ph.get('type', PH_TYPE_OBJ)

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
        elements, other will need to override.
        """
        return self.spPr.xfrm

    @property
    def _nvXxPr(self):
        """
        Non-visual shape properties element for this shape. Actual name
        depends on the shape type, e.g. ``<p:nvPicPr>`` for picture shape.
        """
        return self.xpath('./*[1]', namespaces=_nsmap)[0]


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
        else:
            return super(CT_Transform2D, self).__getattr__(name)

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
