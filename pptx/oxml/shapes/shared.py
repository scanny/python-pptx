# encoding: utf-8

"""
Common shape-related oxml objects
"""

from __future__ import absolute_import

from pptx.oxml.core import BaseOxmlElement
from pptx.oxml.ns import _nsmap, qn
from pptx.spec import PH_ORIENT_HORZ, PH_SZ_FULL, PH_TYPE_OBJ


class BaseShapeElement(BaseOxmlElement):
    """
    Provides common behavior for shape element classes like CT_Shape,
    CT_Picture, etc.
    """
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
    def txBody(self):
        """
        Child ``<p:txBody>`` element, None if not present
        """
        return self.find(qn('p:txBody'))

    @property
    def _nvXxPr(self):
        """
        Non-visual shape properties element for this shape. Actual name
        depends on the shape type, e.g. ``<p:nvPicPr>`` for picture shape.
        """
        return self.xpath('./*[1]', namespaces=_nsmap)[0]
