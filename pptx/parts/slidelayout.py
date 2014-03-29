# encoding: utf-8

"""
Slide layout-related objects.
"""

from __future__ import absolute_import

from warnings import warn

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.parts.slide import BaseSlide
from pptx.shapes.placeholder import BasePlaceholder, BasePlaceholders
from pptx.shapes.shapetree import BaseShapeFactory, BaseShapeTree
from pptx.spec import PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_SLDNUM
from pptx.util import lazyproperty


class SlideLayout(BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    def iter_cloneable_placeholders(self):
        """
        Generate a reference to each layout placeholder on this slide layout
        that should be cloned to a slide when the layout is applied to the
        slide.
        """
        latent_ph_types = (PH_TYPE_DT, PH_TYPE_SLDNUM, PH_TYPE_FTR)
        for ph in self.placeholders:
            if ph.ph_type not in latent_ph_types:
                yield ph

    @lazyproperty
    def placeholders(self):
        """
        Instance of |_LayoutPlaceholders| containing sequence of placeholder
        shapes in this slide layout, sorted in *idx* order.
        """
        return _LayoutPlaceholders(self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |_LayoutShapeTree| containing sequence of shapes
        appearing on this slide layout.
        """
        return _LayoutShapeTree(self)

    @property
    def slide_master(self):
        """
        Slide master from which this slide layout inherits properties.
        """
        return self.part_related_by(RT.SLIDE_MASTER)

    @property
    def slidemaster(self):
        """
        Deprecated. Use ``.slide_master`` property instead.
        """
        msg = (
            'SlideLayout.slidemaster property is deprecated. Use .slide_mast'
            'er instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_master


class _LayoutShapeTree(BaseShapeTree):
    """
    Sequence of shapes appearing on a slide layout. The first shape in the
    sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), index(), and iteration.
    """
    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        parent = self
        return _LayoutShapeFactory(shape_elm, parent)


def _LayoutShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide layout.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return _LayoutPlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class _LayoutPlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide layout, providing differentiated behavior
    for slide layout placeholders, in particular, inheriting shape properties
    from the master placeholder having the same type.
    """
    @property
    def height(self):
        """
        The effective height of this placeholder shape; its directly-applied
        height if it has one, otherwise the height of its parent master
        placeholder.
        """
        return self._direct_or_inherited_value('height')

    @property
    def left(self):
        """
        The effective left of this placeholder shape; its directly-applied
        left if it has one, otherwise the left of its parent master
        placeholder.
        """
        return self._direct_or_inherited_value('left')

    @property
    def top(self):
        """
        The effective top of this placeholder shape; its directly-applied
        top if it has one, otherwise the top of its parent master
        placeholder.
        """
        return self._direct_or_inherited_value('top')

    @property
    def width(self):
        """
        The effective width of this placeholder shape; its directly-applied
        width if it has one, otherwise the width of its parent master
        placeholder.
        """
        return self._direct_or_inherited_value('width')

    def _direct_or_inherited_value(self, attr_name):
        """
        The effective value of *attr_name* on this placeholder shape; its
        directly-applied value if it has one, otherwise the value on the
        master placeholder it inherits from.
        """
        directly_applied_value = getattr(
            super(_LayoutPlaceholder, self), attr_name
        )
        if directly_applied_value is not None:
            return directly_applied_value
        inherited_value = self._inherited_value(attr_name)
        return inherited_value

    def _inherited_value(self, attr_name):
        """
        The attribute value, e.g. 'width' of the parent master placeholder of
        this placeholder shape
        """
        master_placeholder = self._master_placeholder
        if master_placeholder is None:
            return None
        inherited_value = getattr(master_placeholder, attr_name)
        return inherited_value

    @property
    def _master_placeholder(self):
        """
        The master placeholder shape this layout placeholder inherits from.
        """
        inheritee_ph_type = {
            'body':     'body',
            'chart':    'body',
            'clipArt':  'body',
            'ctrTitle': 'title',
            'dgm':      'body',
            'dt':       'dt',
            'ftr':      'ftr',
            'media':    'body',
            'obj':      'body',
            'pic':      'body',
            'sldNum':   'sldNum',
            'subTitle': 'body',
            'tbl':      'body',
            'title':    'title',
        }[self.ph_type]
        slide_master = self._slide_master
        master_placeholder = slide_master.placeholders.get(
            inheritee_ph_type, None
        )
        return master_placeholder

    @property
    def _slide_master(self):
        """
        The slide master this placeholder inherits from.
        """
        slide_layout = self.part
        slide_master = slide_layout.slide_master
        return slide_master


class _LayoutPlaceholders(BasePlaceholders):
    """
    Sequence of _LayoutPlaceholder instances representing the placeholder
    shapes on a slide layout.
    """
    def get(self, idx, default=None):
        """
        Return the first placeholder shape with matching *idx* value, or
        *default* if not found.
        """
        for placeholder in self:
            if placeholder.idx == idx:
                return placeholder
        return default

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        parent = self
        return _LayoutShapeFactory(shape_elm, parent)
