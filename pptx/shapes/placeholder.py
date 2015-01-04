# encoding: utf-8

"""
Placeholder-related objects, specific to shapes having a `p:ph` element.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .autoshape import Shape
from .base import BaseShape
from ..enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from .picture import Picture


class _BaseSlidePlaceholder(BaseShape):
    """
    Base class for placeholders on slides. Provides common behaviors such as
    inherited dimensions.
    """
    @property
    def shape_type(self):
        """
        Member of :ref:`MsoShapeType` specifying the type of this shape.
        Unconditionally ``MSO_SHAPE_TYPE.PLACEHOLDER`` in this case.
        Read-only.
        """
        return MSO_SHAPE_TYPE.PLACEHOLDER


class BasePlaceholder(Shape):
    """
    Base class for placeholder subclasses that differentiate the varying
    behaviors of placeholders on a master, layout, and slide.
    """
    @property
    def idx(self):
        """
        Integer placeholder 'idx' attribute, e.g. 0
        """
        return self._sp.ph_idx

    @property
    def orient(self):
        """
        Placeholder orientation, e.g. ST_Direction.HORZ
        """
        return self._sp.ph_orient

    @property
    def ph_type(self):
        """
        Placeholder type, e.g. PP_PLACEHOLDER.CENTER_TITLE
        """
        return self._sp.ph_type

    @property
    def sz(self):
        """
        Placeholder 'sz' attribute, e.g. ST_PlaceholderSize.FULL
        """
        return self._sp.ph_sz


class LayoutPlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide layout, providing differentiated behavior
    for slide layout placeholders, in particular, inheriting shape properties
    from the master placeholder having the same type, when a matching one
    exists.
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
            super(LayoutPlaceholder, self), attr_name
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
            PP_PLACEHOLDER.BODY:         PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.CHART:        PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.BITMAP:       PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.CENTER_TITLE: PP_PLACEHOLDER.TITLE,
            PP_PLACEHOLDER.ORG_CHART:    PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.DATE:         PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER:       PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.MEDIA_CLIP:   PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT:       PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.PICTURE:      PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.SLIDE_NUMBER: PP_PLACEHOLDER.SLIDE_NUMBER,
            PP_PLACEHOLDER.SUBTITLE:     PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.TABLE:        PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.TITLE:        PP_PLACEHOLDER.TITLE,
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


class MasterPlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide master.
    """


class SlidePlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide. Inherits shape properties from its
    corresponding slide layout placeholder.
    """
    @property
    def height(self):
        """
        The effective height of this placeholder shape; its directly-applied
        height if it has one, otherwise the height of its parent layout
        placeholder.
        """
        return self._effective_value('height')

    @property
    def left(self):
        """
        The effective left of this placeholder shape; its directly-applied
        left if it has one, otherwise the left of its parent layout
        placeholder.
        """
        return self._effective_value('left')

    @property
    def top(self):
        """
        The effective top of this placeholder shape; its directly-applied
        top if it has one, otherwise the top of its parent layout
        placeholder.
        """
        return self._effective_value('top')

    @property
    def width(self):
        """
        The effective width of this placeholder shape; its directly-applied
        width if it has one, otherwise the width of its parent layout
        placeholder.
        """
        return self._effective_value('width')

    def _effective_value(self, attr_name):
        """
        The effective value of *attr_name* on this placeholder shape; its
        directly-applied value if it has one, otherwise the value on the
        layout placeholder it inherits from.
        """
        directly_applied_value = getattr(
            super(SlidePlaceholder, self), attr_name
        )
        if directly_applied_value is not None:
            return directly_applied_value
        return self._inherited_value(attr_name)

    def _inherited_value(self, attr_name):
        """
        The attribute value, e.g. 'width' of the layout placeholder this
        slide placeholder inherits from
        """
        layout_placeholder = self._layout_placeholder
        if layout_placeholder is None:
            return None
        inherited_value = getattr(layout_placeholder, attr_name)
        return inherited_value

    @property
    def _layout_placeholder(self):
        """
        The layout placeholder shape this slide placeholder inherits from
        """
        layout = self._slide_layout
        layout_placeholder = layout.placeholders.get(idx=self.idx)
        return layout_placeholder

    @property
    def _slide_layout(self):
        """
        The slide layout from which the slide this placeholder belongs to
        inherits.
        """
        slide = self.part
        return slide.slide_layout


class PicturePlaceholder(_BaseSlidePlaceholder):
    """
    Placeholder shape that can only accept a picture.
    """


class PlaceholderPicture(Picture):
    """
    Placeholder shape populated with a picture.
    """
