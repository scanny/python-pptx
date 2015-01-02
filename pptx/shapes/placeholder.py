# encoding: utf-8

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

"""
Placeholder-related objects, specific to shapes having a `p:ph` element.
"""

from .autoshape import Shape
from .shapetree import BaseShapeTree


class BasePlaceholders(BaseShapeTree):
    """
    Base class for placeholder collections that differentiate behaviors for
    a master, layout, and slide.
    """
    @staticmethod
    def _is_member_elm(shape_elm):
        """
        True if *shape_elm* is a placeholder shape, False otherwise.
        """
        return shape_elm.has_ph_elm


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
