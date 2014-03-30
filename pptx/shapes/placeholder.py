# encoding: utf-8

"""
Placeholder object, a wrapper (decorator pattern) around an autoshape having
a ``ph`` element. Provides access to placeholder-specific properties of the
shape, such as the placeholder type. All other attribute gets are forwarded
to the underlying shape.
"""

from pptx.shapes.autoshape import Shape
from pptx.shapes.shapetree import BaseShapeTree


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
        Placeholder orientation, e.g. PH_ORIENT_HORZ
        """
        return self._sp.ph_orient

    @property
    def ph_type(self):
        """
        Placeholder type, e.g. PH_TYPE_CTRTITLE
        """
        return self._sp.ph_type

    @property
    def sz(self):
        """
        Placeholder 'sz' attribute, e.g. PH_SZ_FULL
        """
        return self._sp.ph_sz
