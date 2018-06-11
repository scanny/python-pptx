# encoding: utf-8

"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)


class ShadowFormat(object):
    """Provides access to shadow effect on a shape."""

    def __init__(self, spPr):
        # ---spPr may also be a grpSpPr; both have a:effectLst child---
        self._element = spPr

    @property
    def inherit(self):
        """True if shape inherits shadow settings.

        An explicitly-defined shadow setting on a shape causes this property
        to return |False|. A shape with no explicitly-defined shadow setting
        inherits its shadow settings from the style hierarchy (and so returns
        |True|).
        """
        if self._element.effectLst is None:
            return True
        return False

    @inherit.setter
    def inherit(self, value):
        inherit = bool(value)
        if inherit:
            # ---remove any explicitly-defined effects
            self._element._remove_effectLst()
        else:
            # ---ensure at least the effectLst element is present
            self._element.get_or_add_effectLst()
