# encoding: utf-8

"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)


class ShadowFormat(object):
    """Provides access to shadow effect on a shape."""

    def __init__(self, spPr):
        # ---spPr may also be a grpSpPr, but interface is the same---
        self._spPr = self._element = spPr
