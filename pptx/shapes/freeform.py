# encoding: utf-8

"""Objects related to construction of freeform shapes."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from collections import Sequence

from pptx.util import lazyproperty


class FreeformBuilder(Sequence):
    """Allows a freeform shape to be specified and created.

    The initial pen position is provided on construction. From there, drawing
    proceeds using successive calls to draw line segments. The freeform shape
    may be closed by calling the :meth:`close` method.

    A shape may have more than one contour, in which case overlapping areas
    are "subtracted". A contour is a sequence of line segments beginning with
    a "move-to" operation. A move-to operation is automatically inserted in
    each new freeform; additional move-to ops can be inserted with the
    `.move_to()` method.
    """

    def __init__(self, shapes, start_x, start_y, x_scale, y_scale):
        super(FreeformBuilder, self).__init__()
        self._shapes = shapes
        self._start_x = start_x
        self._start_y = start_y
        self._x_scale = x_scale
        self._y_scale = y_scale

    def __getitem__(self, idx):
        return self._drawing_operations.__getitem__(idx)

    def __iter__(self):
        return self._drawing_operations.__iter__()

    def __len__(self):
        return self._drawing_operations.__len__()

    @classmethod
    def new(cls, shapes, start_x, start_y, x_scale, y_scale):
        """Return a new |FreeformBuilder| object.

        The initial pen location is specified (in local coordinates) by
        (*start_x*, *start_y*).
        """
        return cls(
            shapes, int(round(start_x)), int(round(start_y)),
            x_scale, y_scale
        )

    @lazyproperty
    def _drawing_operations(self):
        """Return the sequence of drawing operation objects for freeform."""
        return []
