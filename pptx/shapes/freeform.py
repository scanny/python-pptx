# encoding: utf-8

"""Objects related to construction of freeform shapes."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from collections import Sequence


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

    @classmethod
    def new(cls, shapes, start_x, start_y, x_scale, y_scale):
        """Return a new |FreeformBuilder| object."""
        raise NotImplementedError
