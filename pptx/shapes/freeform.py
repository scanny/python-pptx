# encoding: utf-8

"""The shape tree, the structure that holds a slide's shapes."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from collections import Sequence

from pptx.util import lazyproperty


class FreeformBuilder(object):
    """Allows a freeform shape to be specified and created.

    The initial pen position is provided on construction. From there, drawing
    proceeds using successive calls to draw line segments. The freeform shape
    may be closed by calling the :meth:`close` method. A shape having the
    same start and end point is automatically closed on creation, although
    that behavior can be overridden.

    A shape may have more than one contour, in which case overlapping areas
    are "subtracted".
    """

    def __init__(self, shapes, start_x, start_y, scale):
        super(FreeformBuilder, self).__init__()
        self._shapes = shapes
        self._start_x = start_x
        self._start_y = start_y
        self._scale = scale

    @classmethod
    def new(cls, shapes, start_x, start_y, scale):
        """Return a new |FreeformBuilder| object."""
        return cls(shapes, int(round(start_x)), int(round(start_y)), scale)

    def add_line_segments(self, vertices):
        """Add a straight line segment to each point in *vertices*.

        *vertices* must be an iterable of (x, y) pairs (2-tuples). Each x and
        y value is rounded to the nearest integer before use.
        """
        drawing_operations = self._drawing_operations
        for x, y in vertices:
            drawing_operations.add_line_segment(x, y)

    def convert_to_shape(self, origin_left=0, origin_top=0):
        """Return a new freeform shape positioned at the specified offset.

        The origin of the local coordinate system is specified in slide
        coordinates (EMU), perhaps most conveniently by use of a |Length|
        object.

        Note that this method may be called more than once to add multiple
        shapes of the same geometry in different locations on the slide.
        """
        drawing_operations = self._drawing_operations
        # ---add the empty shape with position and size---
        sp = self._spTree.add_freeform_sp(
            origin_left + self._left,
            origin_top + self._top,
            self._width,
            self._height
        )
        # ---start path---
        custGeom = sp.custGeom
        pathLst = custGeom.pathLst
        path = pathLst.add_path(
            w=drawing_operations.dx,
            h=drawing_operations.dy
        )
        offset_x, offset_y = (
            drawing_operations.min_x,
            drawing_operations.min_y
        )
        path.add_moveTo(
            self._start_x - offset_x,
            self._start_y - offset_y
        )
        # ---add line segments---
        for drawing_operation in drawing_operations:
            drawing_operation.apply_operation_to(path, (offset_x, offset_y))
        print(sp.xml)
        # ---create and return proxy shape---
        return self._shapes._shape_factory(sp)

    @lazyproperty
    def _drawing_operations(self):
        """Return the |DrawingOperations| sequence for this freeform."""
        return _DrawingOperations.new(self._start_x, self._start_y)

    @property
    def _height(self):
        """Return the vertical dimension of this shape in slide coordinates.

        This value is based on the actual extents of the shape and does not
        include any positioning offset.
        """
        return int(round(self._drawing_operations.dy * self._scale))

    @property
    def _left(self):
        """Return the leftmost extent of this shape in slide coordinates.

        Note that this value does not include any positioning offset; it
        assumes the drawing (local) coordinate origin is at (0, 0) on the
        slide.
        """
        print(
            self._drawing_operations.min_x,
            self._scale,
            self._scale,
            int(round(self._drawing_operations.min_x * self._scale))
        )
        return int(round(self._drawing_operations.min_x * self._scale))

    @lazyproperty
    def _spTree(self):
        """Return the `p:spTree` element this freeform will be added to."""
        return self._shapes._spTree

    @property
    def _top(self):
        """Return the topmost extent of this shape in slide coordinates.

        Note that this value does not include any positioning offset; it
        assumes the drawing (local) coordinate origin is at (0, 0) on the
        slide.
        """
        return int(round(self._drawing_operations.min_y * self._scale))

    @property
    def _width(self):
        """Return the width of this shape in slide coordinates.

        This value is based on the actual extents of the shape and does not
        include any positioning offset.
        """
        return int(round(self._drawing_operations.dx * self._scale))


class _DrawingOperations(Sequence):
    """Sequence of drawing operation objects for a freeform shape."""

    def __init__(self, start_x, start_y):
        super(_DrawingOperations, self).__init__()
        self._start_x = start_x
        self._start_y = start_y

    def __getitem__(self, idx):
        return self._operations.__getitem__(idx)

    def __iter__(self):
        return self._operations.__iter__()

    def __len__(self):
        return self._operations.__len__()

    @classmethod
    def new(cls, start_x, start_y):
        """Return |_DrawingOperations| object with specified starting point.

        Both *start_x* and *start_y* are rounded to the nearest integer
        before use.
        """
        return cls(int(round(start_x)), int(round(start_y)))

    def add_line_segment(self, x, y):
        """Add a |LineSegment| operation to the sequence."""
        self._operations.append(_LineSegment.new(x, y))

    @property
    def dx(self):
        """Return integer representing width (delta-x) extent of this shape.

        The returned value is in local coordinates.
        """
        min_x = max_x = self._start_x
        for drawing_operation in self:
            min_x = min(min_x, drawing_operation.x)
            max_x = max(max_x, drawing_operation.x)
        return max_x - min_x

    @property
    def dy(self):
        """Return integer representing width (delta-y) extent of this shape.

        The returned value is in local coordinates.
        """
        min_y = max_y = self._start_y
        for drawing_operation in self:
            min_y = min(min_y, drawing_operation.y)
            max_y = max(max_y, drawing_operation.y)
        return max_y - min_y

    @property
    def max_x(self):
        """Return integer representing the rightmost extent of this shape.

        The returned value is in local coordinates. Note that the bounding
        box of this shape need not include the local origin.
        """
        max_x = self._start_x
        for drawing_operation in self:
            max_x = max(max_x, drawing_operation.x)
        return max_x

    @property
    def max_y(self):
        """Return integer representing the bottommost extent of this shape.

        The returned value is in local coordinates.
        """
        max_y = self._start_y
        for drawing_operation in self:
            max_y = max(max_y, drawing_operation.y)
        return max_y

    @property
    def min_x(self):
        """Return integer representing the leftmost extent of this shape.

        The returned value is in local coordinates.
        """
        min_x = self._start_x
        for drawing_operation in self:
            min_x = min(min_x, drawing_operation.x)
        return min_x

    @property
    def min_y(self):
        """Return integer representing the topmost extent of this shape.

        The returned value is in local coordinates.
        """
        min_y = self._start_y
        for drawing_operation in self:
            min_y = min(min_y, drawing_operation.y)
        return min_y

    @lazyproperty
    def _operations(self):
        """Return the list composed to hold drawing operation objects."""
        return []


class _LineSegment(object):
    """Specifies a straight line segment to the specified point."""

    def __init__(self, x, y):
        super(_LineSegment, self).__init__()
        self._x = x
        self._y = y

    @classmethod
    def new(cls, x, y):
        """Return a new _LineSegment object to point *(x, y)*.

        Both *x* and *y* are rounded to the nearest integer before use.
        """
        return cls(int(round(x)), int(round(y)))

    def apply_operation_to(self, path, offset):
        """Add `a:moveTo` element to *path* for this line segment."""
        offset_x, offset_y = offset
        return path.add_lnTo(
            self._x - offset_x,
            self._y - offset_y
        )

    @property
    def x(self):
        """Return the horizontal location of this segment's end point.

        The returned value is an integer in local coordinates.
        """
        return self._x

    @property
    def y(self):
        """Return the vertical location of this segment's end point.

        The returned value is an integer in local coordinates.
        """
        return self._y
