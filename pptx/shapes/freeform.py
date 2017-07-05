# encoding: utf-8

"""The shape tree, the structure that holds a slide's shapes."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from collections import Sequence

from pptx.util import lazyproperty


class FreeformBuilder(Sequence):
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

    def __getitem__(self, idx):
        return self._drawing_operations.__getitem__(idx)

    def __iter__(self):
        return self._drawing_operations.__iter__()

    def __len__(self):
        return self._drawing_operations.__len__()

    @classmethod
    def new(cls, shapes, start_x, start_y, scale):
        """Return a new |FreeformBuilder| object."""
        return cls(shapes, int(round(start_x)), int(round(start_y)), scale)

    def add_line_segments(self, vertices):
        """Add a straight line segment to each point in *vertices*.

        *vertices* must be an iterable of (x, y) pairs (2-tuples). Each x and
        y value is rounded to the nearest integer before use.
        """
        for x, y in vertices:
            self._add_line_segment(x, y)

    def convert_to_shape(self, origin_left=0, origin_top=0):
        """Return a new freeform shape positioned at the specified offset.

        The origin of the local coordinate system is specified in slide
        coordinates (EMU), perhaps most conveniently by use of a |Length|
        object.

        Note that this method may be called more than once to add multiple
        shapes of the same geometry in different locations on the slide.
        """
        sp = self._add_freeform_sp(origin_left, origin_top)
        path = self._start_path(sp)
        for drawing_operation in self:
            drawing_operation.apply_operation_to(path)
        print(sp.xml)
        return self._shapes._shape_factory(sp)

    @property
    def shape_offset_x(self):
        """Return x distance of shape origin from local coordinate origin.

        The returned integer represents the leftmost extent of the freeform
        shape, in local coordinates. Note that the bounding box of the shape
        need not start at the local origin.
        """
        min_x = self._start_x
        for drawing_operation in self:
            min_x = min(min_x, drawing_operation.x)
        return min_x

    @property
    def shape_offset_y(self):
        """Return y distance of shape origin from local coordinate origin.

        The returned integer represents the topmost extent of the freeform
        shape, in local coordinates. Note that the bounding box of the shape
        need not start at the local origin.
        """
        min_y = self._start_y
        for drawing_operation in self:
            min_y = min(min_y, drawing_operation.y)
        return min_y

    def _add_freeform_sp(self, origin_left, origin_top):
        """Add a freeform `p:sp` element having no drawing elements.

        *origin_left* and *origin_top* are specified in slide coordinates,
        and represent the location of the local coordinates origin on the
        slide.
        """
        spTree = self._shapes._spTree
        return spTree.add_freeform_sp(
            origin_left + self._left,
            origin_top + self._top,
            self._width,
            self._height
        )

    def _add_line_segment(self, x, y):
        """Add a |LineSegment| operation to the sequence."""
        self._drawing_operations.append(_LineSegment.new(self, x, y))

    @lazyproperty
    def _drawing_operations(self):
        """Return the sequence of drawing operation objects for freeform."""
        return []

    @property
    def _dx(self):
        """Return integer width of this shape in local units."""
        min_x = max_x = self._start_x
        for drawing_operation in self:
            min_x = min(min_x, drawing_operation.x)
            max_x = max(max_x, drawing_operation.x)
        return max_x - min_x

    @property
    def _dy(self):
        """Return integer height of this shape in local units."""
        min_y = max_y = self._start_y
        for drawing_operation in self:
            min_y = min(min_y, drawing_operation.y)
            max_y = max(max_y, drawing_operation.y)
        return max_y - min_y

    @property
    def _height(self):
        """Return the vertical dimension of this shape in slide coordinates.

        This value is based on the actual extents of the shape and does not
        include any positioning offset.
        """
        return int(round(self._dy * self._scale))

    @property
    def _left(self):
        """Return the leftmost extent of this shape in slide coordinates.

        Note that this value does not include any positioning offset; it
        assumes the drawing (local) coordinate origin is at (0, 0) on the
        slide.
        """
        return int(round(self.shape_offset_x * self._scale))

    def _local_to_shape(self, local_x, local_y):
        """Translate local coordinates point to shape coordinates.

        Shape coordinates have the same unit as local coordinates, but are
        offset such that the origin of the shape coordinate system (0, 0) is
        located at the top-left corner of the shape bounding box.
        """
        return (
            local_x - self.shape_offset_x,
            local_y - self.shape_offset_y
        )

    def _start_path(self, sp):
        """Return a newly created `a:path` element added to *sp*.

        The returned `a:path` element has an `a:moveTo` element representing
        the shape starting point as its only child.
        """
        path = sp.add_path(w=self._dx, h=self._dy)
        path.add_moveTo(
            *self._local_to_shape(
                self._start_x, self._start_y
            )
        )
        return path

    @property
    def _top(self):
        """Return the topmost extent of this shape in slide coordinates.

        Note that this value does not include any positioning offset; it
        assumes the drawing (local) coordinate origin is located at slide
        coordinates (0, 0) (top-left corner of slide).
        """
        return int(round(self.shape_offset_y * self._scale))

    @property
    def _width(self):
        """Return the width of this shape in slide coordinates.

        This value is based on the actual extents of the shape and does not
        include any positioning offset.
        """
        return int(round(self._dx * self._scale))


class _LineSegment(object):
    """Specifies a straight line segment to the specified point."""

    def __init__(self, freeform_builder, x, y):
        super(_LineSegment, self).__init__()
        self._freeform_builder = freeform_builder
        self._x = x
        self._y = y

    @classmethod
    def new(cls, freeform_builder, x, y):
        """Return a new _LineSegment object to point *(x, y)*.

        Both *x* and *y* are rounded to the nearest integer before use.
        """
        return cls(freeform_builder, int(round(x)), int(round(y)))

    def apply_operation_to(self, path):
        """Add `a:moveTo` element to *path* for this line segment."""
        return path.add_lnTo(
            self._x - self._freeform_builder.shape_offset_x,
            self._y - self._freeform_builder.shape_offset_y
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
