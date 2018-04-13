# encoding: utf-8

"""GroupShape and related objects."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from pptx.shapes.base import BaseShape


class GroupShape(BaseShape):
    """A shape that acts as a container for other shapes."""

    @property
    def has_chart(self):
        raise NotImplementedError

    @property
    def has_table(self):
        raise NotImplementedError

    @property
    def has_text_frame(self):
        raise NotImplementedError

    @property
    def height(self):
        raise NotImplementedError

    @property
    def is_placeholder(self):
        raise NotImplementedError

    @property
    def left(self):
        raise NotImplementedError

    @property
    def name(self):
        raise NotImplementedError

    @property
    def part(self):
        raise NotImplementedError

    @property
    def rotation(self):
        raise NotImplementedError

    @property
    def shape_id(self):
        raise NotImplementedError

    @property
    def shape_type(self):
        raise NotImplementedError

    @property
    def top(self):
        raise NotImplementedError

    @property
    def width(self):
        raise NotImplementedError
