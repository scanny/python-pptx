# encoding: utf-8

"""GroupShape and related objects."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from pptx.shapes.base import BaseShape


class GroupShape(BaseShape):
    """A shape that acts as a container for other shapes."""
