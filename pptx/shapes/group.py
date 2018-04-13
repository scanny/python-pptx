# encoding: utf-8

"""GroupShape and related objects."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from pptx.shapes.base import BaseShape


class GroupShape(BaseShape):
    """A shape that acts as a container for other shapes."""

    @property
    def click_action(self):
        """Unconditionally raises `TypeError`.

        A group shape cannot have a click action or hover action.
        """
        raise TypeError('a group shape cannot have a click action')
