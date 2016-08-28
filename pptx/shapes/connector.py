# encoding: utf-8

"""
Connector (line) shape and related objects. A connector is a line shape
having end-points that can be connected to other objects (but not to other
connectors). A line can be straight, have elbows, or can be curved.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .base import BaseShape


class Connector(BaseShape):
    """
    Connector (line) shape. A connector is a linear shape having end-points
    that can be connected to other objects (but not to other connectors).
    A line can be straight, have elbows, or can be curved.
    """
