# encoding: utf-8

"""
Graphic Frame shape and related objects. A graphic frame is a common
container for table, chart, smart art, and media objects.
"""

from __future__ import absolute_import

from pptx.shapes.shape import BaseShape


class GraphicFrame(BaseShape):
    """
    Container shape for table, chart, smart art, and media objects.
    Corresponds to a ``<p:graphicFrame>`` element in the shape tree.
    """
