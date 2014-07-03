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
    @property
    def has_chart(self):
        """
        |True| if this graphic frame contains a chart object. |False|
        otherwise. When |True|, the chart object can be accessed using the
        ``.chart`` property.
        """
        return self._element.has_chart
