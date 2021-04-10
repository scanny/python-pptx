# encoding: utf-8

"""Graphic Frame shape and related objects.

A graphic frame is a common container for table, chart, smart art, and media
objects.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

import base64

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
from pptx.table import Table


class GraphicFrame(BaseShape):
    """Container shape for table, chart, smart art, and media objects.

    Corresponds to a ``<p:graphicFrame>`` element in the shape tree.
    """

    @property
    def chart(self):
        """
        The |Chart| object containing the chart in this graphic frame. Raises
        |ValueError| if this graphic frame does not contain a chart.
        """
        if not self.has_chart:
            raise ValueError("shape does not contain a chart")
        return self.chart_part.chart

    @property
    def chart_part(self):
        """
        The |ChartPart| object containing the chart in this graphic frame.
        """
        rId = self._element.chart_rId
        chart_part = self.part.related_parts[rId]
        return chart_part

    @property
    def has_chart(self):
        """
        |True| if this graphic frame contains a chart object. |False|
        otherwise. When |True|, the chart object can be accessed using the
        ``.chart`` property.
        """
        return self._element.has_chart

    @property
    def has_table(self):
        """
        |True| if this graphic frame contains a table object. |False|
        otherwise. When |True|, the table object can be accessed using the
        ``.table`` property.
        """
        return self._element.has_table

    @property
    def shadow(self):
        """Unconditionally raises |NotImplementedError|.

        Access to the shadow effect for graphic-frame objects is
        content-specific (i.e. different for charts, tables, etc.) and has
        not yet been implemented.
        """
        raise NotImplementedError("shadow property on GraphicFrame not yet supported")

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, e.g.
        ``MSO_SHAPE_TYPE.TABLE``.
        """
        if self.has_chart:
            return MSO_SHAPE_TYPE.CHART
        elif self.has_table:
            return MSO_SHAPE_TYPE.TABLE
        else:
            return None

    @property
    def table(self):
        """
        The |Table| object contained in this graphic frame. Raises
        |ValueError| if this graphic frame does not contain a table.
        """
        if not self.has_table:
            raise ValueError("shape does not contain a table")
        tbl = self._element.graphic.graphicData.tbl
        return Table(tbl, self)


# --- bytes of embedded-XLSX icon (EMF format) ---
XLSX_ICON_BYTES = base64.b64decode(
    "AQAAAGwAAAAFAAAAAAAAACwAAAAnAAAAAAAAAAAAAAAlAwAAjAYAACBFTUYAAAEADBsAAAkA"
    "AAABAAAAAAAAAAAAAAAAAAAAgAcAADgEAAA1AQAArgAAAAAAAAAAAAAAAAAAAAi3BACwpwIA"
    "EgAAAAwAAAABAAAAGAAAAAwAAAAAAAAAGQAAAAwAAAD///8AcgAAAKAZAAAFAAAAAAAAACwA"
    "AAAnAAAABQAAAAAAAAAoAAAAKAAAAACA/wEAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAA"
    "AAAAAP///wAAAAAAbAAAADQAAACgAAAAABkAACgAAAAoAAAAKAAAACgAAAAoAAAAAQAgAAMA"
    "AAAAGQAAAAAAAAAAAAAAAAAAAAAAAAAA/wAA/wAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AABaXF2ydHd5/3R3ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/"
    "dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3"
    "ef90d3n/WlxdsgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "dHd5//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6/3R3ef8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHR3"
    "ef/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v90d3n/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB0d3n/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "dHd5/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdHd5//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6/3R3"
    "ef8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHR3ef/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v90d3n/"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB0d3n/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/dHd5/wAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdHd5//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6/3R3ef8AAAAA"
    "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"
    "QXwQ/0F8EP99pVv/oL6I/6C+iP9llT3/QXwQ/0F8EP9BfBD/lLZ5/6C+iP+Utnn/QXwQ/0F8"
    "EP9BfBD/QXwQ/0F8EP/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/dHd5/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEF8EP9BfBD/QXwQ/0F8"
    "EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/"
    "QXwQ/0F8EP9BfBD/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6/3R3ef8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABBfBD/QXwQ/0F8EP9BfBD/"
    "QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8"
    "EP9BfBD/QXwQ//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v90d3n/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPXQP8EF8EP9BfBD/QXwQ/0F8"
    "EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/"
    "QXwQ/0yDHv/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/dHd5/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwXAzA9dA/wQXwQ/0F8EP9BfBD/"
    "QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0F8EP9BfBD/QXwQ/0yD"
    "Hv/X4s7/+vr6//r6+v/6+vr/+vr6/6aoqf90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/"
    "dHd5/3R3ef8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHR3"
    "ef/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v90d3n/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/yMnJ/3R3"
    "ef8rLC1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB0d3n/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/dHd5//r6+v/6+vr/+vr6//r6+v/6+vr/yMnJ/3R3ef8rLC1g"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdHd5//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6/3R3ef/6+vr/+vr6//r6+v/6+vr/yMnJ/3R3ef8rLC1gAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHR3ef/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6//r6+v90d3n/+vr6//r6+v/6+vr/yMnJ/3R3ef8rLC1gAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB0d3n/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/dHd5//r6+v/6+vr/yMnJ/3R3ef8rLC1gAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdHd5//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/+vr6/3R3ef/6+vr/yMnJ/3R3ef8rLC1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHR3ef/6+vr/+vr6//r6"
    "+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/"
    "+vr6//r6+v90d3n/yMnJ/3R3ef8rLC1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB0d3n/+vr6//r6+v/6+vr/"
    "+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6+v/6+vr/+vr6//r6"
    "+v/6+vr/dHd5/3R3ef8rLC1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWlxdsnR3ef90d3n/dHd5/3R3"
    "ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/dHd5/3R3ef90d3n/"
    "dHd5/3R3ef8rLC1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAABGAAAAHAAAABAAAABJAGMAbwBuAE8AbgBsAHkARgAAAJwAAACOAAAAQwA6AFwA"
    "VwBJAE4ARABPAFcAUwBcAEkAbgBzAHQAYQBsAGwAZQByAFwAewA5ADAAMQA2ADAAMAAwADAA"
    "LQAwADAAMABGAC0AMAAwADAAMAAtADAAMAAwADAALQAwADAAMAAwADAAMAAwAEYARgAxAEMA"
    "RQB9AFwAeABsAGkAYwBvAG4AcwAuAGUAeABlAAAARgAAABAAAAACAAAAMQAAAA4AAAAUAAAA"
    "AAAAABAAAAAUAAAA"
)
