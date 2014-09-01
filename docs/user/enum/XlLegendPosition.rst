.. _XlLegendPosition:

``XL_LEGEND_POSITION``
======================

Specifies the position of the legend on a chart.

Example::

    from pptx.enum.chart import XL_LEGEND_POSITION

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

----

BOTTOM
    Below the chart.

CORNER
    In the upper-right corner of the chart border.

CUSTOM
    A custom position.

LEFT
    Left of the chart.

RIGHT
    Right of the chart.

TOP
    Above the chart.
