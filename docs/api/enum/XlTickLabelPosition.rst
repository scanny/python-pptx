.. _XlTickLabelPosition:

``XL_TICK_LABEL_POSITION``
==========================

Specifies the position of tick-mark labels on a chart axis.

Example::

    from pptx.enum.chart import XL_TICK_LABEL_POSITION

    category_axis = chart.category_axis
    category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW

----

HIGH
    Top or right side of the chart.

LOW
    Bottom or left side of the chart.

NEXT_TO_AXIS
    Next to axis (where axis is not at either side of the chart).

NONE
    No tick labels.
