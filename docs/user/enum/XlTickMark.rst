.. _XlTickMark:

``XL_TICK_MARK``
================

Specifies a type of axis tick for a chart.

Example::

    from pptx.enum.chart import XL_TICK_MARK

    chart.value_axis.minor_tick_mark = XL_TICK_MARK.INSIDE

----

CROSS
    Tick mark crosses the axis

INSIDE
    Tick mark appears inside the axis

NONE
    No tick mark

OUTSIDE
    Tick mark appears outside the axis
