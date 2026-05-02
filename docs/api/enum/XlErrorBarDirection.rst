.. _XlErrorBarDirection:

``XL_ERROR_BAR_DIRECTION``
==========================

Specifies the direction of the error bars on a chart series.

Example::

    from pptx.enum.chart import XL_ERROR_BAR_DIRECTION

    series.error_bars.direction = XL_ERROR_BAR_DIRECTION.Y

----

X
    Error bars extend in the X direction (horizontal).

Y
    Error bars extend in the Y direction (vertical).
