.. _XlAxisCrosses:

``XL_AXIS_CROSSES``
===================

Specifies the point on the specified axis where the other axis crosses.

Example::

    from pptx.enum.chart import XL_AXIS_CROSSES

    value_axis.crosses = XL_AXIS_CROSSES.MAXIMUM

----

AUTOMATIC
    The axis crossing point is set automatically, often at zero.

CUSTOM
    The `.crosses_at` property specifies the axis crossing point.

MAXIMUM
    The axis crosses at the maximum value.

MINIMUM
    The axis crosses at the minimum value.
