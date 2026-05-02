.. _XlAxisPosition:

``XL_AXIS_POSITION``
====================

Specifies the position of an axis on a chart.

Maps the ``ST_AxPos`` simple type defined in ``dml-chart.xsd``, which takes
one of ``b`` (bottom), ``l`` (left), ``r`` (right), or ``t`` (top).

Example::

    from pptx.enum.chart import XL_AXIS_POSITION

    category_axis = chart.category_axis
    category_axis.position = XL_AXIS_POSITION.BOTTOM

----

BOTTOM
    Axis is drawn at the bottom of the plot area.

LEFT
    Axis is drawn at the left of the plot area.

RIGHT
    Axis is drawn at the right of the plot area.

TOP
    Axis is drawn at the top of the plot area.
