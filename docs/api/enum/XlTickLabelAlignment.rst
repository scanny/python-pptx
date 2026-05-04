.. _XlTickLabelAlignment:

``XL_TICK_LABEL_ALIGNMENT``
===========================

Specifies the horizontal alignment of category-axis tick labels.

Maps the ``ST_LblAlgn`` simple type defined in ``dml-chart.xsd``, which
takes one of ``ctr`` (center), ``l`` (left), or ``r`` (right).

Example::

    from pptx.enum.chart import XL_TICK_LABEL_ALIGNMENT

    category_axis = chart.category_axis
    category_axis.label_align = XL_TICK_LABEL_ALIGNMENT.CENTER

----

CENTER
    Tick labels are center-aligned.

LEFT
    Tick labels are left-aligned.

RIGHT
    Tick labels are right-aligned.
