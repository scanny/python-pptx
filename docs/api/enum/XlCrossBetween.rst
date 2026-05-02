.. _XlCrossBetween:

``XL_CROSS_BETWEEN``
====================

Specifies how a value axis crosses the category axis.

Maps the ``ST_CrossBetween`` simple type defined in ``dml-chart.xsd``, which
takes one of ``between`` (the default for most chart types) or ``midCat``
(crossing occurs at the midpoint of each category).

Example::

    from pptx.enum.chart import XL_CROSS_BETWEEN

    value_axis = chart.value_axis
    value_axis.cross_between = XL_CROSS_BETWEEN.BETWEEN

----

BETWEEN
    Value axis crosses the category axis between categories.

MIDPOINT
    Value axis crosses the category axis at the midpoint of each category.
