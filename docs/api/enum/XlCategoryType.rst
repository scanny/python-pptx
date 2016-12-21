.. _XlCategoryType:

``XL_CATEGORY_TYPE``
====================

Specifies the type of the category axis.

Example::

    from pptx.enum.chart import XL_CATEGORY_TYPE

    date_axis = chart.category_axis
    assert date_axis.category_type == XL_CATEGORY_TYPE.TIME_SCALE

----

AUTOMATIC_SCALE
    The application controls the axis type.

CATEGORY_SCALE
    Axis groups data by an arbitrary set of categories

TIME_SCALE
    Axis groups data on a time scale of days, months, or years.
