.. _XlDataLabelPosition:

``XL_DATA_LABEL_POSITION``
==========================

Specifies where the data label is positioned.

Example::

    from pptx.enum.chart import XL_LABEL_POSITION

    data_labels = chart.plots[0].data_labels
    data_labels.position = XL_LABEL_POSITION.OUTSIDE_END

----

ABOVE
    The data label is positioned above the data point.

BELOW
    The data label is positioned below the data point.

BEST_FIT
    Word sets the position of the data label.

CENTER
    The data label is centered on the data point or inside a bar or a pie
    slice.

INSIDE_BASE
    The data label is positioned inside the data point at the bottom edge.

INSIDE_END
    The data label is positioned inside the data point at the top edge.

LEFT
    The data label is positioned to the left of the data point.

MIXED
    Data labels are in multiple positions.

OUTSIDE_END
    The data label is positioned outside the data point at the top edge.

RIGHT
    The data label is positioned to the right of the data point.
