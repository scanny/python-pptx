.. _XlErrorBarInclude:

``XL_ERROR_BAR_INCLUDE``
========================

Specifies which error-bar parts to display — the positive side, the negative
side, or both.

Example::

    from pptx.enum.chart import XL_ERROR_BAR_INCLUDE

    series.error_bars.include = XL_ERROR_BAR_INCLUDE.BOTH

----

BOTH
    Show both positive and negative error bars.

MINUS_VALUES
    Show only negative-side error bars.

PLUS_VALUES
    Show only positive-side error bars.
