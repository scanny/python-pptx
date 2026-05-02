.. _XlErrorBarType:

``XL_ERROR_BAR_TYPE``
=====================

Specifies how the magnitude of error bars is computed.

Example::

    from pptx.enum.chart import XL_ERROR_BAR_TYPE

    series.error_bars.type = XL_ERROR_BAR_TYPE.FIXED_VALUE

----

CUSTOM
    Values are provided by ``plus`` and ``minus`` references to a range in the
    chart's embedded worksheet — one magnitude per data point. Note that in
    python-pptx 1.x the ``CUSTOM`` type is writable as a tag but the library
    does not populate the per-point ranges; clients needing custom per-point
    error bars must edit the underlying XML directly.

FIXED_VALUE
    All error bars have the same fixed magnitude in the value units of the
    series (e.g. "± 1.5 units").

PERCENT
    Error bar magnitude is a percentage of each data point's value.

STDEV
    Error bar magnitude is a number of standard deviations of the series values
    (e.g. ``±2σ``).

STERROR
    Error bar magnitude is the standard error of the series values. The stored
    ``value`` is ignored by PowerPoint in this mode.
