.. _XlTrendlineType:

``XL_TRENDLINE_TYPE``
=====================

Specifies the regression type fitted by a chart trendline.

Example::

    from pptx.enum.chart import XL_TRENDLINE_TYPE

    series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)

----

EXPONENTIAL
    Exponential trendline (``y = a * e^(b*x)``).

LINEAR
    Linear-regression trendline.

LOGARITHMIC
    Logarithmic trendline (``y = a * ln(x) + b``).

MOVING_AVG
    Moving-average trendline; the window size is given by the ``period``
    argument to :meth:`~pptx.chart.series._BaseSeries.add_trendline` (2 or
    greater).

POLYNOMIAL
    Polynomial-regression trendline; the order is given by the ``order``
    argument to :meth:`~pptx.chart.series._BaseSeries.add_trendline` (integer
    in the range 2..6 inclusive).

POWER
    Power trendline (``y = a * x^b``).
