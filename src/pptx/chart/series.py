"""Series-related objects."""

from __future__ import annotations

from collections.abc import Sequence

from pptx.chart.datalabel import DataLabels
from pptx.chart.marker import Marker
from pptx.chart.point import BubblePoints, CategoryPoints, XyPoints
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import (
    XL_ERROR_BAR_DIRECTION,
    XL_ERROR_BAR_INCLUDE,
    XL_ERROR_BAR_TYPE,
)
from pptx.oxml.chart.series import CT_ErrBars
from pptx.oxml.ns import qn
from pptx.util import lazyproperty


class ErrorBars(object):
    """Provides access to the properties of a set of error bars on a chart series.

    Error bars annotate each data point with a vertical (or horizontal) whisker that
    communicates variability — confidence interval, standard deviation, standard error,
    or a fixed or percentage magnitude.

    An |ErrorBars| instance wraps a single ``c:errBars`` element.
    """

    def __init__(self, errBars):
        super(ErrorBars, self).__init__()
        self._element = errBars
        self._errBars = errBars

    @property
    def direction(self):
        """Read/write |XL_ERROR_BAR_DIRECTION| member or |None|.

        Specifies the axis along which the error bars extend. |None| when the
        ``c:errDir`` child is absent — in which case PowerPoint infers Y for
        category-axis charts (bar, column, line, area) and X/Y per errBars element
        for XY/bubble charts.
        """
        errDir = self._errBars.errDir
        if errDir is None:
            return None
        return errDir.val

    @direction.setter
    def direction(self, value):
        if value is None:
            self._errBars._remove_errDir()
            return
        errDir = self._errBars.get_or_add_errDir()
        errDir.val = value

    @property
    def end_cap(self):
        """Read/write boolean.

        |True| if each error bar ends with a small perpendicular "T" cap (the
        PowerPoint default). |False| when the bar is drawn as a plain line.
        Reflects the *logical* meaning rather than the raw ``c:noEndCap`` boolean,
        which is inverted.
        """
        noEndCap = self._errBars.noEndCap
        if noEndCap is None:
            return True
        # -- noEndCap val=1 means "no end cap", so end_cap is False --
        return not bool(noEndCap.val)

    @end_cap.setter
    def end_cap(self, value):
        noEndCap = self._errBars.get_or_add_noEndCap()
        noEndCap.val = not bool(value)

    @lazyproperty
    def format(self):
        """The |ChartFormat| object providing line/fill properties for the bars."""
        return ChartFormat(self._errBars)

    @property
    def include(self):
        """Read/write |XL_ERROR_BAR_INCLUDE| member.

        Specifies whether bars are drawn on the plus side, the minus side, or both
        sides of the data point. Defaults to :attr:`XL_ERROR_BAR_INCLUDE.BOTH` when
        the ``c:errBarType`` attribute is omitted (per the schema default).
        """
        errBarType = self._errBars.errBarType
        if errBarType is None:
            return XL_ERROR_BAR_INCLUDE.BOTH
        return errBarType.val

    @include.setter
    def include(self, value):
        errBarType = self._errBars.get_or_add_errBarType()
        errBarType.val = value

    @property
    def type(self):
        """Read/write |XL_ERROR_BAR_TYPE| member specifying how magnitudes are computed.

        One of FIXED_VALUE, PERCENT, STDEV, STERROR, or CUSTOM. Defaults to
        :attr:`XL_ERROR_BAR_TYPE.FIXED_VALUE` when ``c:errValType`` is absent.
        """
        errValType = self._errBars.errValType
        if errValType is None:
            return XL_ERROR_BAR_TYPE.FIXED_VALUE
        return errValType.val

    @type.setter
    def type(self, value):
        errValType = self._errBars.get_or_add_errValType()
        errValType.val = value

    @property
    def value(self):
        """Read/write float specifying the fixed magnitude or percentage for the bars.

        Meaningful only when :attr:`type` is :attr:`XL_ERROR_BAR_TYPE.FIXED_VALUE`
        (an absolute value), :attr:`PERCENT` (a percentage expressed as e.g. ``5.0``
        for 5%), or :attr:`STDEV` (a multiplier for the series' standard deviation).
        PowerPoint ignores the stored value for STERROR but a harmless value is
        still written. Returns |None| when no ``c:val`` child is present.
        """
        val = self._errBars.val
        if val is None:
            return None
        raw = val.get("val")
        if raw is None:
            return None
        return float(raw)

    @value.setter
    def value(self, number):
        val = self._errBars.get_or_add_val()
        # -- c:val is registered to CT_NumDataSource (which has no `val` attribute --
        # -- accessor); write the attribute directly. --
        val.set("val", str(float(number)))


def _new_errBars(direction, include, type_, value):
    """Return a loose ``c:errBars`` element configured from Python-API values.

    `direction` is an |XL_ERROR_BAR_DIRECTION| member or |None| (omit `c:errDir`).
    `include` is an |XL_ERROR_BAR_INCLUDE| member.
    `type_` is an |XL_ERROR_BAR_TYPE| member.
    `value` is a float magnitude (fixed value, percentage, or stddev multiplier).
    """
    direction_str = None if direction is None else direction.xml_value
    include_str = include.xml_value
    type_str = type_.xml_value
    return CT_ErrBars.new_errBars(
        err_val_type=type_str,
        val=float(value),
        include=include_str,
        direction=direction_str,
    )


class _BaseSeries(object):
    """
    Base class for |BarSeries| and other series classes.
    """

    def __init__(self, ser):
        super(_BaseSeries, self).__init__()
        self._element = ser
        self._ser = ser

    @property
    def error_bars(self):
        """The |ErrorBars| object describing this series' error bars, or |None|.

        Returns |None| when the series has no ``c:errBars`` child. Assigning a
        freshly-constructed |ErrorBars| (e.g. one returned by
        :meth:`set_error_bars`) replaces any existing error bars. Assigning |None|
        removes the ``c:errBars`` element.

        Note this covers the most common single-``c:errBars`` case; XY/bubble
        series may legally have both an X and a Y ``c:errBars`` child, in which
        case only the first is surfaced here (setting Y-direction error bars is
        supported explicitly via :meth:`set_error_bars`).
        """
        errBars = self._ser.errBars
        if errBars is None:
            return None
        return ErrorBars(errBars)

    @error_bars.setter
    def error_bars(self, value):
        if value is None:
            self._ser._remove_errBars()
            return
        if not isinstance(value, ErrorBars):
            raise TypeError(
                "series.error_bars must be an ErrorBars instance or None, got %r" % type(value)
            )
        # -- replace any existing c:errBars with the new one --
        self._ser._remove_errBars()
        self._ser._insert_errBars(value._errBars)

    def set_error_bars(
        self,
        type_=XL_ERROR_BAR_TYPE.FIXED_VALUE,
        value=1.0,
        include=XL_ERROR_BAR_INCLUDE.BOTH,
        direction=None,
    ):
        """Attach a freshly configured |ErrorBars| block to this series and return it.

        `type_` is an |XL_ERROR_BAR_TYPE| member selecting how magnitude is computed —
        FIXED_VALUE, PERCENT, STDEV, STERROR, or CUSTOM (custom per-point values are
        not configured by this convenience method and would require client code to
        populate ``plus`` and ``minus`` children directly via the returned object's
        underlying element).

        `value` is the fixed magnitude (a number in the series' value units for
        FIXED_VALUE, a percentage such as ``5.0`` for PERCENT, or a standard-deviation
        multiplier for STDEV). Ignored by PowerPoint for STERROR.

        `include` is an |XL_ERROR_BAR_INCLUDE| member selecting which side of each
        data point is drawn (BOTH, PLUS_VALUES, or MINUS_VALUES).

        `direction` is an |XL_ERROR_BAR_DIRECTION| member or |None|. Category-axis
        charts (bar, column, line, area) accept |None| and PowerPoint infers Y.
        """
        errBars = _new_errBars(direction, include, type_, value)
        # -- replace any existing c:errBars --
        self._ser._remove_errBars()
        self._ser._insert_errBars(errBars)
        return ErrorBars(errBars)

    @lazyproperty
    def format(self):
        """
        The |ChartFormat| instance for this series, providing access to shape
        properties such as fill and line.
        """
        return ChartFormat(self._ser)

    @property
    def has_error_bars(self):
        """|True| if this series has a ``c:errBars`` child element, |False| otherwise.

        Equivalent to ``series.error_bars is not None`` but more idiomatic when the
        caller only needs a boolean.
        """
        return self._ser.errBars is not None

    @property
    def index(self):
        """
        The zero-based integer index of this series as reported in its
        `c:ser/c:idx` element.
        """
        return self._element.idx.val

    @property
    def name(self):
        """
        The string label given to this series, appears as the title of the
        column for this series in the Excel worksheet. It also appears as the
        label for this series in the legend.
        """
        names = self._element.xpath("./c:tx//c:pt/c:v/text()")
        name = names[0] if names else ""
        return name


class _BaseCategorySeries(_BaseSeries):
    """Base class for |BarSeries| and other category chart series classes."""

    @lazyproperty
    def data_labels(self):
        """|DataLabels| object controlling data labels for this series."""
        return DataLabels(self._ser.get_or_add_dLbls())

    @lazyproperty
    def points(self):
        """
        The |CategoryPoints| object providing access to individual data
        points in this series.
        """
        return CategoryPoints(self._ser)

    @property
    def values(self):
        """
        Read-only. A sequence containing the float values for this series, in
        the order they appear on the chart.
        """

        def iter_values():
            val = self._element.val
            if val is None:
                return
            for idx in range(val.ptCount_val):
                yield val.pt_v(idx)

        return tuple(iter_values())


class _MarkerMixin(object):
    """
    Mixin class providing `.marker` property for line-type chart series. The
    line-type charts are Line, XY, and Radar.
    """

    @lazyproperty
    def marker(self):
        """
        The |Marker| instance for this series, providing access to data point
        marker properties such as fill and line. Setting these properties
        determines the appearance of markers for all points in this series
        that are not overridden by settings at the point level.
        """
        return Marker(self._ser)


class AreaSeries(_BaseCategorySeries):
    """
    A data point series belonging to an area plot.
    """


class BarSeries(_BaseCategorySeries):
    """A data point series belonging to a bar plot."""

    @property
    def invert_if_negative(self):
        """
        |True| if a point having a value less than zero should appear with a
        fill different than those with a positive value. |False| if the fill
        should be the same regardless of the bar's value. When |True|, a bar
        with a solid fill appears with white fill; in a bar with gradient
        fill, the direction of the gradient is reversed, e.g. dark -> light
        instead of light -> dark. The term "invert" here should be understood
        to mean "invert the *direction* of the *fill gradient*".
        """
        invertIfNegative = self._element.invertIfNegative
        if invertIfNegative is None:
            return True
        return invertIfNegative.val

    @invert_if_negative.setter
    def invert_if_negative(self, value):
        invertIfNegative = self._element.get_or_add_invertIfNegative()
        invertIfNegative.val = value


class LineSeries(_BaseCategorySeries, _MarkerMixin):
    """
    A data point series belonging to a line plot.
    """

    @property
    def smooth(self):
        """
        Read/write boolean specifying whether to use curve smoothing to
        form the line connecting the data points in this series into
        a continuous curve. If |False|, a series of straight line segments
        are used to connect the points.
        """
        smooth = self._element.smooth
        if smooth is None:
            return True
        return smooth.val

    @smooth.setter
    def smooth(self, value):
        self._element.get_or_add_smooth().val = value


class PieSeries(_BaseCategorySeries):
    """
    A data point series belonging to a pie plot.
    """


class RadarSeries(_BaseCategorySeries, _MarkerMixin):
    """
    A data point series belonging to a radar plot.
    """


class XySeries(_BaseSeries, _MarkerMixin):
    """
    A data point series belonging to an XY (scatter) plot.
    """

    def iter_values(self):
        """
        Generate each float Y value in this series, in the order they appear
        on the chart. A value of `None` represents a missing Y value
        (corresponding to a blank Excel cell).
        """
        yVal = self._element.yVal
        if yVal is None:
            return

        for idx in range(yVal.ptCount_val):
            yield yVal.pt_v(idx)

    @lazyproperty
    def points(self):
        """
        The |XyPoints| object providing access to individual data points in
        this series.
        """
        return XyPoints(self._ser)

    @property
    def values(self):
        """
        Read-only. A sequence containing the float values for this series, in
        the order they appear on the chart.
        """
        return tuple(self.iter_values())


class BubbleSeries(XySeries):
    """
    A data point series belonging to a bubble plot.
    """

    @lazyproperty
    def points(self):
        """
        The |BubblePoints| object providing access to individual data point
        objects used to discover and adjust the formatting and data labels of
        a data point.
        """
        return BubblePoints(self._ser)


class SeriesCollection(Sequence):
    """
    A sequence of |Series| objects.
    """

    def __init__(self, parent_elm):
        # *parent_elm* can be either a c:plotArea or xChart element
        super(SeriesCollection, self).__init__()
        self._element = parent_elm

    def __getitem__(self, index):
        ser = self._element.sers[index]
        return _SeriesFactory(ser)

    def __len__(self):
        return len(self._element.sers)


def _SeriesFactory(ser):
    """
    Return an instance of the appropriate subclass of _BaseSeries based on the
    xChart element *ser* appears in.
    """
    xChart_tag = ser.getparent().tag

    try:
        SeriesCls = {
            qn("c:areaChart"): AreaSeries,
            qn("c:barChart"): BarSeries,
            qn("c:bubbleChart"): BubbleSeries,
            qn("c:doughnutChart"): PieSeries,
            qn("c:lineChart"): LineSeries,
            qn("c:pieChart"): PieSeries,
            qn("c:radarChart"): RadarSeries,
            qn("c:scatterChart"): XySeries,
        }[xChart_tag]
    except KeyError:
        raise NotImplementedError("series class for %s not yet implemented" % xChart_tag)

    return SeriesCls(ser)
