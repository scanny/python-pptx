"""Series-related objects."""

from __future__ import annotations

from collections import namedtuple
from collections.abc import Sequence

from pptx.chart.datalabel import DataLabels
from pptx.chart.marker import Marker
from pptx.chart.point import BubblePoints, CategoryPoints, XyPoints
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import (
    XL_ERROR_BAR_INCLUDE,
    XL_ERROR_BAR_TYPE,
    XL_TRENDLINE_TYPE,
)
from pptx.oxml.chart.series import CT_ErrBars, CT_Trendline
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_TrendlineOrder, ST_TrendlinePeriod
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import lazyproperty

SheetReference = namedtuple("SheetReference", ("sheet_name", "a1_range"))
"""Parsed form of a ``c:f`` formula: ``(sheet_name, a1_range)``.

For example the formula ``Sheet1!$B$2:$F$2`` parses as
``SheetReference(sheet_name="Sheet1", a1_range="$B$2:$F$2")``. Sheet names
that contain spaces or punctuation are quoted with single quotes in the
formula (``'My Sheet'!$A$1:$A$3``); the quotes are stripped from
``sheet_name``.

.. versionadded:: 2026.05.0
"""


def _parse_sheet_reference(formula):
    """Return a |SheetReference| parsed from `formula`, or |None|.

    `formula` is the text content of a ``c:f`` element — e.g.
    ``Sheet1!$B$2:$F$2`` or ``'My Sheet'!$A$1``. Returns |None| when
    `formula` is |None|, empty, or has no ``!`` separator.
    """
    if not formula:
        return None
    sep = formula.rfind("!")
    if sep == -1:
        return None
    sheet_name = formula[:sep]
    a1_range = formula[sep + 1 :]
    # -- strip surrounding single quotes from sheet_name (e.g. 'My Sheet') --
    if len(sheet_name) >= 2 and sheet_name.startswith("'") and sheet_name.endswith("'"):
        # -- Excel escapes an embedded apostrophe as '' inside a quoted name --
        sheet_name = sheet_name[1:-1].replace("''", "'")
    return SheetReference(sheet_name=sheet_name, a1_range=a1_range)


class ErrorBars(object):
    """Provides access to the properties of a set of error bars on a chart series.

    Error bars annotate each data point with a vertical (or horizontal) whisker that
    communicates variability — confidence interval, standard deviation, standard error,
    or a fixed or percentage magnitude.

    An |ErrorBars| instance wraps a single ``c:errBars`` element.

    .. versionadded:: 2026.05.0
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
    def end_cap(self) -> bool:
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
    def format(self) -> ChartFormat:
        """The |ChartFormat| object providing line/fill properties for the bars.

        .. versionadded:: 2026.05.0
        """
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
    def value(self) -> float | None:
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


class Trendline(object):
    """Proxy wrapping a ``c:trendline`` element on a chart series.

    A trendline overlays a fitted curve — linear regression, logarithmic,
    polynomial, power, exponential, or a moving-average — on the data points
    of a chart series. A series may carry any number of trendlines, each
    drawn as its own line on the plot.

    Instances are constructed by :meth:`_BaseSeries.add_trendline` and
    enumerated via :attr:`_BaseSeries.trendlines`. Assigning properties on
    this object updates the underlying XML in place; to delete the
    trendline, call :meth:`delete`.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, trendline):
        super(Trendline, self).__init__()
        self._element = trendline
        self._trendline = trendline

    @property
    def trendline_type(self):
        """Read/write |XL_TRENDLINE_TYPE| member.

        Specifies the regression type — one of :attr:`LINEAR`, :attr:`LOGARITHMIC`,
        :attr:`POLYNOMIAL`, :attr:`POWER`, :attr:`EXPONENTIAL`, or :attr:`MOVING_AVG`.
        Defaults to :attr:`XL_TRENDLINE_TYPE.LINEAR` when the underlying
        ``c:trendlineType`` attribute is omitted (per the schema default).
        """
        trendlineType = self._trendline.trendlineType
        return trendlineType.val

    @trendline_type.setter
    def trendline_type(self, value):
        self._trendline.trendlineType.val = value

    @property
    def order(self) -> int:
        """Read/write int in range 2..6 specifying the polynomial order.

        Only meaningful when :attr:`trendline_type` is
        :attr:`XL_TRENDLINE_TYPE.POLYNOMIAL`. Returns the schema default (2)
        when the ``c:order`` child is absent.
        """
        order = self._trendline.order
        if order is None:
            return 2
        return order.val

    @order.setter
    def order(self, value):
        ST_TrendlineOrder.validate(value)
        self._trendline.get_or_add_order().val = value

    @property
    def period(self) -> int:
        """Read/write int >= 2 specifying the moving-average window size.

        Only meaningful when :attr:`trendline_type` is
        :attr:`XL_TRENDLINE_TYPE.MOVING_AVG`. Returns the schema default (2)
        when the ``c:period`` child is absent.
        """
        period = self._trendline.period
        if period is None:
            return 2
        return period.val

    @period.setter
    def period(self, value):
        ST_TrendlinePeriod.validate(value)
        self._trendline.get_or_add_period().val = value

    @property
    def forward(self) -> float | None:
        """Read/write float specifying categories to extend past the data.

        How many category units the fitted curve extends *beyond* the last
        data point (extrapolation). |None| when no ``c:forward`` child is
        present (PowerPoint treats this as zero).
        """
        forward = self._trendline.forward
        if forward is None:
            return None
        return forward.val

    @forward.setter
    def forward(self, value):
        if value is None:
            self._trendline._remove_forward()
            return
        self._trendline.get_or_add_forward().val = float(value)

    @property
    def backward(self) -> float | None:
        """Read/write float specifying categories to extend before the data.

        How many category units the fitted curve extends *before* the first
        data point (back-extrapolation). |None| when no ``c:backward`` child
        is present.
        """
        backward = self._trendline.backward
        if backward is None:
            return None
        return backward.val

    @backward.setter
    def backward(self, value):
        if value is None:
            self._trendline._remove_backward()
            return
        self._trendline.get_or_add_backward().val = float(value)

    @property
    def intercept(self) -> float | None:
        """Read/write float forcing the curve through a specific y-intercept.

        |None| when no ``c:intercept`` child is present — PowerPoint then
        lets the regression solve for its own intercept.
        """
        intercept = self._trendline.intercept
        if intercept is None:
            return None
        return intercept.val

    @intercept.setter
    def intercept(self, value):
        if value is None:
            self._trendline._remove_intercept()
            return
        self._trendline.get_or_add_intercept().val = float(value)

    @property
    def display_equation(self) -> bool:
        """Read/write bool: draw the fitted equation on the chart.

        Maps ``c:dispEq``. Defaults to |False| when the element is absent.
        """
        dispEq = self._trendline.dispEq
        if dispEq is None:
            return False
        return bool(dispEq.val)

    @display_equation.setter
    def display_equation(self, value):
        if not value:
            self._trendline._remove_dispEq()
            return
        self._trendline.get_or_add_dispEq().val = True

    @property
    def display_r_squared(self) -> bool:
        """Read/write bool: draw the fit's R-squared value on the chart.

        Maps ``c:dispRSqr``. Defaults to |False| when the element is absent.
        """
        dispRSqr = self._trendline.dispRSqr
        if dispRSqr is None:
            return False
        return bool(dispRSqr.val)

    @display_r_squared.setter
    def display_r_squared(self, value):
        if not value:
            self._trendline._remove_dispRSqr()
            return
        self._trendline.get_or_add_dispRSqr().val = True

    @property
    def name(self) -> str:
        """Read/write string: the user-visible label for this trendline.

        Maps ``c:name``. Returns an empty string when the element is absent —
        PowerPoint then synthesizes a default label (e.g. "Linear (Series1)").
        """
        name = self._trendline.name
        if name is None:
            return ""
        return name.text or ""

    @name.setter
    def name(self, value):
        if value is None or value == "":
            self._trendline._remove_name()
            return
        name = self._trendline.get_or_add_name()
        name.text = str(value)

    @lazyproperty
    def format(self) -> ChartFormat:
        """The |ChartFormat| object providing line/fill properties for this trendline.

        .. versionadded:: 2026.05.0
        """
        return ChartFormat(self._trendline)

    def delete(self):
        """Remove this ``c:trendline`` from its parent series.

        .. versionadded:: 2026.05.0
        """
        parent = self._trendline.getparent()
        if parent is not None:
            parent.remove(self._trendline)


def _new_trendline(
    trendline_type,
    order,
    period,
    forward,
    backward,
    intercept,
    display_equation,
    display_r_squared,
):
    """Return a freshly-constructed ``c:trendline`` element from Python-API values.

    `trendline_type` is an |XL_TRENDLINE_TYPE| member.
    The other parameters mirror the keyword arguments of
    :meth:`_BaseSeries.add_trendline`.
    """
    if order is not None:
        ST_TrendlineOrder.validate(order)
    if period is not None:
        ST_TrendlinePeriod.validate(period)
    return CT_Trendline.new_trendline(
        trendline_type=trendline_type.xml_value,
        order=order,
        period=period,
        forward=forward,
        backward=backward,
        intercept=intercept,
        display_equation=bool(display_equation),
        display_r_squared=bool(display_r_squared),
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
    def error_bars(self) -> ErrorBars | None:
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
                "series.error_bars must be an ErrorBars instance or None, got %r" % type(value),
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

    def delete(self):
        """Remove this series from its parent chart.

        The underlying ``c:ser`` element is detached from its parent
        ``c:{x}Chart``. After deletion the |Series| instance is stale —
        further use raises or silently no-ops; fetch a fresh
        ``chart.series[i]`` reference as needed.

        Deletion is typically used to exclude embedded-workbook columns
        from a chart without touching the workbook itself (GitHub issue
        #1043) — analogous to PowerPoint's right-click "Select Data"
        dialog unchecking a series. The ``c:idx`` and ``c:order`` values
        on remaining sibling series are left unchanged; PowerPoint
        tolerates gaps in either sequence and re-normalizes them on its
        next save.

        .. versionadded:: 2026.05.0
        """
        parent = self._ser.getparent()
        if parent is not None:
            parent.remove(self._ser)

    @lazyproperty
    def format(self) -> ChartFormat:
        """
        The |ChartFormat| instance for this series, providing access to shape
        properties such as fill and line.
        """
        return ChartFormat(self._ser)

    @property
    def has_error_bars(self) -> bool:
        """|True| if this series has a ``c:errBars`` child element, |False| otherwise.

        Equivalent to ``series.error_bars is not None`` but more idiomatic when the
        caller only needs a boolean.
        """
        return self._ser.errBars is not None

    @property
    def index(self) -> int:
        """
        The zero-based integer index of this series as reported in its
        `c:ser/c:idx` element.
        """
        return self._element.idx.val

    @property
    def name(self) -> str:
        """
        The string label given to this series, appears as the title of the
        column for this series in the Excel worksheet. It also appears as the
        label for this series in the legend.
        """
        names = self._element.xpath("./c:tx//c:pt/c:v/text()")
        name = names[0] if names else ""
        return name

    @property
    def source_range(self) -> str | None:
        """Read/write formula text of ``c:val/c:numRef/c:f``, or |None|.

        Returns the formula — e.g. ``"Sheet1!$B$2:$F$2"`` — that points the
        series' numeric values into a range of the chart's embedded workbook.
        |None| when the series has no ``c:val/c:numRef/c:f`` child (typically
        a literal-value series that stores values inline under ``c:numLit``
        rather than referencing the workbook).

        Assigning a string replaces the existing ``c:f`` text. A ``c:f`` child
        is created under ``c:val/c:numRef`` if one is not already present,
        but the series must already have a ``c:val/c:numRef`` container —
        attempting to set ``source_range`` on a series with no ``c:val`` or
        with a literal-value ``c:val/c:numLit`` raises :class:`ValueError`.
        Assigning |None| is not supported (raises :class:`ValueError`);
        surface a ``c:f``-less series by removing the ``c:numRef`` directly
        through the element tree.

        The cached values under ``c:val/c:numRef/c:numCache`` are **not**
        updated by this setter — they still reflect the prior range. Call
        :meth:`Chart.update_cached_values` afterward (which re-reads the
        embedded xlsx) or use
        :meth:`Chart.replace_data_preserve_formulas` to reconcile the cache
        with the new range's values.

        .. versionadded:: 2026.05.0
        """
        return self._f_text("./c:val/c:numRef/c:f")

    @source_range.setter
    def source_range(self, value):
        if value is None:
            raise ValueError(
                "assigning None to series.source_range is not supported; remove the "
                "c:numRef element directly through the XML tree if that is intended"
            )
        numRefs = self._element.xpath("./c:val/c:numRef")
        if not numRefs:
            raise ValueError(
                "series has no c:val/c:numRef element; source_range can only be set "
                "on a series that already references a range in the embedded workbook"
            )
        numRef = numRefs[0]
        fs = self._element.xpath("./c:val/c:numRef/c:f")
        if fs:
            f = fs[0]
        else:
            # -- c:f is the first child of c:numRef (before c:numCache/c:extLst) --
            f = OxmlElement("c:f")
            numRef.insert(0, f)
        f.text = str(value)

    @property
    def category_range(self) -> str | None:
        """Read-only formula text of ``c:cat/c:strRef/c:f`` or ``c:cat/c:numRef/c:f``.

        Returns the formula — e.g. ``"Sheet1!$A$2:$A$6"`` — that points the
        series' category labels into a range of the chart's embedded
        workbook. |None| when the series has no ``c:cat`` child, or has
        inline category data (``c:cat/c:strLit`` / ``c:cat/c:numLit``) rather
        than a workbook reference. Category data may be carried under either
        a ``c:strRef`` (text categories) or a ``c:numRef`` (numeric or date
        categories); this property returns whichever is present.

        .. versionadded:: 2026.05.0
        """
        return self._f_text("./c:cat/c:strRef/c:f") or self._f_text("./c:cat/c:numRef/c:f")

    @property
    def name_range(self) -> str | None:
        """Read-only formula text of ``c:tx/c:strRef/c:f``.

        Returns the formula — e.g. ``"Sheet1!$B$1"`` — that points the
        series name into a single cell of the chart's embedded workbook.
        |None| when the series has no ``c:tx/c:strRef/c:f`` child (a series
        whose name is stored inline as ``c:tx/c:v`` literal text or is
        omitted entirely returns |None| here; use :attr:`name` to read the
        cached literal string instead).

        .. versionadded:: 2026.05.0
        """
        return self._f_text("./c:tx/c:strRef/c:f")

    @property
    def values_sheet_reference(self) -> SheetReference | None:
        """|SheetReference| parsed from :attr:`source_range`, or |None|.

        Returns a ``(sheet_name, a1_range)`` :class:`~pptx.chart.series.SheetReference`
        namedtuple split out of the ``c:val/c:numRef/c:f`` formula. |None|
        when :attr:`source_range` is |None| or the formula has no ``!``
        sheet separator. Quoted sheet names (``'My Sheet'!$A$1``) are
        unquoted and any ``''`` apostrophe escapes are unescaped.

        .. versionadded:: 2026.05.0
        """
        return _parse_sheet_reference(self.source_range)

    def _f_text(self, xpath):
        """Return the text of the first ``c:f`` matched by `xpath`, or |None|.

        A helper for the ``*_range`` properties — returns |None| when no
        match is found and also when the match exists but is empty (``c:f``
        with no text content).
        """
        matches = self._element.xpath(xpath)
        if not matches:
            return None
        text = matches[0].text
        return text if text else None

    @property
    def trendlines(self) -> list[Trendline]:
        """List of |Trendline| objects currently attached to this series.

        Returns a fresh list on each access (mutations to the list itself are
        harmless — to remove a trendline use :meth:`Trendline.delete`). The
        list is empty when the series has no ``c:trendline`` children.
        """
        return [Trendline(tl) for tl in self._ser.trendline_lst]

    def add_trendline(
        self,
        trendline_type=XL_TRENDLINE_TYPE.LINEAR,
        order=None,
        period=None,
        forward=None,
        backward=None,
        intercept=None,
        display_equation=False,
        display_r_squared=False,
    ):
        """Attach a freshly-configured |Trendline| to this series and return it.

        `trendline_type` is an |XL_TRENDLINE_TYPE| member selecting the curve
        shape — LINEAR, LOGARITHMIC, POLYNOMIAL, POWER, EXPONENTIAL, or
        MOVING_AVG.

        `order` is the polynomial order (2..6). Required-for-effect when
        `trendline_type` is :attr:`POLYNOMIAL`; ignored (but accepted) by
        other types.

        `period` is the moving-average window size (>= 2). Required-for-effect
        when `trendline_type` is :attr:`MOVING_AVG`; ignored otherwise.

        `forward` / `backward` are category units to extend the fitted curve
        past / before the data points (extrapolation). Pass |None| (default)
        to omit the corresponding XML child, which PowerPoint treats as zero.

        `intercept` is a float forcing the curve through a specific y-value;
        |None| lets the regression solve for its own intercept.

        `display_equation` / `display_r_squared` control whether the fitted
        equation / R-squared are drawn on the chart as a label next to the
        trendline.
        """
        trendline = _new_trendline(
            trendline_type,
            order,
            period,
            forward,
            backward,
            intercept,
            display_equation,
            display_r_squared,
        )
        self._ser._insert_trendline(trendline)
        return Trendline(trendline)


class _BaseCategorySeries(_BaseSeries):
    """Base class for |BarSeries| and other category chart series classes."""

    @lazyproperty
    def data_labels(self):
        """|DataLabels| object controlling data labels for this series."""
        return DataLabels(
            self._ser.get_or_add_dLbls(),
            chart_type=self._chart_type_or_none,
        )

    @property
    def _chart_type_or_none(self):
        """The ``XL_CHART_TYPE`` of the plot owning this series, or ``None``.

        Used to scope :class:`DataLabels` position validation (issue #789).
        Resolution can fail on partial XML fixtures used in unit tests, in
        which case this returns ``None`` so validation is skipped rather
        than blowing up during attribute access.
        """
        # -- local import avoids a circular dependency at module import --
        from pptx.chart.plot import PlotFactory, PlotTypeInspector

        xChart = self._ser.getparent()
        if xChart is None:
            return None
        try:
            plot = PlotFactory(xChart, None)
            return PlotTypeInspector.chart_type(plot)
        except (KeyError, ValueError, IndexError, NotImplementedError):
            return None

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

    def iter_x_values(self):
        """Generate each X value in this series in on-chart order.

        A value of |None| represents a missing X value — either a blank
        cell in the Excel range or a missing ``c:pt`` in the inline
        ``c:numLit`` / cached ``c:numCache``. The generator reads whichever
        of ``c:xVal/c:numRef/c:numCache`` or ``c:xVal/c:numLit`` is present;
        XY/scatter and bubble series alike carry X values under ``c:xVal``.
        An empty sequence results when the series has no ``c:xVal`` child
        or its point cache is empty.

        .. versionadded:: 2026.05.0
        """
        xVal = self._element.xVal
        if xVal is None:
            return

        for idx in range(xVal.ptCount_val):
            yield xVal.pt_v(idx)

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

    @property
    def x_values(self):
        """Read-only tuple of float X values for this series, in on-chart order.

        Complements :attr:`values` (the Y values) — together they give the
        (x, y) coordinates of each data point on an XY/scatter or bubble
        chart. Reads whichever of ``c:xVal/c:numRef/c:numCache`` or
        ``c:xVal/c:numLit`` is present; a tuple element is |None| for any
        point where the X value is missing. Returns an empty tuple on a
        series that has no ``c:xVal`` child.

        .. versionadded:: 2026.05.0
        """
        return tuple(self.iter_x_values())


class BubbleSeries(XySeries):
    """
    A data point series belonging to a bubble plot.
    """

    def iter_bubble_sizes(self):
        """Generate each bubble-size value in this series in on-chart order.

        A value of |None| represents a missing bubble size — either a blank
        cell in the Excel range or a missing ``c:pt`` in the inline
        ``c:numLit`` / cached ``c:numCache``. Reads whichever of
        ``c:bubbleSize/c:numRef/c:numCache`` or ``c:bubbleSize/c:numLit``
        is present. An empty sequence results when the series has no
        ``c:bubbleSize`` child.

        .. versionadded:: 2026.05.0
        """
        bubbleSize = self._element.bubbleSize
        if bubbleSize is None:
            return

        for idx in range(bubbleSize.ptCount_val):
            yield bubbleSize.pt_v(idx)

    @lazyproperty
    def points(self):
        """
        The |BubblePoints| object providing access to individual data point
        objects used to discover and adjust the formatting and data labels of
        a data point.
        """
        return BubblePoints(self._ser)

    @property
    def bubble_sizes(self):
        """Read-only tuple of float bubble-size values in on-chart order.

        A third parallel sequence alongside :attr:`x_values` and
        :attr:`values` — together they give the (x, y, size) triple of each
        data point on a bubble chart. Reads whichever of
        ``c:bubbleSize/c:numRef/c:numCache`` or ``c:bubbleSize/c:numLit``
        is present; a tuple element is |None| for any point where the
        bubble size is missing. Returns an empty tuple when the series has
        no ``c:bubbleSize`` child.

        .. versionadded:: 2026.05.0
        """
        return tuple(self.iter_bubble_sizes())


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
