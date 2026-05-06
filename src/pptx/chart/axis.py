"""Axis-related chart objects."""

from __future__ import annotations

from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import (
    XL_AXIS_CROSSES,
    XL_AXIS_POSITION,
    XL_CATEGORY_TYPE,
    XL_CROSS_BETWEEN,
    XL_TICK_LABEL_ALIGNMENT,
    XL_TICK_LABEL_POSITION,
    XL_TICK_MARK,
)
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_Orientation
from pptx.shared import ElementProxy
from pptx.text.text import Font, TextFrame
from pptx.util import lazyproperty


class _BaseAxis(object):
    """Base class for chart axis objects. All axis objects share these properties."""

    def __init__(self, xAx):
        super(_BaseAxis, self).__init__()
        self._element = xAx  # axis element, c:catAx or c:valAx
        self._xAx = xAx

    @property
    def axis_title(self) -> AxisTitle:
        """An |AxisTitle| object providing access to title properties.

        Calling this property is destructive in the sense that it adds an
        axis title element (`c:title`) to the axis XML if one is not already
        present. Use :attr:`has_title` to test for presence of axis title
        non-destructively.
        """
        return AxisTitle(self._element.get_or_add_title())

    @lazyproperty
    def format(self) -> ChartFormat:
        """
        The |ChartFormat| object providing access to the shape formatting
        properties of this axis, such as its line color and fill.
        """
        return ChartFormat(self._element)

    @property
    def has_major_gridlines(self) -> bool:
        """
        Read/write boolean value specifying whether this axis has gridlines
        at its major tick mark locations. Assigning |True| to this property
        causes major gridlines to be displayed. Assigning |False| causes them
        to be removed.
        """
        return self._element.majorGridlines is not None

    @has_major_gridlines.setter
    def has_major_gridlines(self, value):
        if bool(value) is True:
            self._element.get_or_add_majorGridlines()
        else:
            self._element._remove_majorGridlines()

    @property
    def has_minor_gridlines(self) -> bool:
        """
        Read/write boolean value specifying whether this axis has gridlines
        at its minor tick mark locations. Assigning |True| to this property
        causes minor gridlines to be displayed. Assigning |False| causes them
        to be removed.
        """
        return self._element.minorGridlines is not None

    @has_minor_gridlines.setter
    def has_minor_gridlines(self, value):
        if bool(value) is True:
            self._element.get_or_add_minorGridlines()
        else:
            self._element._remove_minorGridlines()

    @property
    def has_title(self) -> bool:
        """Read/write boolean specifying whether this axis has a title.

        |True| if this axis has a title, |False| otherwise. Assigning |True|
        causes an axis title to be added if not already present. Assigning
        |False| causes any existing title to be deleted.
        """
        return self._element.title is not None

    @has_title.setter
    def has_title(self, value):
        if bool(value) is True:
            self._element.get_or_add_title()
        else:
            self._element._remove_title()

    @lazyproperty
    def major_gridlines(self) -> MajorGridlines:
        """
        The |MajorGridlines| object representing the major gridlines for
        this axis.
        """
        return MajorGridlines(self._element)

    @lazyproperty
    def minor_gridlines(self) -> MinorGridlines:
        """|MinorGridlines| object providing access to minor-gridline formatting.

        Accessing this property is destructive in the sense that a `c:minorGridlines`
        element is added to the axis XML if one is not already present. Use
        :attr:`has_minor_gridlines` to test for presence non-destructively. The
        returned proxy exposes a ``.format`` property of type |ChartFormat| for
        setting the gridline color, width, dash style, etc.
        """
        return MinorGridlines(self._element)

    @property
    def major_tick_mark(self):
        """
        Read/write :ref:`XlTickMark` value specifying the type of major tick
        mark to display on this axis.
        """
        majorTickMark = self._element.majorTickMark
        if majorTickMark is None:
            return XL_TICK_MARK.CROSS
        return majorTickMark.val

    @major_tick_mark.setter
    def major_tick_mark(self, value):
        self._element._remove_majorTickMark()
        if value is XL_TICK_MARK.CROSS:
            return
        self._element._add_majorTickMark(val=value)

    @property
    def maximum_scale(self):
        """
        Read/write float value specifying the upper limit of the value range
        for this axis, the number at the top or right of the vertical or
        horizontal value scale, respectively. The value |None| indicates the
        upper limit should be determined automatically based on the range of
        data point values associated with the axis.
        """
        return self._element.scaling.maximum

    @maximum_scale.setter
    def maximum_scale(self, value):
        scaling = self._element.scaling
        scaling.maximum = value

    @property
    def minimum_scale(self):
        """
        Read/write float value specifying lower limit of value range, the
        number at the bottom or left of the value scale. |None| if no minimum
        scale has been set. The value |None| indicates the lower limit should
        be determined automatically based on the range of data point values
        associated with the axis.
        """
        return self._element.scaling.minimum

    @minimum_scale.setter
    def minimum_scale(self, value):
        scaling = self._element.scaling
        scaling.minimum = value

    @property
    def minor_tick_mark(self):
        """
        Read/write :ref:`XlTickMark` value specifying the type of minor tick
        mark for this axis.
        """
        minorTickMark = self._element.minorTickMark
        if minorTickMark is None:
            return XL_TICK_MARK.CROSS
        return minorTickMark.val

    @minor_tick_mark.setter
    def minor_tick_mark(self, value):
        self._element._remove_minorTickMark()
        if value is XL_TICK_MARK.CROSS:
            return
        self._element._add_minorTickMark(val=value)

    @property
    def position(self):
        """Read/write :ref:`XlAxisPosition` value specifying axis position on plot area.

        The value identifies on which side of the plot area the axis is drawn: `BOTTOM`,
        `LEFT`, `RIGHT`, or `TOP`. Returns |None| when no `c:axPos` element is present
        on the axis, which generally happens only on freshly-authored charts that have
        not yet been rendered by PowerPoint; PowerPoint inserts a `c:axPos` element on
        its first save, so in practice the value is rarely |None| on existing charts.

        Assigning |None| removes the `c:axPos` element. Assigning a member of
        |XL_AXIS_POSITION| sets the corresponding XML attribute value.
        """
        axPos = self._element.axPos
        if axPos is None:
            return None
        return axPos.val

    @position.setter
    def position(self, value):
        if value is None:
            self._element._remove_axPos()
            return
        XL_AXIS_POSITION.validate(value)
        axPos = self._element.get_or_add_axPos()
        axPos.val = value

    @property
    def reverse_order(self) -> bool:
        """Read/write bool value specifying whether to reverse plotting order for axis.

        For a category axis, this reverses the order in which the categories are
        displayed. This may be desired, for example, on a (horizontal) bar-chart where
        by default the first category appears at the bottom. Since we read from
        top-to-bottom, many viewers may find it most natural for the first category to
        appear on top.

        For a value axis, it reverses the direction of increasing value from
        bottom-to-top to top-to-bottom.
        """
        return self._element.orientation == ST_Orientation.MAX_MIN

    @reverse_order.setter
    def reverse_order(self, value):
        self._element.orientation = (
            ST_Orientation.MAX_MIN if bool(value) is True else ST_Orientation.MIN_MAX
        )

    @lazyproperty
    def tick_labels(self) -> TickLabels:
        """
        The |TickLabels| instance providing access to axis tick label
        formatting properties. Tick labels are the numbers appearing on
        a value axis or the category names appearing on a category axis.
        """
        return TickLabels(self._element)

    @property
    def tick_label_position(self):
        """
        Read/write :ref:`XlTickLabelPosition` value specifying where the tick
        labels for this axis should appear.
        """
        tickLblPos = self._element.tickLblPos
        if tickLblPos is None:
            return XL_TICK_LABEL_POSITION.NEXT_TO_AXIS
        if tickLblPos.val is None:
            return XL_TICK_LABEL_POSITION.NEXT_TO_AXIS
        return tickLblPos.val

    @tick_label_position.setter
    def tick_label_position(self, value):
        tickLblPos = self._element.get_or_add_tickLblPos()
        tickLblPos.val = value

    @property
    def visible(self) -> bool:
        """Read/write. |True| if axis is visible, |False| otherwise.

        The backing `c:delete` element has an OOXML-specified default of `val="true"`
        when the `val` attribute is omitted, so a `<c:delete/>` element without a
        `val` attribute is interpreted as `visible=False`. In practice PowerPoint does
        not follow the spec on this point (see `CT_Boolean_Explicit` for a similar
        case with `c:overlay`); python-pptx always writes the `val` attribute
        explicitly when assigning to `visible` so the result is interoperable. See
        issue #852.
        """
        delete = self._element.delete_
        if delete is None:
            return False
        # -- treat a missing `val` attribute as the OOXML default of `true` --
        return not delete.val

    @visible.setter
    def visible(self, value):
        if value not in (True, False):
            raise ValueError("assigned value must be True or False, got: %s" % value)
        delete = self._element.get_or_add_delete_()
        # -- Always write the `val` attribute explicitly. PowerPoint does not
        # -- honor the OOXML default of `true` for `c:delete` (see issue #852),
        # -- so relying on attribute omission would produce a file whose axis
        # -- visibility does not match the value the caller assigned.
        delete.set("val", "1" if value is False else "0")


class AxisTitle(ElementProxy):
    """Provides properties for manipulating axis title."""

    def __init__(self, title):
        super(AxisTitle, self).__init__(title)
        self._title = title

    @lazyproperty
    def format(self) -> ChartFormat:
        """|ChartFormat| object providing access to shape formatting.

        Return the |ChartFormat| object providing shape formatting properties
        for this axis title, such as its line color and fill.
        """
        return ChartFormat(self._element)

    @property
    def has_text_frame(self) -> bool:
        """Read/write Boolean specifying presence of a text frame.

        Return |True| if this axis title has a text frame, and |False|
        otherwise. Assigning |True| causes a text frame to be added if not
        already present. Assigning |False| causes any existing text frame to
        be removed along with any text contained in the text frame.
        """
        return self._title.tx_rich is not None

    @has_text_frame.setter
    def has_text_frame(self, value):
        if bool(value) is True:
            self._title.get_or_add_tx_rich()
        else:
            self._title._remove_tx()

    @property
    def text_frame(self) -> TextFrame:
        """|TextFrame| instance for this axis title.

        Return a |TextFrame| instance allowing read/write access to the text
        of this axis title and its text formatting properties. Accessing this
        property is destructive as it adds a new text frame if not already
        present.
        """
        rich = self._title.get_or_add_tx_rich()
        return TextFrame(rich, self)


class CategoryAxis(_BaseAxis):
    """A category axis of a chart."""

    @property
    def category_type(self):
        """
        A member of :ref:`XlCategoryType` specifying the scale type of this
        axis. Unconditionally ``CATEGORY_SCALE`` for a |CategoryAxis| object.
        """
        return XL_CATEGORY_TYPE.CATEGORY_SCALE

    @property
    def label_align(self) -> XL_TICK_LABEL_ALIGNMENT:
        """Horizontal alignment of category-axis tick labels.

        Read/write. Corresponds to ``c:lblAlgn/@val`` on this category axis,
        one of ``XL_TICK_LABEL_ALIGNMENT`` ``CENTER``, ``LEFT``, or ``RIGHT``.
        Returns ``CENTER`` (the PowerPoint default) when no ``c:lblAlgn``
        element is present. Assigning ``CENTER`` removes the backing element
        so the XML stays minimal. Assigning a non-member raises ``ValueError``.
        """
        lblAlgn = self._element.lblAlgn
        if lblAlgn is None:
            return XL_TICK_LABEL_ALIGNMENT.CENTER
        return lblAlgn.val

    @label_align.setter
    def label_align(self, value: XL_TICK_LABEL_ALIGNMENT):
        XL_TICK_LABEL_ALIGNMENT.validate(value)
        self._element._remove_lblAlgn()
        if value == XL_TICK_LABEL_ALIGNMENT.CENTER:
            return
        lblAlgn = self._element._add_lblAlgn()
        lblAlgn.val = value

    @property
    def tick_label_skip(self) -> int:
        """Integer count of categories between drawn tick-labels.

        Read/write. Corresponds to ``c:tickLblSkip/@val`` on this category
        axis. ``1`` (the default) draws a label for every category, ``2``
        draws every other category label, and so on. Returns ``1`` when no
        ``c:tickLblSkip`` element is present. Must be ``>= 1`` (ECMA-376
        ``ST_Skip``); a lesser value raises ``ValueError``.
        """
        tickLblSkip = self._element.tickLblSkip
        if tickLblSkip is None:
            return 1
        return tickLblSkip.val

    @tick_label_skip.setter
    def tick_label_skip(self, value: int):
        if not isinstance(value, int) or isinstance(value, bool) or value < 1:
            raise ValueError("tick_label_skip must be an int >= 1, got %r" % (value,))
        self._element._remove_tickLblSkip()
        if value == 1:
            return
        self._element._add_tickLblSkip(val=value)

    @property
    def tick_mark_skip(self) -> int:
        """Integer count of categories between drawn major tick marks.

        Read/write. Corresponds to ``c:tickMarkSkip/@val`` on this category
        axis. ``1`` (the default) draws a tick at every category, ``2`` draws
        a tick every other category, and so on. Returns ``1`` when no
        ``c:tickMarkSkip`` element is present. Must be ``>= 1`` (ECMA-376
        ``ST_Skip``); a lesser value raises ``ValueError``.
        """
        tickMarkSkip = self._element.tickMarkSkip
        if tickMarkSkip is None:
            return 1
        return tickMarkSkip.val

    @tick_mark_skip.setter
    def tick_mark_skip(self, value: int):
        if not isinstance(value, int) or isinstance(value, bool) or value < 1:
            raise ValueError("tick_mark_skip must be an int >= 1, got %r" % (value,))
        self._element._remove_tickMarkSkip()
        if value == 1:
            return
        self._element._add_tickMarkSkip(val=value)


class DateAxis(_BaseAxis):
    """A category axis with dates as its category labels.

    This axis-type has some special display behaviors such as making length of equal
    periods equal and normalizing month start dates despite unequal month lengths.
    """

    @property
    def category_type(self):
        """
        A member of :ref:`XlCategoryType` specifying the scale type of this
        axis. Unconditionally ``TIME_SCALE`` for a |DateAxis| object.
        """
        return XL_CATEGORY_TYPE.TIME_SCALE

    @property
    def major_unit(self):
        """The float number of units between major tick marks on this date axis.

        |None| corresponds to the 'Auto' setting in the UI, and indicates the
        value should be calculated by PowerPoint based on the date range of
        the chart data. The unit of measure is determined by the associated
        ``c:majorTimeUnit`` element (days by default).
        """
        majorUnit = self._element.majorUnit
        if majorUnit is None:
            return None
        return majorUnit.val

    @major_unit.setter
    def major_unit(self, value):
        self._element._remove_majorUnit()
        if value is None:
            return
        self._element._add_majorUnit(val=value)

    @property
    def minor_unit(self):
        """The float number of units between minor tick marks on this date axis.

        |None| corresponds to the 'Auto' setting in the UI, and indicates the
        value should be calculated by PowerPoint based on the date range of
        the chart data. The unit of measure is determined by the associated
        ``c:minorTimeUnit`` element (days by default).
        """
        minorUnit = self._element.minorUnit
        if minorUnit is None:
            return None
        return minorUnit.val

    @minor_unit.setter
    def minor_unit(self, value):
        self._element._remove_minorUnit()
        if value is None:
            return
        self._element._add_minorUnit(val=value)


class MajorGridlines(ElementProxy):
    """Provides access to the properties of the major gridlines appearing on an axis."""

    def __init__(self, xAx):
        super(MajorGridlines, self).__init__(xAx)
        self._xAx = xAx  # axis element, catAx or valAx

    @lazyproperty
    def format(self) -> ChartFormat:
        """
        The |ChartFormat| object providing access to the shape formatting
        properties of this data point, such as line and fill.
        """
        majorGridlines = self._xAx.get_or_add_majorGridlines()
        return ChartFormat(majorGridlines)


class MinorGridlines(ElementProxy):
    """Provides access to the properties of the minor gridlines appearing on an axis."""

    def __init__(self, xAx):
        super(MinorGridlines, self).__init__(xAx)
        self._xAx = xAx  # axis element, catAx, dateAx, or valAx

    @lazyproperty
    def format(self) -> ChartFormat:
        """|ChartFormat| object providing access to shape formatting.

        Returns the |ChartFormat| object providing line, fill, and shadow
        formatting for these minor gridlines. Accessing this property is
        destructive in that it adds a ``c:minorGridlines`` element to the
        axis XML if not already present.
        """
        minorGridlines = self._xAx.get_or_add_minorGridlines()
        return ChartFormat(minorGridlines)


class TickLabels(object):
    """A service class providing access to formatting of axis tick mark labels."""

    def __init__(self, xAx_elm):
        super(TickLabels, self).__init__()
        self._element = xAx_elm

    @lazyproperty
    def font(self) -> Font:
        """
        The |Font| object that provides access to the text properties for
        these tick labels, such as bold, italic, etc.
        """
        defRPr = self._element.defRPr
        font = Font(defRPr)
        return font

    @property
    def number_format(self):
        """
        Read/write string (e.g. "$#,##0.00") specifying the format for the
        numbers on this axis. The syntax for these strings is the same as it
        appears in the PowerPoint or Excel UI. Returns 'General' if no number
        format has been set. Note that this format string has no effect on
        rendered tick labels when :meth:`number_format_is_linked` is |True|.
        Assigning a format string to this property automatically sets
        :meth:`number_format_is_linked` to |False|.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return "General"
        return numFmt.formatCode

    @number_format.setter
    def number_format(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.formatCode = value
        self.number_format_is_linked = False

    @property
    def number_format_is_linked(self) -> bool:
        """
        Read/write boolean specifying whether number formatting should be
        taken from the source spreadsheet rather than the value of
        :meth:`number_format`.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return False
        souceLinked = numFmt.sourceLinked
        if souceLinked is None:
            return True
        return numFmt.sourceLinked

    @number_format_is_linked.setter
    def number_format_is_linked(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.sourceLinked = value

    @property
    def offset(self):
        """
        Read/write int value in range 0-1000 specifying the spacing between
        the tick mark labels and the axis as a percentange of the default
        value. 100 if no label offset setting is present.
        """
        lblOffset = self._element.lblOffset
        if lblOffset is None:
            return 100
        return lblOffset.val

    @offset.setter
    def offset(self, value):
        if self._element.tag != qn("c:catAx"):
            raise ValueError("only a category axis has an offset")
        self._element._remove_lblOffset()
        if value == 100:
            return
        lblOffset = self._element._add_lblOffset()
        lblOffset.val = value

    @property
    def rotation(self) -> float:
        """Clockwise rotation in degrees applied to these axis tick labels.

        Read/write. Corresponds to the ``rot`` attribute of the ``a:bodyPr``
        descendant of the axis ``c:txPr`` element. Returns ``0.0`` (the default)
        when no explicit rotation is set.

        Both ``int`` and ``float`` values are accepted; PowerPoint stores the
        value in 60000ths of a degree under the hood. Negative values are
        normalized to the equivalent positive rotation in the range
        ``[0, 360)``. See issue #329.
        """
        txPr = self._element.txPr
        if txPr is None:
            return 0.0
        return txPr.bodyPr.rot

    @rotation.setter
    def rotation(self, value: float):
        txPr = self._element.get_or_add_txPr()
        txPr.bodyPr.rot = value


class ValueAxis(_BaseAxis):
    """An axis having continuous (as opposed to discrete) values.

    The vertical axis is generally a value axis, however both axes of an XY-type chart
    are value axes.
    """

    @property
    def crosses(self):
        """
        Member of :ref:`XlAxisCrosses` enumeration specifying the point on
        this axis where the other axis crosses, such as auto/zero, minimum,
        or maximum. Returns `XL_AXIS_CROSSES.CUSTOM` when a specific numeric
        crossing point (e.g. 1.5) is defined.
        """
        crosses = self._cross_xAx.crosses
        if crosses is None:
            return XL_AXIS_CROSSES.CUSTOM
        return crosses.val

    @crosses.setter
    def crosses(self, value):
        cross_xAx = self._cross_xAx
        if value == XL_AXIS_CROSSES.CUSTOM and cross_xAx.crossesAt is not None:
            return
        cross_xAx._remove_crosses()
        cross_xAx._remove_crossesAt()
        if value == XL_AXIS_CROSSES.CUSTOM:
            cross_xAx._add_crossesAt(val=0.0)
        else:
            cross_xAx._add_crosses(val=value)

    @property
    def crosses_at(self):
        """
        Numeric value on this axis at which the perpendicular axis crosses.
        Returns |None| if no crossing value is set.
        """
        crossesAt = self._cross_xAx.crossesAt
        if crossesAt is None:
            return None
        return crossesAt.val

    @crosses_at.setter
    def crosses_at(self, value):
        cross_xAx = self._cross_xAx
        cross_xAx._remove_crosses()
        cross_xAx._remove_crossesAt()
        if value is None:
            return
        cross_xAx._add_crossesAt(val=value)

    @property
    def cross_between(self):
        """Read/write :ref:`XlCrossBetween` value specifying value-axis crossing mode.

        Indicates whether this value axis crosses the category axis `BETWEEN`
        category tick marks (the default for most chart types) or at the `MIDPOINT`
        of each category (mapped to the OOXML `midCat` enum value). Returns |None|
        when no `c:crossBetween` element is present on the axis.

        Assigning |None| removes the `c:crossBetween` element. Assigning a member
        of |XL_CROSS_BETWEEN| sets the corresponding XML attribute value.
        """
        crossBetween = self._element.crossBetween
        if crossBetween is None:
            return None
        return crossBetween.val

    @cross_between.setter
    def cross_between(self, value):
        if value is None:
            self._element._remove_crossBetween()
            return
        XL_CROSS_BETWEEN.validate(value)
        crossBetween = self._element.get_or_add_crossBetween()
        crossBetween.val = value

    @property
    def log_base(self) -> float | None:
        """Logarithm base for this value axis, or |None| for linear scaling.

        Read/write ``float | int | None``. Corresponds to
        ``c:valAx/c:scaling/c:logBase/@val``; presence of that element
        flips the axis from linear to logarithmic scaling, with the base
        given by ``@val`` (typically ``10.0`` or ``2.0``). Returns
        |None| when no ``c:logBase`` element is present (the linear
        default).

        Assigning |None| removes the element, reverting the axis to
        linear. Assigning a number outside ECMA-376 ``ST_LogBase`` — the
        inclusive range ``[2.0, 1000.0]`` — raises ``ValueError``.
        """
        return self._element.scaling.log_base

    @log_base.setter
    def log_base(self, value: float | int | None):
        scaling = self._element.scaling
        scaling.log_base = value

    @property
    def major_unit(self):
        """
        The float number of units between major tick marks on this value
        axis. |None| corresponds to the 'Auto' setting in the UI, and
        specifies the value should be calculated by PowerPoint based on the
        underlying chart data.
        """
        majorUnit = self._element.majorUnit
        if majorUnit is None:
            return None
        return majorUnit.val

    @major_unit.setter
    def major_unit(self, value):
        self._element._remove_majorUnit()
        if value is None:
            return
        self._element._add_majorUnit(val=value)

    @property
    def minor_unit(self):
        """
        The float number of units between minor tick marks on this value
        axis. |None| corresponds to the 'Auto' setting in the UI, and
        specifies the value should be calculated by PowerPoint based on the
        underlying chart data.
        """
        minorUnit = self._element.minorUnit
        if minorUnit is None:
            return None
        return minorUnit.val

    @minor_unit.setter
    def minor_unit(self, value):
        self._element._remove_minorUnit()
        if value is None:
            return
        self._element._add_minorUnit(val=value)

    @property
    def _cross_xAx(self):
        """
        The axis element in the same group (primary/secondary) that crosses
        this axis.
        """
        crossAx_id = self._element.crossAx.val
        expr = '(../c:catAx | ../c:valAx | ../c:dateAx)/c:axId[@val="%d"]' % crossAx_id
        cross_axId = self._element.xpath(expr)[0]
        return cross_axId.getparent()
