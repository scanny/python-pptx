"""Axis-related oxml objects."""

from __future__ import annotations

from pptx.enum.chart import (
    XL_AXIS_CROSSES,
    XL_AXIS_POSITION,
    XL_CROSS_BETWEEN,
    XL_TICK_LABEL_ALIGNMENT,
    XL_TICK_LABEL_POSITION,
    XL_TICK_MARK,
)
from pptx.oxml.chart.shared import CT_Title
from pptx.oxml.simpletypes import ST_AxisUnit, ST_LblOffset, ST_Orientation, ST_Skip
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class BaseAxisElement(BaseOxmlElement):
    """Base class for catAx, dateAx, valAx, and perhaps other axis elements."""

    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` great-great-grandchild element, added with its
        ancestors if not present.
        """
        txPr = self.get_or_add_txPr()
        defRPr = txPr.defRPr
        return defRPr

    @property
    def orientation(self):
        """Value of `val` attribute of `c:scaling/c:orientation` grandchild element.

        Defaults to `ST_Orientation.MIN_MAX` if attribute or any ancestors are not
        present.
        """
        orientation = self.scaling.orientation
        if orientation is None:
            return ST_Orientation.MIN_MAX
        return orientation.val

    @orientation.setter
    def orientation(self, value):
        """`value` is a member of `ST_Orientation`."""
        self.scaling._remove_orientation()
        if value == ST_Orientation.MAX_MIN:
            self.scaling.get_or_add_orientation().val = value

    def _new_title(self):
        return CT_Title.new_title()

    def _new_txPr(self):
        return CT_TextBody.new_txPr()


class CT_AxisUnit(BaseOxmlElement):
    """Used for `c:majorUnit` and `c:minorUnit` elements, and others."""

    val = RequiredAttribute("val", ST_AxisUnit)


class CT_AxPos(BaseOxmlElement):
    """`c:axPos` element, specifying the position of an axis on a chart.

    The `val` attribute takes one of `b` (bottom), `l` (left), `r` (right), or
    `t` (top) — the values defined by `ST_AxPos` in `dml-chart.xsd`.
    """

    val = RequiredAttribute("val", XL_AXIS_POSITION)


class CT_CatAx(BaseAxisElement):
    """`c:catAx` element, defining a category axis."""

    _tag_seq = (
        "c:axId",
        "c:scaling",
        "c:delete",
        "c:axPos",
        "c:majorGridlines",
        "c:minorGridlines",
        "c:title",
        "c:numFmt",
        "c:majorTickMark",
        "c:minorTickMark",
        "c:tickLblPos",
        "c:spPr",
        "c:txPr",
        "c:crossAx",
        "c:crosses",
        "c:crossesAt",
        "c:auto",
        "c:lblAlgn",
        "c:lblOffset",
        "c:tickLblSkip",
        "c:tickMarkSkip",
        "c:noMultiLvlLbl",
        "c:extLst",
    )
    scaling = OneAndOnlyOne("c:scaling")
    delete_ = ZeroOrOne("c:delete", successors=_tag_seq[3:])
    axPos = ZeroOrOne("c:axPos", successors=_tag_seq[4:])
    majorGridlines = ZeroOrOne("c:majorGridlines", successors=_tag_seq[5:])
    minorGridlines = ZeroOrOne("c:minorGridlines", successors=_tag_seq[6:])
    title = ZeroOrOne("c:title", successors=_tag_seq[7:])
    numFmt = ZeroOrOne("c:numFmt", successors=_tag_seq[8:])
    majorTickMark = ZeroOrOne("c:majorTickMark", successors=_tag_seq[9:])
    minorTickMark = ZeroOrOne("c:minorTickMark", successors=_tag_seq[10:])
    tickLblPos = ZeroOrOne("c:tickLblPos", successors=_tag_seq[11:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[12:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[13:])
    crosses = ZeroOrOne("c:crosses", successors=_tag_seq[15:])
    crossesAt = ZeroOrOne("c:crossesAt", successors=_tag_seq[16:])
    lblAlgn = ZeroOrOne("c:lblAlgn", successors=_tag_seq[18:])
    lblOffset = ZeroOrOne("c:lblOffset", successors=_tag_seq[19:])
    tickLblSkip = ZeroOrOne("c:tickLblSkip", successors=_tag_seq[20:])
    tickMarkSkip = ZeroOrOne("c:tickMarkSkip", successors=_tag_seq[21:])
    del _tag_seq


class CT_ChartLines(BaseOxmlElement):
    """Used for `c:majorGridlines` and `c:minorGridlines`.

    Specifies gridlines visual properties such as color and width.
    """

    spPr = ZeroOrOne("c:spPr", successors=())


class CT_Crosses(BaseOxmlElement):
    """`c:crosses` element, specifying where the other axis crosses this one."""

    val = RequiredAttribute("val", XL_AXIS_CROSSES)


class CT_CrossBetween(BaseOxmlElement):
    """`c:crossBetween` element, specifying how the value axis crosses categories.

    The `val` attribute takes either `between` or `midCat` — the values defined
    by `ST_CrossBetween` in `dml-chart.xsd`.
    """

    val = RequiredAttribute("val", XL_CROSS_BETWEEN)


class CT_DateAx(BaseAxisElement):
    """`c:dateAx` element, defining a date (category) axis."""

    _tag_seq = (
        "c:axId",
        "c:scaling",
        "c:delete",
        "c:axPos",
        "c:majorGridlines",
        "c:minorGridlines",
        "c:title",
        "c:numFmt",
        "c:majorTickMark",
        "c:minorTickMark",
        "c:tickLblPos",
        "c:spPr",
        "c:txPr",
        "c:crossAx",
        "c:crosses",
        "c:crossesAt",
        "c:auto",
        "c:lblOffset",
        "c:baseTimeUnit",
        "c:majorUnit",
        "c:majorTimeUnit",
        "c:minorUnit",
        "c:minorTimeUnit",
        "c:extLst",
    )
    scaling = OneAndOnlyOne("c:scaling")
    delete_ = ZeroOrOne("c:delete", successors=_tag_seq[3:])
    axPos = ZeroOrOne("c:axPos", successors=_tag_seq[4:])
    majorGridlines = ZeroOrOne("c:majorGridlines", successors=_tag_seq[5:])
    minorGridlines = ZeroOrOne("c:minorGridlines", successors=_tag_seq[6:])
    title = ZeroOrOne("c:title", successors=_tag_seq[7:])
    numFmt = ZeroOrOne("c:numFmt", successors=_tag_seq[8:])
    majorTickMark = ZeroOrOne("c:majorTickMark", successors=_tag_seq[9:])
    minorTickMark = ZeroOrOne("c:minorTickMark", successors=_tag_seq[10:])
    tickLblPos = ZeroOrOne("c:tickLblPos", successors=_tag_seq[11:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[12:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[13:])
    crosses = ZeroOrOne("c:crosses", successors=_tag_seq[15:])
    crossesAt = ZeroOrOne("c:crossesAt", successors=_tag_seq[16:])
    lblOffset = ZeroOrOne("c:lblOffset", successors=_tag_seq[18:])
    majorUnit = ZeroOrOne("c:majorUnit", successors=_tag_seq[20:])
    minorUnit = ZeroOrOne("c:minorUnit", successors=_tag_seq[22:])
    del _tag_seq


class CT_LblAlgn(BaseOxmlElement):
    """`c:lblAlgn` element, specifying category-axis tick-label horizontal alignment.

    The `val` attribute takes one of `ctr`, `l`, or `r` — the values defined by
    `ST_LblAlgn` in `dml-chart.xsd`.
    """

    val = RequiredAttribute("val", XL_TICK_LABEL_ALIGNMENT)


class CT_LblOffset(BaseOxmlElement):
    """`c:lblOffset` custom element class."""

    val = OptionalAttribute("val", ST_LblOffset, default=100)


class CT_Skip(BaseOxmlElement):
    """`c:tickLblSkip` and `c:tickMarkSkip` element class.

    Defines the integer skip count for category-axis tick labels or major tick
    marks — the interval at which categories are drawn (e.g. `val="2"` skips
    every second category). Minimum value is 1 (ECMA-376 Part 1 `ST_Skip`).
    """

    val = RequiredAttribute("val", ST_Skip)


class CT_Orientation(BaseOxmlElement):
    """`c:xAx/c:scaling/c:orientation` element, defining category order.

    Used to reverse the order categories appear in on a bar chart so they start at the
    top rather than the bottom. Because we read top-to-bottom, the default way looks odd
    to many and perhaps most folks. Also applicable to value and date axes.
    """

    val = OptionalAttribute("val", ST_Orientation, default=ST_Orientation.MIN_MAX)


class CT_Scaling(BaseOxmlElement):
    """`c:scaling` element.

    Defines axis scale characteristics such as maximum value, log vs. linear, etc.
    """

    _tag_seq = ("c:logBase", "c:orientation", "c:max", "c:min", "c:extLst")
    orientation = ZeroOrOne("c:orientation", successors=_tag_seq[2:])
    max = ZeroOrOne("c:max", successors=_tag_seq[3:])
    min = ZeroOrOne("c:min", successors=_tag_seq[4:])
    del _tag_seq

    @property
    def maximum(self):
        """
        The float value of the ``<c:max>`` child element, or |None| if no max
        element is present.
        """
        max = self.max
        if max is None:
            return None
        return max.val

    @maximum.setter
    def maximum(self, value):
        """
        Set the value of the ``<c:max>`` child element to the float *value*,
        or remove the max element if *value* is |None|.
        """
        self._remove_max()
        if value is None:
            return
        self._add_max(val=value)

    @property
    def minimum(self):
        """
        The float value of the ``<c:min>`` child element, or |None| if no min
        element is present.
        """
        min = self.min
        if min is None:
            return None
        return min.val

    @minimum.setter
    def minimum(self, value):
        """
        Set the value of the ``<c:min>`` child element to the float *value*,
        or remove the min element if *value* is |None|.
        """
        self._remove_min()
        if value is None:
            return
        self._add_min(val=value)


class CT_TickLblPos(BaseOxmlElement):
    """`c:tickLblPos` element."""

    val = OptionalAttribute("val", XL_TICK_LABEL_POSITION)


class CT_TickMark(BaseOxmlElement):
    """Used for `c:minorTickMark` and `c:majorTickMark`."""

    val = OptionalAttribute("val", XL_TICK_MARK, default=XL_TICK_MARK.CROSS)


class CT_ValAx(BaseAxisElement):
    """`c:valAx` element, defining a value axis."""

    _tag_seq = (
        "c:axId",
        "c:scaling",
        "c:delete",
        "c:axPos",
        "c:majorGridlines",
        "c:minorGridlines",
        "c:title",
        "c:numFmt",
        "c:majorTickMark",
        "c:minorTickMark",
        "c:tickLblPos",
        "c:spPr",
        "c:txPr",
        "c:crossAx",
        "c:crosses",
        "c:crossesAt",
        "c:crossBetween",
        "c:majorUnit",
        "c:minorUnit",
        "c:dispUnits",
        "c:extLst",
    )
    scaling = OneAndOnlyOne("c:scaling")
    delete_ = ZeroOrOne("c:delete", successors=_tag_seq[3:])
    axPos = ZeroOrOne("c:axPos", successors=_tag_seq[4:])
    majorGridlines = ZeroOrOne("c:majorGridlines", successors=_tag_seq[5:])
    minorGridlines = ZeroOrOne("c:minorGridlines", successors=_tag_seq[6:])
    title = ZeroOrOne("c:title", successors=_tag_seq[7:])
    numFmt = ZeroOrOne("c:numFmt", successors=_tag_seq[8:])
    majorTickMark = ZeroOrOne("c:majorTickMark", successors=_tag_seq[9:])
    minorTickMark = ZeroOrOne("c:minorTickMark", successors=_tag_seq[10:])
    tickLblPos = ZeroOrOne("c:tickLblPos", successors=_tag_seq[11:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[12:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[13:])
    crossAx = ZeroOrOne("c:crossAx", successors=_tag_seq[14:])
    crosses = ZeroOrOne("c:crosses", successors=_tag_seq[15:])
    crossesAt = ZeroOrOne("c:crossesAt", successors=_tag_seq[16:])
    crossBetween = ZeroOrOne("c:crossBetween", successors=_tag_seq[17:])
    majorUnit = ZeroOrOne("c:majorUnit", successors=_tag_seq[18:])
    minorUnit = ZeroOrOne("c:minorUnit", successors=_tag_seq[19:])
    del _tag_seq
