"""Series-related oxml objects."""

from __future__ import annotations

from copy import deepcopy

from pptx.enum.chart import (
    XL_ERROR_BAR_DIRECTION,
    XL_ERROR_BAR_INCLUDE,
    XL_ERROR_BAR_TYPE,
    XL_TRENDLINE_TYPE,
)
from pptx.oxml import parse_xml
from pptx.oxml.chart.datalabel import CT_DLbls
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    OxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_AxDataSource(BaseOxmlElement):
    """
    ``<c:cat>`` custom element class used in category charts to specify
    category labels and hierarchy.
    """

    multiLvlStrRef = ZeroOrOne("c:multiLvlStrRef", successors=())

    @property
    def lvls(self):
        """
        Return a list containing the `c:lvl` descendent elements in document
        order. These will only be present when the required single child
        is a `c:multiLvlStrRef` element. Returns an empty list when no
        `c:lvl` descendent elements are present.
        """
        return self.xpath(".//c:lvl")


class CT_DPt(BaseOxmlElement):
    """
    ``<c:dPt>`` custom element class, containing visual properties for a data
    point.
    """

    _tag_seq = (
        "c:idx",
        "c:invertIfNegative",
        "c:marker",
        "c:bubble3D",
        "c:explosion",
        "c:spPr",
        "c:pictureOptions",
        "c:extLst",
    )
    idx = OneAndOnlyOne("c:idx")
    invertIfNegative = ZeroOrOne("c:invertIfNegative", successors=_tag_seq[2:])
    marker = ZeroOrOne("c:marker", successors=_tag_seq[3:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[6:])
    del _tag_seq

    @classmethod
    def new_dPt(cls):
        """
        Return a newly created "loose" `c:dPt` element containing its default
        subtree.
        """
        dPt = OxmlElement("c:dPt")
        dPt.append(OxmlElement("c:idx"))
        return dPt

    def _new_spPr(self):
        """Create a new `c:spPr` element pre-populated with inherited overrides.

        When a data-point-level `c:spPr` is introduced, PowerPoint treats it as a
        complete replacement for series-level shape properties -- any
        `a:ln`, `a:effectLst`, `a:effectDag`, `a:scene3d`, or `a:sp3d` present
        on the series `c:spPr` is no longer inherited at render time. To preserve
        round-trip fidelity (see issue #450, where a series-level shadow is lost
        when a single data point's fill color is overridden), pre-seed the new
        data-point `c:spPr` with deep copies of these non-fill children when the
        parent `c:ser` has its own `c:spPr`. The caller then adds a fill on top.
        """
        spPr = OxmlElement("c:spPr")
        ser = self.getparent()
        if ser is None:
            return spPr
        ser_spPr = ser.find(qn("c:spPr"))
        if ser_spPr is None:
            return spPr
        # -- copy non-fill children in schema order so that fill added later
        # -- inserts ahead of them (ln/effectLst/... are successors of the fill
        # -- choice group in CT_ShapeProperties). --
        _inherited_tags = (
            "a:ln",
            "a:effectLst",
            "a:effectDag",
            "a:scene3d",
            "a:sp3d",
        )
        for tag in _inherited_tags:
            child = ser_spPr.find(qn(tag))
            if child is not None:
                spPr.append(deepcopy(child))
        return spPr


class CT_Lvl(BaseOxmlElement):
    """
    ``<c:lvl>`` custom element class used in multi-level categories to
    specify a level of hierarchy.
    """

    pt = ZeroOrMore("c:pt", successors=())


class CT_NumDataSource(BaseOxmlElement):
    """
    ``<c:yVal>`` custom element class used in XY and bubble charts, and
    perhaps others.
    """

    numRef = OneAndOnlyOne("c:numRef")

    @property
    def ptCount_val(self):
        """
        Return the value of `./c:numRef/c:numCache/c:ptCount/@val`,
        specifying how many `c:pt` elements are in this numeric data cache.
        Returns 0 if no `c:ptCount` element is present, as this is the least
        disruptive way to degrade when no cached point data is available.
        This situation is not expected, but is valid according to the schema.
        """
        results = self.xpath(".//c:ptCount/@val")
        return int(results[0]) if results else 0

    def pt_v(self, idx):
        """
        Return the Y value for data point *idx* in this cache, or None if no
        value is present for that data point.
        """
        results = self.xpath(".//c:pt[@idx=%d]" % idx)
        return results[0].value if results else None


class CT_ErrBarType(BaseOxmlElement):
    """`c:errBarType` element specifying which side(s) of the data point the error bars appear.

    Child of `c:errBars`. Attribute `val` is one of "both", "minus", "plus" (default "both").
    """

    val = OptionalAttribute(
        "val", XL_ERROR_BAR_INCLUDE, default=XL_ERROR_BAR_INCLUDE.BOTH
    )  # pyright: ignore[reportAssignmentType]


class CT_ErrDir(BaseOxmlElement):
    """`c:errDir` element specifying the axis along which error bars extend.

    Child of `c:errBars`. Attribute `val` is one of "x" or "y".
    """

    val = RequiredAttribute("val", XL_ERROR_BAR_DIRECTION)  # pyright: ignore[reportAssignmentType]


class CT_ErrValType(BaseOxmlElement):
    """`c:errValType` element specifying how error-bar magnitudes are computed.

    Child of `c:errBars`. Attribute `val` is one of "cust", "fixedVal", "percentage",
    "stdDev", "stdErr" (default "fixedVal").
    """

    val = OptionalAttribute(
        "val", XL_ERROR_BAR_TYPE, default=XL_ERROR_BAR_TYPE.FIXED_VALUE
    )  # pyright: ignore[reportAssignmentType]


class CT_ErrBars(BaseOxmlElement):
    """`c:errBars` element describing a set of error bars for a data series.

    A series may have zero, one, or (on XY/bubble charts) two `c:errBars` children — in
    the two-child case one carries `errDir="x"` and the other `errDir="y"`.
    """

    _tag_seq = (
        "c:errDir",
        "c:errBarType",
        "c:errValType",
        "c:noEndCap",
        "c:plus",
        "c:minus",
        "c:val",
        "c:spPr",
        "c:extLst",
    )
    errDir = ZeroOrOne("c:errDir", successors=_tag_seq[1:])
    errBarType = ZeroOrOne("c:errBarType", successors=_tag_seq[2:])
    errValType = ZeroOrOne("c:errValType", successors=_tag_seq[3:])
    noEndCap = ZeroOrOne("c:noEndCap", successors=_tag_seq[4:])
    plus = ZeroOrOne("c:plus", successors=_tag_seq[5:])
    minus = ZeroOrOne("c:minus", successors=_tag_seq[6:])
    val = ZeroOrOne("c:val", successors=_tag_seq[7:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[8:])
    del _tag_seq

    @classmethod
    def new_errBars(cls, err_val_type="fixedVal", val=1.0, include="both", direction=None):
        """Return a newly-created `c:errBars` element with common defaults filled in.

        `err_val_type` is the OOXML string for `c:errValType/@val`
        ("fixedVal", "percentage", "stdDev", "stdErr", or "cust").
        `val` is the fixed magnitude for the `c:val` child (used when err_val_type is
        "fixedVal" or "percentage"); it is emitted regardless as it is harmless when
        the effective type ignores it.
        `include` is the OOXML string for `c:errBarType/@val` — "both", "plus", or "minus".
        `direction` is "x" or "y" (or |None| to omit `c:errDir`, which is valid for
        category-axis series).
        """
        parts = ["<c:errBars %s>" % nsdecls("c")]
        if direction is not None:
            parts.append('  <c:errDir val="%s"/>' % direction)
        parts.append('  <c:errBarType val="%s"/>' % include)
        parts.append('  <c:errValType val="%s"/>' % err_val_type)
        parts.append('  <c:noEndCap val="0"/>')
        parts.append('  <c:val val="%s"/>' % val)
        parts.append("</c:errBars>")
        return parse_xml("\n".join(parts))


class CT_TrendlineType(BaseOxmlElement):
    """`c:trendlineType` element specifying a trendline's regression type.

    Child of `c:trendline`. Attribute `val` is one of "exp", "linear", "log",
    "movingAvg", "poly", "power" (default "linear").
    """

    val = OptionalAttribute(
        "val", XL_TRENDLINE_TYPE, default=XL_TRENDLINE_TYPE.LINEAR
    )  # pyright: ignore[reportAssignmentType]


class CT_Trendline(BaseOxmlElement):
    """`c:trendline` element describing a regression/moving-average overlay for a series.

    A series may have zero or more `c:trendline` children (one per fitted curve
    the user wants drawn on the chart). The element carries a required
    `c:trendlineType` child plus several optional children that customize the
    fit (`c:order`, `c:period`, `c:forward`, `c:backward`, `c:intercept`) and
    its on-chart annotation (`c:dispEq`, `c:dispRSqr`).
    """

    _tag_seq = (
        "c:name",
        "c:spPr",
        "c:trendlineType",
        "c:order",
        "c:period",
        "c:forward",
        "c:backward",
        "c:intercept",
        "c:dispRSqr",
        "c:dispEq",
        "c:trendlineLbl",
        "c:extLst",
    )
    name = ZeroOrOne("c:name", successors=_tag_seq[1:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[2:])
    trendlineType = OneAndOnlyOne("c:trendlineType")
    order = ZeroOrOne("c:order", successors=_tag_seq[4:])
    period = ZeroOrOne("c:period", successors=_tag_seq[5:])
    forward = ZeroOrOne("c:forward", successors=_tag_seq[6:])
    backward = ZeroOrOne("c:backward", successors=_tag_seq[7:])
    intercept = ZeroOrOne("c:intercept", successors=_tag_seq[8:])
    dispRSqr = ZeroOrOne("c:dispRSqr", successors=_tag_seq[9:])
    dispEq = ZeroOrOne("c:dispEq", successors=_tag_seq[10:])
    del _tag_seq

    @classmethod
    def new_trendline(
        cls,
        trendline_type="linear",
        order=None,
        period=None,
        forward=None,
        backward=None,
        intercept=None,
        display_equation=False,
        display_r_squared=False,
    ):
        """Return a newly-created `c:trendline` element with common defaults filled in.

        `trendline_type` is the OOXML string for `c:trendlineType/@val`
        ("linear", "log", "poly", "power", "movingAvg", or "exp").
        `order` sets `c:order/@val` for polynomial trendlines (2..6).
        `period` sets `c:period/@val` for moving-average trendlines (>= 2).
        `forward` / `backward` extend the fitted curve that many category units
        beyond the last / before the first data point; floats, emitted only when
        not |None|.
        `intercept` forces the curve through a specific y-intercept; float,
        emitted only when not |None|.
        `display_equation` / `display_r_squared` control whether the fitted
        equation / R-squared value is drawn on the chart as a label.
        """
        parts = ["<c:trendline %s>" % nsdecls("c")]
        parts.append('  <c:trendlineType val="%s"/>' % trendline_type)
        if order is not None:
            parts.append('  <c:order val="%d"/>' % int(order))
        if period is not None:
            parts.append('  <c:period val="%d"/>' % int(period))
        if forward is not None:
            parts.append('  <c:forward val="%s"/>' % float(forward))
        if backward is not None:
            parts.append('  <c:backward val="%s"/>' % float(backward))
        if intercept is not None:
            parts.append('  <c:intercept val="%s"/>' % float(intercept))
        if display_r_squared:
            parts.append('  <c:dispRSqr val="1"/>')
        if display_equation:
            parts.append('  <c:dispEq val="1"/>')
        parts.append("</c:trendline>")
        return parse_xml("\n".join(parts))


class CT_SeriesComposite(BaseOxmlElement):
    """
    ``<c:ser>`` custom element class. Note there are several different series
    element types in the schema, such as ``CT_LineSer`` and ``CT_BarSer``,
    but they all share the same tag name. This class acts as a composite and
    depends on the caller not to do anything invalid for a series belonging
    to a particular plot type.
    """

    _tag_seq = (
        "c:idx",
        "c:order",
        "c:tx",
        "c:spPr",
        "c:invertIfNegative",
        "c:pictureOptions",
        "c:marker",
        "c:explosion",
        "c:dPt",
        "c:dLbls",
        "c:trendline",
        "c:errBars",
        "c:cat",
        "c:val",
        "c:xVal",
        "c:yVal",
        "c:shape",
        "c:smooth",
        "c:bubbleSize",
        "c:bubble3D",
        "c:extLst",
    )
    idx = OneAndOnlyOne("c:idx")
    order = OneAndOnlyOne("c:order")
    tx = ZeroOrOne("c:tx", successors=_tag_seq[3:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[4:])
    invertIfNegative = ZeroOrOne("c:invertIfNegative", successors=_tag_seq[5:])
    marker = ZeroOrOne("c:marker", successors=_tag_seq[7:])
    dPt = ZeroOrMore("c:dPt", successors=_tag_seq[9:])
    dLbls = ZeroOrOne("c:dLbls", successors=_tag_seq[10:])
    trendline = ZeroOrMore("c:trendline", successors=_tag_seq[11:])
    errBars = ZeroOrOne("c:errBars", successors=_tag_seq[12:])
    cat = ZeroOrOne("c:cat", successors=_tag_seq[13:])
    val = ZeroOrOne("c:val", successors=_tag_seq[14:])
    xVal = ZeroOrOne("c:xVal", successors=_tag_seq[15:])
    yVal = ZeroOrOne("c:yVal", successors=_tag_seq[16:])
    smooth = ZeroOrOne("c:smooth", successors=_tag_seq[18:])
    bubbleSize = ZeroOrOne("c:bubbleSize", successors=_tag_seq[19:])
    del _tag_seq

    @property
    def bubbleSize_ptCount_val(self):
        """
        Return the number of bubble size values as reflected in the `val`
        attribute of `./c:bubbleSize//c:ptCount`, or 0 if not present.
        """
        vals = self.xpath("./c:bubbleSize//c:ptCount/@val")
        if not vals:
            return 0
        return int(vals[0])

    @property
    def cat_ptCount_val(self):
        """
        Return the number of categories as reflected in the `val` attribute
        of `./c:cat//c:ptCount`, or 0 if not present.
        """
        vals = self.xpath("./c:cat//c:ptCount/@val")
        if not vals:
            return 0
        return int(vals[0])

    def get_dLbl(self, idx):
        """
        Return the `c:dLbl` element representing the label for the data point
        at offset *idx* in this series, or |None| if not present.
        """
        dLbls = self.dLbls
        if dLbls is None:
            return None
        return dLbls.get_dLbl_for_point(idx)

    def get_or_add_dLbl(self, idx):
        """
        Return the `c:dLbl` element representing the label of the point at
        offset *idx* in this series, newly created if not yet present.
        """
        dLbls = self.get_or_add_dLbls()
        return dLbls.get_or_add_dLbl_for_point(idx)

    def get_or_add_dPt_for_point(self, idx):
        """
        Return the `c:dPt` child representing the visual properties of the
        data point at index *idx*.
        """
        matches = self.xpath('c:dPt[c:idx[@val="%d"]]' % idx)
        if matches:
            return matches[0]
        dPt = self._add_dPt()
        dPt.idx.val = idx
        return dPt

    @property
    def xVal_ptCount_val(self):
        """
        Return the number of X values as reflected in the `val` attribute of
        `./c:xVal//c:ptCount`, or 0 if not present.
        """
        vals = self.xpath("./c:xVal//c:ptCount/@val")
        if not vals:
            return 0
        return int(vals[0])

    @property
    def yVal_ptCount_val(self):
        """
        Return the number of Y values as reflected in the `val` attribute of
        `./c:yVal//c:ptCount`, or 0 if not present.
        """
        vals = self.xpath("./c:yVal//c:ptCount/@val")
        if not vals:
            return 0
        return int(vals[0])

    def _new_dLbls(self):
        """Override metaclass method that creates `c:dLbls` element."""
        return CT_DLbls.new_dLbls()

    def _new_dPt(self):
        """
        Overrides the metaclass generated method to get `c:dPt` with minimal
        subtree.
        """
        return CT_DPt.new_dPt()


class CT_StrVal_NumVal_Composite(BaseOxmlElement):
    """
    ``<c:pt>`` element, can be either CT_StrVal or CT_NumVal complex type.
    Using this class for both, differentiating as needed.
    """

    v = OneAndOnlyOne("c:v")
    idx = RequiredAttribute("idx", XsdUnsignedInt)

    @property
    def value(self):
        """
        The float value of the text in the required ``<c:v>`` child.
        """
        return float(self.v.text)
