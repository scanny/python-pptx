# encoding: utf-8

"""
Axis-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ...enum.chart import XL_AXIS_CROSSES, XL_TICK_LABEL_POSITION, XL_TICK_MARK
from .shared import CT_Title
from ..simpletypes import ST_AxisUnit, ST_LblOffset
from ..text import CT_TextBody
from ..xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class BaseAxisElement(BaseOxmlElement):
    """
    Base class for catAx, valAx, and perhaps other axis elements.
    """

    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` great-great-grandchild element, added with its
        ancestors if not present.
        """
        txPr = self.get_or_add_txPr()
        defRPr = txPr.defRPr
        return defRPr

    def _new_title(self):
        return CT_Title.new_title()

    def _new_txPr(self):
        return CT_TextBody.new_txPr()


class CT_AxisUnit(BaseOxmlElement):
    """
    Used for ``<c:majorUnit>`` and ``<c:minorUnit>`` elements, and others.
    """

    val = RequiredAttribute("val", ST_AxisUnit)


class CT_CatAx(BaseAxisElement):
    """
    ``<c:catAx>`` element, defining a category axis.
    """

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
    lblOffset = ZeroOrOne("c:lblOffset", successors=_tag_seq[19:])
    del _tag_seq


class CT_ChartLines(BaseOxmlElement):
    """
    Used for c:majorGridlines and c:minorGridlines, specifies gridlines
    visual properties such as color and width.
    """

    spPr = ZeroOrOne("c:spPr", successors=())


class CT_Crosses(BaseAxisElement):
    """
    ``<c:crosses>`` element, specifying where the other axis crosses this
    one.
    """

    val = RequiredAttribute("val", XL_AXIS_CROSSES)


class CT_DateAx(BaseAxisElement):
    """
    ``<c:dateAx>`` element, defining a date (category) axis.
    """

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
    del _tag_seq


class CT_LblOffset(BaseOxmlElement):
    """
    ``<c:lblOffset>`` custom element class
    """

    val = OptionalAttribute("val", ST_LblOffset, default=100)


class CT_Scaling(BaseOxmlElement):
    """
    ``<c:scaling>`` element, defining axis scale characteristics such as
    maximum value, log vs. linear, etc.
    """

    max = ZeroOrOne("c:max", successors=("c:min", "c:extLst"))
    min = ZeroOrOne("c:min", successors=("c:extLst",))

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
    """
    ``<c:tickLblPos>`` element.
    """

    val = OptionalAttribute("val", XL_TICK_LABEL_POSITION)


class CT_TickMark(BaseOxmlElement):
    """
    Used for ``<c:minorTickMark>`` and ``<c:majorTickMark>``.
    """

    val = OptionalAttribute("val", XL_TICK_MARK, default=XL_TICK_MARK.CROSS)


class CT_ValAx(BaseAxisElement):
    """
    ``<c:valAx>`` element, defining a value axis.
    """

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
    majorUnit = ZeroOrOne("c:majorUnit", successors=_tag_seq[18:])
    minorUnit = ZeroOrOne("c:minorUnit", successors=_tag_seq[19:])
    del _tag_seq
