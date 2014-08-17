# encoding: utf-8

"""
Axis-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ...enum.chart import XL_TICK_LABEL_POSITION, XL_TICK_MARK
from ..simpletypes import ST_AxisUnit, ST_LblOffset
from ..text import CT_TextBody
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute,
    ZeroOrOne
)


class BaseAxisElement(BaseOxmlElement):
    """
    Base class for catAx, valAx, and perhaps other axis elements.
    """
    scaling = OneAndOnlyOne('c:scaling')
    delete = ZeroOrOne('c:delete', successors=('c:axPos',))
    majorGridlines = ZeroOrOne('c:majorGridlines', successors=(
        'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark',
        'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx'
    ))
    minorGridlines = ZeroOrOne('c:minorGridlines', successors=(
        'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark',
        'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx'
    ))
    numFmt = ZeroOrOne('c:numFmt', successors=(
        'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr',
        'c:txPr', 'c:crossAx'
    ))
    majorTickMark = ZeroOrOne('c:majorTickMark', successors=(
        'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx'
    ))
    minorTickMark = ZeroOrOne('c:minorTickMark', successors=(
        'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx'
    ))
    tickLblPos = ZeroOrOne('c:tickLblPos', successors=(
        'c:spPr', 'c:txPr', 'c:crossAx'
    ))
    txPr = ZeroOrOne('c:txPr', successors=('c:crossAx',))

    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` great-great-grandchild element, added with its
        ancestors if not present.
        """
        txPr = self.get_or_add_txPr()
        defRPr = txPr.defRPr
        return defRPr

    def _new_txPr(self):
        return CT_TextBody.new_txPr()


class CT_AxisUnit(BaseOxmlElement):
    """
    Used for ``<c:majorUnit>`` and ``<c:minorUnit>`` elements, and others.
    """
    val = RequiredAttribute('val', ST_AxisUnit)


class CT_CatAx(BaseAxisElement):
    """
    ``<c:catAx>`` element, defining a category axis.
    """
    lblOffset = ZeroOrOne('c:lblOffset', successors=(
        'c:tickLblSkip', 'c:tickMarkSkip', 'c:noMultiLvlLbl', 'c:extLst'
    ))


class CT_LblOffset(BaseOxmlElement):
    """
    ``<c:lblOffset>`` custom element class
    """
    val = OptionalAttribute('val', ST_LblOffset, default=100)


class CT_Scaling(BaseOxmlElement):
    """
    ``<c:scaling>`` element, defining axis scale characteristics such as
    maximum value, log vs. linear, etc.
    """
    max = ZeroOrOne('c:max', successors=('c:min', 'c:extLst'))
    min = ZeroOrOne('c:min', successors=('c:extLst',))

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
    val = OptionalAttribute('val', XL_TICK_LABEL_POSITION)


class CT_TickMark(BaseOxmlElement):
    """
    Used for ``<c:minorTickMark>`` and ``<c:majorTickMark>``.
    """
    val = OptionalAttribute('val', XL_TICK_MARK, default=XL_TICK_MARK.CROSS)


class CT_ValAx(BaseAxisElement):
    """
    ``<c:valAx>`` element, defining a value axis.
    """
    majorUnit = ZeroOrOne('c:majorUnit', successors=(
        'c:minorUnit', 'c:dispUnits', 'c:extLst'
    ))
    minorUnit = ZeroOrOne('c:minorUnit', successors=(
        'c:dispUnits', 'c:extLst'
    ))
