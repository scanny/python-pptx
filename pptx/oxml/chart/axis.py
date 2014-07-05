# encoding: utf-8

"""
lxml custom element classes for chart axis-related XML elements.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ...enum.chart import XL_TICK_MARK
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, ZeroOrOne
)


class BaseAxisElement(BaseOxmlElement):
    """
    Base class for catAx, valAx, and perhaps other axis elements.
    """
    scaling = OneAndOnlyOne('c:scaling')
    delete = ZeroOrOne('c:delete', successors=('c:axPos',))
    majorTickMark = ZeroOrOne('c:majorTickMark', successors=(
        'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx'
    ))


class CT_CatAx(BaseAxisElement):
    """
    ``<c:catAx>`` element, defining a category axis.
    """


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


class CT_TickMark(BaseOxmlElement):
    """
    Used for ``<c:minorTickMark>`` and ``<c:majorTickMark>``.
    """
    val = OptionalAttribute('val', XL_TICK_MARK, default=XL_TICK_MARK.CROSS)


class CT_ValAx(BaseAxisElement):
    """
    ``<c:valAx>`` element, defining a value axis.
    """
