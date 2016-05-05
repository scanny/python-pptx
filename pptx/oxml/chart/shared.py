# encoding: utf-8

"""
Shared oxml objects for charts.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..simpletypes import (
    ST_LayoutMode, XsdBoolean, XsdDouble, XsdString, XsdUnsignedInt
)
from ..xmlchemy import (
    BaseOxmlElement, OptionalAttribute, OxmlElement, RequiredAttribute,
    ZeroOrOne
)


class CT_Boolean(BaseOxmlElement):
    """
    Common complex type used for elements having a True/False value.
    """
    val = OptionalAttribute('val', XsdBoolean, default=True)


class CT_Double(BaseOxmlElement):
    """
    Used for floating point values.
    """
    val = RequiredAttribute('val', XsdDouble)


class CT_Layout(BaseOxmlElement):
    """
    ``<c:layout>`` custom element class
    """
    manualLayout = ZeroOrOne('c:manualLayout', successors=('c:extLst',))

    @property
    def horz_offset(self):
        """
        The float value in ./c:manualLayout/c:x when
        c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
        expression finds no match.
        """
        manualLayout = self.manualLayout
        if manualLayout is None:
            return 0.0
        return manualLayout.horz_offset

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:manualLayout/c:x@val to *offset* and
        ./c:manualLayout/c:xMode@val to "factor". Remove ./c:manualLayout if
        *offset* == 0.
        """
        if offset == 0.0:
            self._remove_manualLayout()
            return
        manualLayout = self.get_or_add_manualLayout()
        manualLayout.horz_offset = offset


class CT_LayoutMode(BaseOxmlElement):
    """
    Used for ``<c:xMode>``, ``<c:yMode>``, ``<c:wMode>``, and ``<c:hMode>``
    child elements of CT_ManualLayout.
    """
    val = OptionalAttribute(
        'val', ST_LayoutMode, default=ST_LayoutMode.FACTOR
    )


class CT_ManualLayout(BaseOxmlElement):
    """
    ``<c:manualLayout>`` custom element class
    """
    _tag_seq = (
        'c:layoutTarget', 'c:xMode', 'c:yMode', 'c:wMode', 'c:hMode', 'c:x',
        'c:y', 'c:w', 'c:h', 'c:extLst'
    )
    xMode = ZeroOrOne('c:xMode', successors=_tag_seq[2:])
    x = ZeroOrOne('c:x', successors=_tag_seq[6:])
    del _tag_seq

    @property
    def horz_offset(self):
        """
        The float value in ./c:x@val when ./c:xMode@val == "factor". 0.0 when
        ./c:x is not present or ./c:xMode@val != "factor".
        """
        x, xMode = self.x, self.xMode
        if x is None or xMode is None or xMode.val != ST_LayoutMode.FACTOR:
            return 0.0
        return x.val

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:x@val to *offset* and ./c:xMode@val to "factor".
        """
        self.get_or_add_xMode().val = ST_LayoutMode.FACTOR
        self.get_or_add_x().val = offset


class CT_NumFmt(BaseOxmlElement):
    """
    ``<c:numFmt>`` element specifying the formatting for number labels on a
    tick mark or data point.
    """
    formatCode = RequiredAttribute('formatCode', XsdString)
    sourceLinked = OptionalAttribute('sourceLinked', XsdBoolean)


class CT_Tx(BaseOxmlElement):
    """
    ``<c:tx>`` element containing the text for a label on a data point or
    other chart item.
    """
    strRef = ZeroOrOne('c:strRef')
    rich = ZeroOrOne('c:rich')

    def _new_rich(self):
        rich = OxmlElement('c:rich')
        rich.append(OxmlElement('a:bodyPr'))
        rich.add_p()
        return rich


class CT_UnsignedInt(BaseOxmlElement):
    """
    ``<c:idx>`` element and others.
    """
    val = RequiredAttribute('val', XsdUnsignedInt)
