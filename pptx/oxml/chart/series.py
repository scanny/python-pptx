# encoding: utf-8

"""
Series-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..simpletypes import XsdUnsignedInt
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne
)


class CT_SeriesComposite(BaseOxmlElement):
    """
    ``<c:ser>`` custom element class. Note there are several different series
    element types in the schema, such as ``CT_LineSer`` and ``CT_BarSer``,
    but they all share the same tag name. This class acts as a composite and
    depends on the caller not to do anything invalid for a series belonging
    to a particular plot type.
    """
    _tag_seq = (
        'c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative',
        'c:pictureOptions', 'c:marker', 'c:explosion', 'c:dPt', 'c:dLbls',
        'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:xVal', 'c:yVal',
        'c:shape', 'c:smooth', 'c:bubbleSize', 'c:bubble3D', 'c:extLst'
    )
    idx = OneAndOnlyOne('c:idx')
    order = OneAndOnlyOne('c:order')
    tx = ZeroOrOne('c:tx', successors=_tag_seq[3:])
    spPr = ZeroOrOne('c:spPr', successors=_tag_seq[4:])
    invertIfNegative = ZeroOrOne(
        'c:invertIfNegative', successors=_tag_seq[5:]
    )
    cat = ZeroOrOne('c:cat', successors=_tag_seq[13:])
    val = ZeroOrOne('c:val', successors=_tag_seq[14:])
    smooth = ZeroOrOne('c:smooth', successors=_tag_seq[18:])
    del _tag_seq

    @property
    def val_pts(self):
        """
        The sequence of ``<c:pt>`` elements under the ``<c:val>`` child
        element, ordered by the value of their ``idx`` attribute.
        """
        val_pts = self.xpath('./c:val//c:pt')
        return sorted(val_pts, key=lambda pt: pt.idx)


class CT_StrVal_NumVal_Composite(BaseOxmlElement):
    """
    ``<c:pt>`` element, can be either CT_StrVal or CT_NumVal complex type.
    Using this class for both, differentiating as needed.
    """
    v = OneAndOnlyOne('c:v')
    idx = RequiredAttribute('idx', XsdUnsignedInt)

    @property
    def value(self):
        """
        The float value of the text in the required ``<c:v>`` child.
        """
        return float(self.v.text)
