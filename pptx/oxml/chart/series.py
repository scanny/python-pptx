# encoding: utf-8

"""
Series-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..simpletypes import XsdUnsignedInt
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne
)


class CT_NumDataSource(BaseOxmlElement):
    """
    ``<c:yVal>`` custom element class used in XY and bubble charts, and
    perhaps others.
    """
    numRef = OneAndOnlyOne('c:numRef')

    @property
    def ptCount_val(self):
        """
        Return the value of `./c:numRef/c:numCache/c:ptCount/@val`,
        specifying how many `c:pt` elements are in this numeric data cache.
        Returns 0 if no `c:ptCount` element is present, as this is the least
        disruptive way to degrade when no cached point data is available.
        This situation is not expected, but is valid according to the schema.
        """
        results = self.xpath('.//c:ptCount/@val')
        return int(results[0]) if results else 0

    def pt_v(self, idx):
        """
        Return the Y value for data point *idx* in this cache, or None if no
        value is present for that data point.
        """
        results = self.xpath('.//c:pt[@idx=%d]' % idx)
        return results[0].value if results else None


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
    yVal = ZeroOrOne('c:yVal', successors=_tag_seq[16:])
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

    @property
    def xVal_ptCount_val(self):
        """
        Return the number of X values as reflected in the `val` attribute of
        `./c:xVals/c:ptCount`, or 0 if not present.
        """
        vals = self.xpath('./c:xVal//c:ptCount/@val')
        if not vals:
            return 0
        return int(vals[0])

    @property
    def yVal_ptCount_val(self):
        """
        Return the number of Y values as reflected in the `val` attribute of
        `./c:yVals/c:ptCount`, or 0 if not present.
        """
        vals = self.xpath('./c:yVal//c:ptCount/@val')
        if not vals:
            return 0
        return int(vals[0])


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
