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
    idx = OneAndOnlyOne('c:idx')
    order = OneAndOnlyOne('c:order')
    tx = ZeroOrOne('c:tx')      # provide override for _insert_tx()
    spPr = ZeroOrOne('c:spPr')  # provide override for _insert_spPr()
    invertIfNegative = ZeroOrOne('c:invertIfNegative')  # provide _insert..()
    cat = ZeroOrOne('c:cat', successors=(
        'c:val', 'c:smooth', 'c:shape', 'c:extLst'
    ))
    val = ZeroOrOne('c:val', successors=('c:smooth', 'c:shape', 'c:extLst'))
    smooth = ZeroOrOne('c:smooth', successors=('c:extLst',))

    @property
    def val_pts(self):
        """
        The sequence of ``<c:pt>`` elements under the ``<c:val>`` child
        element, ordered by the value of their ``idx`` attribute.
        """
        val_pts = self.xpath('./c:val//c:pt')
        return sorted(val_pts, key=lambda pt: pt.idx)

    def _insert_invertIfNegative(self, invertIfNegative):
        """
        invertIfNegative has a lot of successors and they vary depending on
        the series type, so easier just to insert it "manually" as it's close
        to a required element.
        """
        if self.spPr is not None:
            self.spPr.addnext(invertIfNegative)
        elif self.tx is not None:
            self.tx.addnext(invertIfNegative)
        else:
            self.order.addnext(invertIfNegative)
        return invertIfNegative

    def _insert_spPr(self, spPr):
        """
        spPr has a lot of successors and it varies depending on the series
        type, so easier just to insert it "manually" as it's close to a
        required element.
        """
        if self.tx is not None:
            self.tx.addnext(spPr)
        else:
            self.order.addnext(spPr)
        return spPr

    def _insert_tx(self, tx):
        self.order.addnext(tx)
        return tx


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
