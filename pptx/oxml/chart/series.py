# encoding: utf-8

"""
Series-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..xmlchemy import BaseOxmlElement, OneAndOnlyOne, ZeroOrOne


class CT_SeriesComposite(BaseOxmlElement):
    """
    ``<c:ser>`` custom element class. Note there are several different series
    element types in the schema, such as ``CT_LineSer`` and ``CT_BarSer``,
    but they all share the same tag name. This class acts as a composite and
    depends on the caller not to do anything invalid for a series belonging
    to a particular plot type.
    """
    order = OneAndOnlyOne('c:order')
    tx = ZeroOrOne('c:tx')      # provide override for _insert_tx()
    spPr = ZeroOrOne('c:spPr')  # provide override for _insert_spPr()

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
