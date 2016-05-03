# encoding: utf-8

"""
Data label-related oxml objects.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .. import parse_xml
from ...enum.chart import XL_DATA_LABEL_POSITION
from ..ns import nsdecls
from ..text import CT_TextBody
from ..xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrOne


class CT_DLblPos(BaseOxmlElement):
    """
    ``<c:dLblPos>`` element specifying the positioning of a data label with
    respect to its data point.
    """
    val = RequiredAttribute('val', XL_DATA_LABEL_POSITION)


class CT_DLbls(BaseOxmlElement):
    """
    ``<c:dLbls>`` element specifying the properties of a set of data labels.
    """
    _tag_seq = (
        'c:numFmt', 'c:spPr', 'c:txPr', 'c:dLblPos', 'c:showLegendKey',
        'c:showVal', 'c:showCatName', 'c:showSerName', 'c:showPercent',
        'c:showBubbleSize', 'c:separator', 'c:showLeaderLines',
        'c:leaderLines', 'c:extLst'
    )
    numFmt = ZeroOrOne('c:numFmt', successors=(_tag_seq[1:]))
    txPr = ZeroOrOne('c:txPr', successors=(_tag_seq[3:]))
    dLblPos = ZeroOrOne('c:dLblPos', successors=(_tag_seq[4:]))
    del _tag_seq

    _default_xml = (
        '<c:dLbls %s>\n'
        '  <c:showLegendKey val="0"/>\n'
        '  <c:showVal val="1"/>\n'
        '  <c:showCatName val="0"/>\n'
        '  <c:showSerName val="0"/>\n'
        '  <c:showPercent val="0"/>\n'
        '  <c:showBubbleSize val="0"/>\n'
        '</c:dLbls>' % nsdecls('c')
    )

    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` great-great-grandchild element, added with its
        ancestors if not present.
        """
        txPr = self.get_or_add_txPr()
        defRPr = txPr.defRPr
        return defRPr

    @classmethod
    def new_default(cls):
        """
        Return a new default ``<c:dLbls>`` element.
        """
        xml = cls._default_xml
        dLbls = parse_xml(xml)
        return dLbls

    def _new_txPr(self):
        return CT_TextBody.new_txPr()
