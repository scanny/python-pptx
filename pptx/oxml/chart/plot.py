# encoding: utf-8

"""
Plot-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .. import parse_xml
from ..ns import nsdecls, qn
from ..simpletypes import ST_GapAmount
from ..text import CT_TextBody
from ..xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class BaseChartElement(BaseOxmlElement):
    """
    Base class for barChart, lineChart, and other plot elements.
    """
    def iter_sers(self):
        """
        Generate each ``<c:ser>`` child element in the order it appears.
        """
        for child in self.iterchildren():
            if child.tag == qn('c:ser'):
                yield child

    def _new_dLbls(self):
        return CT_DLbls.new_default()


class CT_BarChart(BaseChartElement):
    """
    ``<c:barChart>`` element.
    """
    dLbls = ZeroOrOne('c:dLbls', successors=(
        'c:gapWidth', 'c:overlap', 'c:serLines', 'c:axId'
    ))
    gapWidth = ZeroOrOne('c:gapWidth', successors=(
        'c:overlap', 'c:serLines', 'c:axId'
    ))


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


class CT_GapAmount(BaseOxmlElement):
    """
    ``<c:gapWidth>`` child of ``<c:barChart>`` element, also used for other
    purposes like error bars.
    """
    val = OptionalAttribute('val', ST_GapAmount, default=150)


class CT_LineChart(BaseChartElement):
    """
    ``<c:lineChart>`` custom element class
    """
    dLbls = ZeroOrOne('c:dLbls', successors=(
        'c:dropLines', 'c:hiLowLines', 'c:upDownBars', 'c:marker',
        'c:smooth', 'c:axId'
    ))


class CT_PieChart(BaseChartElement):
    """
    ``<c:pieChart>`` custom element class
    """
    dLbls = ZeroOrOne('c:dLbls', successors=('c:firstSliceAng', 'c:extLst'))
