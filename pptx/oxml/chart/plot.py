# encoding: utf-8

"""
Plot-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .. import parse_xml
from ..ns import nsdecls
from ..simpletypes import ST_GapAmount
from ..xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class BaseChartElement(BaseOxmlElement):
    """
    Base class for barChart, lineChart, and other plot elements.
    """
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
    numFmt = ZeroOrOne('c:numFmt', successors=(
        'c:spPr', 'c:txPr', 'c:dLblPos', 'c:showLegendKey', 'c:showVal',
        'c:showCatName', 'c:showSerName', 'c:showPercent',
        'c:showBubbleSize', 'c:separator', 'c:showLeaderLines',
        'c:leaderLines', 'c:extLst'
    ))

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

    @classmethod
    def new_default(cls):
        """
        Return a new default ``<c:dLbls>`` element.
        """
        xml = cls._default_xml
        dLbls = parse_xml(xml)
        return dLbls


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
