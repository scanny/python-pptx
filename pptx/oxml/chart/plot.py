# encoding: utf-8

"""
Plot-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .. import parse_xml
from ..ns import nsdecls, qn
from ..simpletypes import ST_BarDir, ST_GapAmount, ST_Grouping
from ..text import CT_TextBody
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, ZeroOrOne
)


class BaseChartElement(BaseOxmlElement):
    """
    Base class for barChart, lineChart, and other plot elements.
    """
    @property
    def grouping_val(self):
        """
        Return the value of the ``./c:grouping{val=?}`` attribute, taking
        defaults into account when items are not present.
        """
        grouping = self.grouping
        if grouping is None:
            return ST_Grouping.STANDARD
        val = grouping.val
        if val is None:
            return ST_Grouping.STANDARD
        return val

    def iter_sers(self):
        """
        Generate each ``<c:ser>`` child element in the order it appears.
        """
        for child in self.iterchildren():
            if child.tag == qn('c:ser'):
                yield child

    def _new_dLbls(self):
        return CT_DLbls.new_default()


class CT_Area3DChart(BaseChartElement):
    """
    ``<c:area3DChart>`` element.
    """
    grouping = ZeroOrOne('c:grouping', successors=(
        'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:gapDepth',
        'c:axId'
    ))


class CT_AreaChart(BaseChartElement):
    """
    ``<c:areaChart>`` element.
    """
    grouping = ZeroOrOne('c:grouping', successors=(
        'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:axId'
    ))


class CT_BarChart(BaseChartElement):
    """
    ``<c:barChart>`` element.
    """
    _tag_seq = (
        'c:barDir', 'c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls',
        'c:gapWidth', 'c:overlap', 'c:serLines', 'c:axId', 'c:extLst'
    )
    barDir = OneAndOnlyOne('c:barDir')
    grouping = ZeroOrOne('c:grouping', successors=_tag_seq[2:])
    dLbls = ZeroOrOne('c:dLbls', successors=_tag_seq[5:])
    gapWidth = ZeroOrOne('c:gapWidth', successors=_tag_seq[6:])
    del _tag_seq

    @property
    def grouping_val(self):
        """
        Return the value of the ``./c:grouping{val=?}`` attribute, taking
        defaults into account when items are not present.
        """
        grouping = self.grouping
        if grouping is None:
            return ST_Grouping.CLUSTERED
        val = grouping.val
        if val is None:
            return ST_Grouping.CLUSTERED
        return val


class CT_BarDir(BaseOxmlElement):
    """
    ``<c:barDir>`` child of a barChart element, specifying the orientation of
    the bars, 'bar' if they are horizontal and 'col' if they are vertical.
    """
    val = OptionalAttribute('val', ST_BarDir, default=ST_BarDir.COL)


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


class CT_Grouping(BaseOxmlElement):
    """
    ``<c:grouping>`` child of an xChart element, specifying a value like
    'clustered' or 'stacked'. Also used for variants with the same tag name
    like CT_BarGrouping.
    """
    val = OptionalAttribute('val', ST_Grouping)


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
