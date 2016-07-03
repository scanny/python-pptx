# encoding: utf-8

"""
Plot-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .datalabel import CT_DLbls
from ..ns import qn
from ..simpletypes import (
    ST_BarDir, ST_BubbleScale, ST_GapAmount, ST_Grouping, ST_Overlap
)
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, ZeroOrOne, ZeroOrMore
)


class BaseChartElement(BaseOxmlElement):
    """
    Base class for barChart, lineChart, and other plot elements.
    """
    @property
    def cat_pts(self):
        """
        The sequence of ``<c:pt>`` elements under the ``<c:cat>`` element of
        the first series in this xChart element, ordered by the value of
        their ``idx`` attribute.
        """
        cat_pts = self.xpath('./c:ser[1]/c:cat//c:lvl[1]/c:pt')
        if not cat_pts:
            cat_pts = self.xpath('./c:ser[1]/c:cat//c:pt')
        return sorted(cat_pts, key=lambda pt: pt.idx)

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

    @property
    def sers(self):
        """
        Sequence of ``<c:ser>`` child elements in document order.
        """
        return list(self.iter_sers())

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
    _tag_seq = (
        'c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines',
        'c:axId', 'c:extLst'
    )
    grouping = ZeroOrOne('c:grouping', successors=_tag_seq[1:])
    varyColors = ZeroOrOne('c:varyColors', successors=_tag_seq[2:])
    ser = ZeroOrMore('c:ser', successors=_tag_seq[3:])
    dLbls = ZeroOrOne('c:dLbls', successors=_tag_seq[4:])
    del _tag_seq


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
    varyColors = ZeroOrOne('c:varyColors', successors=_tag_seq[3:])
    ser = ZeroOrMore('c:ser', successors=_tag_seq[4:])
    dLbls = ZeroOrOne('c:dLbls', successors=_tag_seq[5:])
    gapWidth = ZeroOrOne('c:gapWidth', successors=_tag_seq[6:])
    overlap = ZeroOrOne('c:overlap', successors=_tag_seq[7:])
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


class CT_BubbleChart(BaseChartElement):
    """
    ``<c:bubbleChart>`` custom element class
    """
    _tag_seq = (
        'c:varyColors', 'c:ser', 'c:dLbls', 'c:axId', 'c:bubble3D',
        'c:bubbleScale', 'c:showNegBubbles', 'c:sizeRepresents', 'c:axId',
        'c:extLst'
    )
    ser = ZeroOrMore('c:ser', successors=_tag_seq[2:])
    bubble3D = ZeroOrOne('c:bubble3D', successors=_tag_seq[5:])
    bubbleScale = ZeroOrOne('c:bubbleScale', successors=_tag_seq[6:])
    del _tag_seq


class CT_BubbleScale(BaseChartElement):
    """
    ``<c:bubbleScale>`` custom element class
    """
    val = OptionalAttribute('val', ST_BubbleScale, default=100)


class CT_DoughnutChart(BaseChartElement):
    """
    ``<c:doughnutChart>`` element.
    """
    _tag_seq = (
        'c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng', 'c:holeSize',
        'c:extLst'
    )
    varyColors = ZeroOrOne('c:varyColors', successors=_tag_seq[1:])
    ser = ZeroOrMore('c:ser', successors=_tag_seq[2:])
    dLbls = ZeroOrOne('c:dLbls', successors=_tag_seq[3:])
    del _tag_seq


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
    _tag_seq = (
        'c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines',
        'c:hiLowLines', 'c:upDownBars', 'c:marker', 'c:smooth', 'c:axId',
        'c:extLst'
    )
    grouping = ZeroOrOne('c:grouping', successors=(_tag_seq[1:]))
    varyColors = ZeroOrOne('c:varyColors', successors=_tag_seq[2:])
    ser = ZeroOrMore('c:ser', successors=_tag_seq[3:])
    dLbls = ZeroOrOne('c:dLbls', successors=(_tag_seq[4:]))
    del _tag_seq


class CT_Overlap(BaseOxmlElement):
    """
    ``<c:overlap>`` element specifying bar overlap as an integer percentage
    of bar width, in range -100 to 100.
    """
    val = OptionalAttribute('val', ST_Overlap, default=0)


class CT_PieChart(BaseChartElement):
    """
    ``<c:pieChart>`` custom element class
    """
    _tag_seq = (
        'c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng', 'c:extLst'
    )
    varyColors = ZeroOrOne('c:varyColors', successors=_tag_seq[1:])
    ser = ZeroOrMore('c:ser', successors=_tag_seq[2:])
    dLbls = ZeroOrOne('c:dLbls', successors=_tag_seq[3:])
    del _tag_seq


class CT_RadarChart(BaseChartElement):
    """
    ``<c:radarChart>`` custom element class
    """
    _tag_seq = (
        'c:radarStyle', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:axId',
        'c:extLst'
    )
    varyColors = ZeroOrOne('c:varyColors', successors=_tag_seq[2:])
    ser = ZeroOrMore('c:ser', successors=_tag_seq[3:])
    dLbls = ZeroOrOne('c:dLbls', successors=(_tag_seq[4:]))
    del _tag_seq


class CT_ScatterChart(BaseChartElement):
    """
    ``<c:scatterChart>`` custom element class
    """
    _tag_seq = (
        'c:scatterStyle', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:axId',
        'c:extLst'
    )
    varyColors = ZeroOrOne('c:varyColors', successors=_tag_seq[2:])
    ser = ZeroOrMore('c:ser', successors=_tag_seq[3:])
    del _tag_seq
