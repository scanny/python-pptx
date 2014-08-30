# encoding: utf-8

"""
lxml custom element classes for chart-related XML elements.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..ns import qn
from ..simpletypes import ST_Style, XsdString
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne
)


class CT_Chart(BaseOxmlElement):
    """
    ``<c:chart>`` custom element class
    """
    plotArea = OneAndOnlyOne('c:plotArea')
    rId = RequiredAttribute('r:id', XsdString)

    @property
    def catAx(self):
        return self.plotArea.catAx

    @property
    def valAx(self):
        return self.plotArea.valAx


class CT_ChartSpace(BaseOxmlElement):
    """
    ``<c:chartSpace>`` element class, the root element of a chart part.
    """
    style = ZeroOrOne('c:style', successors=(
        'c:clrMapOvr', 'c:pivotSource', 'c:protection', 'c:chart'
    ))
    chart = OneAndOnlyOne('c:chart')
    externalData = ZeroOrOne('c:externalData', successors=(
        'c:printSettings', 'c:userShapes', 'c:extLst'
    ))

    @property
    def catAx(self):
        return self.chart.catAx

    @property
    def valAx(self):
        return self.chart.valAx

    @property
    def xlsx_part_rId(self):
        """
        The string in the required ``r:id`` attribute of the
        `<c:externalData>` child, or |None| if no externalData element is
        present.
        """
        externalData = self.externalData
        if externalData is None:
            return None
        return externalData.rId

    def _add_externalData(self):
        """
        Always add a ``<c:autoUpdate val="0"/>`` child so auto-updating
        behavior is off by default.
        """
        externalData = self._new_externalData()
        externalData._add_autoUpdate(val=False)
        self._insert_externalData(externalData)
        return externalData


class CT_ExternalData(BaseOxmlElement):
    """
    `<c:externalData>` element, defining link to embedded Excel package part
    containing the chart data.
    """
    autoUpdate = ZeroOrOne('c:autoUpdate')
    rId = RequiredAttribute('r:id', XsdString)


class CT_PlotArea(BaseOxmlElement):
    """
    ``<c:plotArea>`` element.
    """
    # catAx and valAx are actually ZeroOrMore, but don't need list bit yet
    catAx = ZeroOrOne('c:catAx')
    valAx = ZeroOrOne('c:valAx')

    def iter_plots(self):
        """
        Generate each xChart child element in the order it appears.
        """
        plot_tags = (
            qn('c:area3DChart'), qn('c:areaChart'), qn('c:bar3DChart'),
            qn('c:barChart'), qn('c:bubbleChart'), qn('c:doughnutChart'),
            qn('c:line3DChart'), qn('c:lineChart'), qn('c:ofPieChart'),
            qn('c:pie3DChart'), qn('c:pieChart'), qn('c:radarChart'),
            qn('c:scatterChart'), qn('c:stockChart'), qn('c:surface3DChart'),
            qn('c:surfaceChart')
        )

        for child in self.iterchildren():
            if child.tag not in plot_tags:
                continue
            yield child


class CT_Style(BaseOxmlElement):
    """
    ``<c:style>`` element; defines the chart style.
    """
    val = RequiredAttribute('val', ST_Style)
