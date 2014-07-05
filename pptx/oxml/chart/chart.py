# encoding: utf-8

"""
lxml custom element classes for chart-related XML elements.
"""

from __future__ import absolute_import, print_function, unicode_literals

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

    @property
    def catAx(self):
        return self.chart.catAx

    @property
    def valAx(self):
        return self.chart.valAx


class CT_PlotArea(BaseOxmlElement):
    """
    ``<c:plotArea>`` element.
    """
    # catAx and valAx are actually ZeroOrMore, but don't need list bit yet
    catAx = ZeroOrOne('c:catAx')
    valAx = ZeroOrOne('c:valAx')


class CT_Style(BaseOxmlElement):
    """
    ``<c:style>`` element; defines the chart style.
    """
    val = RequiredAttribute('val', ST_Style)
