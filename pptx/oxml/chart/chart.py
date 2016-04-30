# encoding: utf-8

"""
lxml custom element classes for chart-related XML elements.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .. import parse_xml
from ..ns import nsdecls, qn
from ..simpletypes import ST_Style, XsdString
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne
)


class CT_Chart(BaseOxmlElement):
    """
    ``<c:chart>`` custom element class
    """
    plotArea = OneAndOnlyOne('c:plotArea')
    legend = ZeroOrOne('c:legend', successors=(
        'c:plotVisOnly', 'c:dispBlanksAs', 'c:showDLblsOverMax', 'c:extLst'
    ))
    rId = RequiredAttribute('r:id', XsdString)

    _chart_tmpl = (
        '<c:chart %s %s r:id="%%s"/>' % (nsdecls('c'), nsdecls('r'))
    )

    @property
    def has_legend(self):
        """
        True if this chart has a legend defined, False otherwise.
        """
        legend = self.legend
        if legend is None:
            return False
        return True

    @has_legend.setter
    def has_legend(self, bool_value):
        """
        Add, remove, or leave alone the ``<c:legend>`` child element depending
        on current state and *bool_value*. If *bool_value* is |True| and no
        ``<c:legend>`` element is present, a new default element is added.
        When |False|, any existing legend element is removed.
        """
        if bool(bool_value) is False:
            self._remove_legend()
        else:
            if self.legend is None:
                self._add_legend()

    @staticmethod
    def new_chart(rId):
        """
        Return a new ``<c:chart>`` element
        """
        xml = CT_Chart._chart_tmpl % (rId)
        chart = parse_xml(xml)
        return chart


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
    def catAx_lst(self):
        return self.chart.plotArea.catAx_lst

    @property
    def last_doc_order_ser(self):
        """
        Return the last `<c:ser>` element in the last xChart element, based
        on document order. Note this is not necessarily the same element as
        ``self.sers[-1]``.
        """
        last_xChart = list(self.chart.plotArea.iter_plots())[-1]
        sers = list(last_xChart.iter_sers())
        if not sers:
            return None
        return sers[-1]

    @property
    def sers(self):
        """
        An immutable sequence of the `c:ser` elements under this chartSpace
        element, sorted in order of their `c:ser/c:idx/@val` value and with
        any gaps in numbering collapsed.
        """
        def ser_idx(ser):
            return ser.idx.val

        sers = sorted(self.xpath('.//c:ser'), key=ser_idx)
        for idx, ser in enumerate(sers):
            if ser.idx.val != idx:
                ser.idx.val = idx
                ser.order.val = idx
        return sers

    @property
    def valAx_lst(self):
        return self.chart.plotArea.valAx_lst

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
    catAx = ZeroOrMore('c:catAx')
    valAx = ZeroOrMore('c:valAx')

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
