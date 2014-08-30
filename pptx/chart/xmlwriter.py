# encoding: utf-8

"""
Composers for default chart XML for various chart types.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..enum.chart import XL_CHART_TYPE


def ChartXmlWriter(chart_type, series_seq):
    """
    Factory function returning appropriate XML writer object for
    *chart_type*, loaded with *series_seq*.
    """
    try:
        BuilderCls = {
            XL_CHART_TYPE.BAR_CLUSTERED:    _BarChartXmlWriter,
            XL_CHART_TYPE.BAR_STACKED_100:  _BarChartXmlWriter,
            XL_CHART_TYPE.COLUMN_CLUSTERED: _BarChartXmlWriter,
            XL_CHART_TYPE.LINE:             _LineChartXmlWriter,
            XL_CHART_TYPE.PIE:              _PieChartXmlWriter,
        }[chart_type]
    except KeyError:
        raise NotImplementedError(
            'XML writer for chart type %s not yet implemented' % chart_type
        )
    return BuilderCls(chart_type, series_seq)


class _BaseChartXmlWriter(object):
    """
    Generates XML text (unicode) for a default chart, like the one added by
    PowerPoint when you click the *Add Column Chart* button on the ribbon.
    Differentiated XML for different chart types is provided by subclasses.
    """
    def __init__(self, chart_type, series_seq):
        super(_BaseChartXmlWriter, self).__init__()
        self._chart_type = chart_type
        self._series_lst = list(series_seq)

    @property
    def xml(self):
        """
        The full XML stream for the chart specified by this chart builder, as
        unicode text. This method must be overridden by each subclass.
        """
        raise NotImplementedError('must be implemented by all subclasses')


class _BarChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:barChart>`` element.
    """
    @property
    def xml(self):
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:chart>\n'
            '    <c:plotArea>\n'
            '      <c:barChart>\n'
            '%s%s%s%s'
            '        <c:axId val="-2068027336"/>\n'
            '        <c:axId val="-2113994440"/>\n'
            '      </c:barChart>\n'
            '      <c:catAx>\n'
            '        <c:axId val="-2068027336"/>\n'
            '        <c:scaling/>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="%s"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2113994440"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:lblAlgn val="ctr"/>\n'
            '        <c:lblOffset val="100"/>\n'
            '      </c:catAx>\n'
            '      <c:valAx>\n'
            '        <c:axId val="-2113994440"/>\n'
            '        <c:scaling/>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="%s"/>\n'
            '        <c:majorGridlines/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2068027336"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '      </c:valAx>\n'
            '    </c:plotArea>\n'
            '  </c:chart>\n'
            '</c:chartSpace>\n'
        ) % (
            self._barDir_xml, self._grouping_xml, self._ser_xml,
            self._overlap_xml, self._cat_ax_pos, self._val_ax_pos
        )
        return xml

    @property
    def _barDir_xml(self):
        XL = XL_CHART_TYPE
        bar_types = (XL.BAR_CLUSTERED, XL.BAR_STACKED_100)
        col_types = (XL.COLUMN_CLUSTERED,)
        if self._chart_type in bar_types:
            return '        <c:barDir val="bar"/>\n'
        elif self._chart_type in col_types:
            return '        <c:barDir val="col"/>\n'
        raise NotImplementedError(
            'no _barDir_xml() for chart type %s' % self._chart_type
        )

    @property
    def _cat_ax_pos(self):
        return {
            XL_CHART_TYPE.BAR_CLUSTERED:    'l',
            XL_CHART_TYPE.BAR_STACKED_100:  'l',
            XL_CHART_TYPE.COLUMN_CLUSTERED: 'b',
        }[self._chart_type]

    @property
    def _grouping_xml(self):
        XL = XL_CHART_TYPE
        clustered_types = (XL.BAR_CLUSTERED, XL.COLUMN_CLUSTERED)
        percentStacked_types = (XL.BAR_STACKED_100,)
        if self._chart_type in clustered_types:
            return '        <c:grouping val="clustered"/>\n'
        elif self._chart_type in percentStacked_types:
            return '        <c:grouping val="percentStacked"/>\n'
        raise NotImplementedError(
            'no _grouping_xml() for chart type %s' % self._chart_type
        )

    @property
    def _overlap_xml(self):
        XL = XL_CHART_TYPE
        percentStacked_types = (XL.BAR_STACKED_100,)
        if self._chart_type in percentStacked_types:
            return '        <c:overlap val="100"/>\n'
        return ''

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._series_lst:
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="%d"/>\n'
                '          <c:order val="%d"/>\n'
                '%s%s%s'
                '        </c:ser>\n'
            ) % (
                series.index, series.index, series.tx_xml, series.cat_xml,
                series.val_xml
            )
        return xml

    @property
    def _val_ax_pos(self):
        return {
            XL_CHART_TYPE.BAR_CLUSTERED:    'b',
            XL_CHART_TYPE.BAR_STACKED_100:  'b',
            XL_CHART_TYPE.COLUMN_CLUSTERED: 'l',
        }[self._chart_type]


class _LineChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:lineChart>`` element.
    """


class _PieChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:pieChart>`` element.
    """
