# encoding: utf-8

"""
Composers for default chart XML for various chart types.
"""

from __future__ import absolute_import, print_function, unicode_literals

from xml.sax.saxutils import escape

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
            XL_CHART_TYPE.XY_SCATTER:       _XyChartXmlWriter,
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
        self._chart_data = series_seq
        self._series_lst = list(series_seq)

    @property
    def xml(self):
        """
        The full XML stream for the chart specified by this chart builder, as
        unicode text. This method must be overridden by each subclass.
        """
        raise NotImplementedError('must be implemented by all subclasses')


class _BaseSeriesXmlWriter(object):
    """
    Provides shared members for series XML writers.
    """
    def __init__(self, series):
        super(_BaseSeriesXmlWriter, self).__init__()
        self._series = series

    @property
    def name(self):
        """
        The XML-escaped name for this series.
        """
        return escape(self._series.name)

    def numRef_xml(self, wksht_ref, values):
        """
        Return the ``<c:numRef>`` element specified by the parameters as
        unicode text.
        """
        pt_xml = self.pt_xml(values)
        return (
            '            <c:numRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>General</c:formatCode>\n'
            '{pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
        ).format(
            wksht_ref=wksht_ref, pt_xml=pt_xml
        )

    def pt_xml(self, values):
        """
        Return the ``<c:ptCount>`` and sequence of ``<c:pt>`` elements
        corresponding to *values* as a single unicode text string.
        `c:ptCount` refers to the number of `c:pt` elements in this sequence.
        The `idx` attribute value for `c:pt` elements locates the data point
        in the overall data point sequence of the chart and is started at
        *offset*.
        """
        xml = (
            '                <c:ptCount val="{pt_count}"/>\n'
        ).format(
            pt_count=len(values)
        )

        pt_tmpl = (
            '                <c:pt idx="{idx}">\n'
            '                  <c:v>{value}</c:v>\n'
            '                </c:pt>\n'
        )
        for idx, value in enumerate(values):
            xml += pt_tmpl.format(idx=idx, value=value)

        return xml

    @property
    def tx_xml(self):
        """
        Return the ``<c:tx>`` (tx is short for 'text') element for this
        series as unicode text. This element contains the series name.
        """
        return (
            '          <c:tx>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="1"/>\n'
            '                <c:pt idx="0">\n'
            '                  <c:v>{series_name}</c:v>\n'
            '                </c:pt>\n'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:tx>\n'
        ).format(
            wksht_ref=self._series.name_ref, series_name=self.name
        )


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
            '  <c:txPr>\n'
            '    <a:bodyPr/>\n'
            '    <a:lstStyle/>\n'
            '    <a:p>\n'
            '      <a:pPr>\n'
            '        <a:defRPr sz="1800"/>\n'
            '      </a:pPr>\n'
            '      <a:endParaRPr lang="en-US"/>\n'
            '    </a:p>\n'
            '  </c:txPr>\n'
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
            '      <c:lineChart>\n'
            '        <c:grouping val="standard"/>\n'
            '%s'
            '        <c:axId val="2118791784"/>\n'
            '        <c:axId val="2140495176"/>\n'
            '      </c:lineChart>\n'
            '      <c:catAx>\n'
            '        <c:axId val="2118791784"/>\n'
            '        <c:scaling/>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="b"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="2140495176"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:lblAlgn val="ctr"/>\n'
            '        <c:lblOffset val="100"/>\n'
            '      </c:catAx>\n'
            '      <c:valAx>\n'
            '        <c:axId val="2140495176"/>\n'
            '        <c:scaling/>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            '        <c:majorGridlines/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="2118791784"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '      </c:valAx>\n'
            '    </c:plotArea>\n'
            '  </c:chart>\n'
            '  <c:txPr>\n'
            '    <a:bodyPr/>\n'
            '    <a:lstStyle/>\n'
            '    <a:p>\n'
            '      <a:pPr>\n'
            '        <a:defRPr sz="1800"/>\n'
            '      </a:pPr>\n'
            '      <a:endParaRPr lang="en-US"/>\n'
            '    </a:p>\n'
            '  </c:txPr>\n'
            '</c:chartSpace>\n'
        ) % self._ser_xml
        return xml

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._series_lst:
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="%d"/>\n'
                '          <c:order val="%d"/>\n'
                '%s'
                '          <c:marker>\n'
                '            <c:symbol val="none"/>\n'
                '          </c:marker>\n'
                '%s%s'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ) % (
                series.index, series.index, series.tx_xml, series.cat_xml,
                series.val_xml
            )
        return xml


class _PieChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:pieChart>`` element.
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
            '      <c:pieChart>\n'
            '        <c:varyColors val="1"/>\n'
            '%s'
            '      </c:pieChart>\n'
            '    </c:plotArea>\n'
            '  </c:chart>\n'
            '  <c:txPr>\n'
            '    <a:bodyPr/>\n'
            '    <a:lstStyle/>\n'
            '    <a:p>\n'
            '      <a:pPr>\n'
            '        <a:defRPr sz="1800"/>\n'
            '      </a:pPr>\n'
            '      <a:endParaRPr lang="en-US"/>\n'
            '    </a:p>\n'
            '  </c:txPr>\n'
            '</c:chartSpace>\n'
        ) % self._ser_xml
        return xml

    @property
    def _ser_xml(self):
        series = self._series_lst[0]
        xml = (
            '        <c:ser>\n'
            '          <c:idx val="0"/>\n'
            '          <c:order val="0"/>\n'
            '%s%s%s'
            '        </c:ser>\n'
        ) % (series.tx_xml, series.cat_xml, series.val_xml)
        return xml


class _XyChartXmlWriter(_BaseChartXmlWriter):
    """
    Generates XML for the ``<c:scatterChart>`` element.
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
            '      <c:scatterChart>\n'
            '        <c:scatterStyle val="lineMarker"/>\n'
            '        <c:varyColors val="0"/>\n'
            '%s'
            '        <c:axId val="-2128940872"/>\n'
            '        <c:axId val="-2129643912"/>\n'
            '      </c:scatterChart>\n'
            '      <c:valAx>\n'
            '        <c:axId val="-2128940872"/>\n'
            '        <c:scaling>\n'
            '          <c:orientation val="minMax"/>\n'
            '        </c:scaling>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="b"/>\n'
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2129643912"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            '      </c:valAx>\n'
            '      <c:valAx>\n'
            '        <c:axId val="-2129643912"/>\n'
            '        <c:scaling>\n'
            '          <c:orientation val="minMax"/>\n'
            '        </c:scaling>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            '        <c:majorGridlines/>\n'
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2128940872"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            '      </c:valAx>\n'
            '    </c:plotArea>\n'
            '    <c:legend>\n'
            '      <c:legendPos val="r"/>\n'
            '      <c:layout/>\n'
            '      <c:overlay val="0"/>\n'
            '    </c:legend>\n'
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="gap"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            '  </c:chart>\n'
            '  <c:txPr>\n'
            '    <a:bodyPr/>\n'
            '    <a:lstStyle/>\n'
            '    <a:p>\n'
            '      <a:pPr>\n'
            '        <a:defRPr sz="1800"/>\n'
            '      </a:pPr>\n'
            '      <a:endParaRPr lang="en-US"/>\n'
            '    </a:p>\n'
            '  </c:txPr>\n'
            '</c:chartSpace>\n'
        ) % self._ser_xml
        return xml

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._chart_data:
            xml_writer = _XySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '          <c:spPr>\n'
                '            <a:ln w="47625">\n'
                '              <a:noFill/>\n'
                '            </a:ln>\n'
                '          </c:spPr>\n'
                '{xVal_xml}'
                '{yVal_xml}'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ).format(
                ser_idx=series.index, ser_order=series.index,
                tx_xml=xml_writer.tx_xml, xVal_xml=xml_writer.xVal_xml,
                yVal_xml=xml_writer.yVal_xml
            )
        return xml


class _XySeriesXmlWriter(_BaseSeriesXmlWriter):
    """
    Generates XML snippets particular to an XY series.
    """
    @property
    def xVal_xml(self):
        """
        Return the ``<c:xVal>`` element for this series as unicode text. This
        element contains the X values for this series (which is actually all
        the X values for the entire chart).
        """
        return (
            '          <c:xVal>\n'
            '{numRef_xml}'
            '          </c:xVal>\n'
        ).format(
            numRef_xml=self.numRef_xml(
                self._series.x_values_ref, self._series.x_values
            )
        )

    @property
    def yVal_xml(self):
        """
        Return the ``<c:yVal>`` element for this series as unicode text. This
        element contains the Y values for this series.
        """
        return (
            '          <c:yVal>\n'
            '{numRef_xml}'
            '          </c:yVal>\n'
        ).format(
            numRef_xml=self.numRef_xml(
                self._series.y_values_ref, self._series.y_values
            )
        )
