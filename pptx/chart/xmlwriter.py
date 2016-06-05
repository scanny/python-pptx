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
    XL_CT = XL_CHART_TYPE
    try:
        BuilderCls = {
            XL_CT.BAR_CLUSTERED:                _BarChartXmlWriter,
            XL_CT.BAR_STACKED_100:              _BarChartXmlWriter,
            XL_CT.BUBBLE:                       _BubbleChartXmlWriter,
            XL_CT.BUBBLE_THREE_D_EFFECT:        _BubbleChartXmlWriter,
            XL_CT.COLUMN_CLUSTERED:             _BarChartXmlWriter,
            XL_CT.LINE:                         _LineChartXmlWriter,
            XL_CT.PIE:                          _PieChartXmlWriter,
            XL_CT.RADAR:                        _RadarChartXmlWriter,
            XL_CT.RADAR_FILLED:                 _RadarChartXmlWriter,
            XL_CT.RADAR_MARKERS:                _RadarChartXmlWriter,
            XL_CT.XY_SCATTER:                   _XyChartXmlWriter,
            XL_CT.XY_SCATTER_LINES:             _XyChartXmlWriter,
            XL_CT.XY_SCATTER_LINES_NO_MARKERS:  _XyChartXmlWriter,
            XL_CT.XY_SCATTER_SMOOTH:            _XyChartXmlWriter,
            XL_CT.XY_SCATTER_SMOOTH_NO_MARKERS: _XyChartXmlWriter,
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
        self._series_seq = list(series_seq)

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
        return self._tx_tmpl.format(**{
            'wksht_ref':   self._series.name_ref,
            'series_name': self.name,
            'nsdecls':     '',
        })

    @property
    def _tx_tmpl(self):
        """
        The string formatting template for the ``<c:tx>`` element for this
        series, containing the series title and spreadsheet range reference.
        """
        return (
            '          <c:tx{nsdecls}>\n'
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
        for series in self._chart_data:
            xml_writer = _CategorySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '{cat_xml}'
                '{val_xml}'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':    series.index,
                'ser_order':  series.index,
                'tx_xml':     xml_writer.tx_xml,
                'cat_xml':    xml_writer.cat_xml,
                'val_xml':    xml_writer.val_xml,
            })
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
        for series in self._series_seq:
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
        series = self._series_seq[0]
        xml = (
            '        <c:ser>\n'
            '          <c:idx val="0"/>\n'
            '          <c:order val="0"/>\n'
            '%s%s%s'
            '        </c:ser>\n'
        ) % (series.tx_xml, series.cat_xml, series.val_xml)
        return xml


class _RadarChartXmlWriter(_BaseChartXmlWriter):
    """
    Generates XML for the ``<c:radarChart>`` element.
    """
    @property
    def xml(self):
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:date1904 val="0"/>\n'
            '  <c:roundedCorners val="0"/>\n'
            '  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.'
            'org/markup-compatibility/2006">\n'
            '    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/d'
            'rawing/2007/8/2/chart" Requires="c14">\n'
            '      <c14:style val="118"/>\n'
            '    </mc:Choice>\n'
            '    <mc:Fallback>\n'
            '      <c:style val="18"/>\n'
            '    </mc:Fallback>\n'
            '  </mc:AlternateContent>\n'
            '  <c:chart>\n'
            '    <c:plotArea>\n'
            '      <c:layout/>\n'
            '      <c:radarChart>\n'
            '        <c:radarStyle val="%s"/>\n'
            '        <c:varyColors val="0"/>\n'
            '%s'
            '        <c:axId val="2073612648"/>\n'
            '        <c:axId val="-2112772216"/>\n'
            '      </c:radarChart>\n'
            '      <c:catAx>\n'
            '        <c:axId val="2073612648"/>\n'
            '        <c:scaling>\n'
            '          <c:orientation val="minMax"/>\n'
            '        </c:scaling>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="b"/>\n'
            '        <c:majorGridlines/>\n'
            '        <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2112772216"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:auto val="1"/>\n'
            '        <c:lblAlgn val="ctr"/>\n'
            '        <c:lblOffset val="100"/>\n'
            '        <c:noMultiLvlLbl val="0"/>\n'
            '      </c:catAx>\n'
            '      <c:valAx>\n'
            '        <c:axId val="-2112772216"/>\n'
            '        <c:scaling>\n'
            '          <c:orientation val="minMax"/>\n'
            '        </c:scaling>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            '        <c:majorGridlines/>\n'
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="cross"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="2073612648"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="between"/>\n'
            '      </c:valAx>\n'
            '    </c:plotArea>\n'
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
        ) % (
            self._radar_style, self._ser_xml
        )
        return xml

    @property
    def _marker_xml(self):
        if self._chart_type == XL_CHART_TYPE.RADAR:
            return (
                '          <c:marker>\n'
                '            <c:symbol val="none"/>\n'
                '          </c:marker>\n'
            )
        return ''

    @property
    def _radar_style(self):
        if self._chart_type == XL_CHART_TYPE.RADAR_FILLED:
            return 'filled'
        return 'marker'

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._series_seq:
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="%d"/>\n'
                '          <c:order val="%d"/>\n'
                '%s'
                '%s'
                '%s%s'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ) % (
                series.index, series.index, series.tx_xml, self._marker_xml,
                series.cat_xml, series.val_xml
            )
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
            '        <c:scatterStyle val="%s"/>\n'
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
        ) % (self._scatterStyle_val, self._ser_xml)
        return xml

    @property
    def _marker_xml(self):
        no_marker_types = (
            XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
            XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
        )
        if self._chart_type in no_marker_types:
            return (
                '          <c:marker>\n'
                '            <c:symbol val="none"/>\n'
                '          </c:marker>\n'
            )
        return ''

    @property
    def _scatterStyle_val(self):
        smooth_types = (
            XL_CHART_TYPE.XY_SCATTER_SMOOTH,
            XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
        )
        if self._chart_type in smooth_types:
            return 'smoothMarker'
        return 'lineMarker'

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
                '{spPr_xml}'
                '{marker_xml}'
                '{xVal_xml}'
                '{yVal_xml}'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':    series.index,
                'ser_order':  series.index,
                'tx_xml':     xml_writer.tx_xml,
                'spPr_xml':   self._spPr_xml,
                'marker_xml': self._marker_xml,
                'xVal_xml':   xml_writer.xVal_xml,
                'yVal_xml':   xml_writer.yVal_xml,
            })
        return xml

    @property
    def _spPr_xml(self):
        if self._chart_type == XL_CHART_TYPE.XY_SCATTER:
            return (
                '          <c:spPr>\n'
                '            <a:ln w="47625">\n'
                '              <a:noFill/>\n'
                '            </a:ln>\n'
                '          </c:spPr>\n'
            )
        return ''


class _BubbleChartXmlWriter(_XyChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:bubbleChart>``
    element.
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
            '    <c:autoTitleDeleted val="0"/>\n'
            '    <c:plotArea>\n'
            '      <c:layout/>\n'
            '      <c:bubbleChart>\n'
            '        <c:varyColors val="0"/>\n'
            '%s'
            '        <c:dLbls>\n'
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            '        </c:dLbls>\n'
            '        <c:bubbleScale val="100"/>\n'
            '        <c:showNegBubbles val="0"/>\n'
            '        <c:axId val="-2115720072"/>\n'
            '        <c:axId val="-2115723560"/>\n'
            '      </c:bubbleChart>\n'
            '      <c:valAx>\n'
            '        <c:axId val="-2115720072"/>\n'
            '        <c:scaling>\n'
            '          <c:orientation val="minMax"/>\n'
            '        </c:scaling>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="b"/>\n'
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2115723560"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            '      </c:valAx>\n'
            '      <c:valAx>\n'
            '        <c:axId val="-2115723560"/>\n'
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
            '        <c:crossAx val="-2115720072"/>\n'
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
    def _bubble3D_val(self):
        if self._chart_type == XL_CHART_TYPE.BUBBLE_THREE_D_EFFECT:
            return '1'
        return '0'

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._chart_data:
            xml_writer = _BubbleSeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '          <c:invertIfNegative val="0"/>\n'
                # '{dLbls_xml}'
                '{xVal_xml}'
                '{yVal_xml}'
                '{bubbleSize_xml}'
                '          <c:bubble3D val="{bubble3D_val}"/>\n'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':        series.index,
                'ser_order':      series.index,
                'tx_xml':         xml_writer.tx_xml,
                # 'dLbls_xml':      xml_writer.dLbls_xml,
                'xVal_xml':       xml_writer.xVal_xml,
                'yVal_xml':       xml_writer.yVal_xml,
                'bubbleSize_xml': xml_writer.bubbleSize_xml,
                'bubble3D_val':   self._bubble3D_val,
            })
        return xml


class _CategorySeriesXmlWriter(_BaseSeriesXmlWriter):
    """
    Generates XML snippets particular to a category chart series.
    """
    @property
    def cat_xml(self):
        """
        The unicode XML snippet for the ``<c:cat>`` element for this series,
        containing the category labels and spreadsheet reference.
        """
        categories = self._series.categories
        if categories.depth == 1:
            return self._cat_tmpl.format(**{
                'wksht_ref':  self._series.categories_ref,
                'cat_count':  len(self._series.categories),
                'cat_pt_xml': self._cat_pt_xml,
                'nsdecls':    '',
            })
        return self._multiLvl_cat_tmpl.format(**{
            'wksht_ref': self._series.categories_ref,
            'cat_count': categories.leaf_count,
            'lvl_xml':   self._lvl_xml(categories),
            'nsdecls':   '',
        })

    @property
    def val_xml(self):
        """
        Return the unicode XML snippet for the ``<c:val>`` element describing
        this series, containing the series values and their spreadsheet range
        reference.
        """
        return self._val_tmpl.format(**{
            'wksht_ref':  self._series.values_ref,
            'val_count':  len(self._series),
            'val_pt_xml': self._val_pt_xml,
            'nsdecls':    '',
        })

    @property
    def _cat_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements containing the
        category names for this series.
        """
        xml = ''
        for idx, category in enumerate(self._series.categories):
            xml += (
                '                <c:pt idx="{cat_name_idx}">\n'
                '                  <c:v>{cat_name}</c:v>\n'
                '                </c:pt>\n'
            ).format(**{
                'cat_name_idx': idx,
                'cat_name':     escape(str(category.name)),
            })
        return xml

    @property
    def _cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series, containing
        the category labels and spreadsheet reference.
        """
        return (
            '          <c:cat{nsdecls}>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="{cat_count}"/>\n'
            '{cat_pt_xml}'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:cat>\n'
        )

    @property
    def _val_pt_xml(self):
        """
        The unicode XML snippet containing the ``<c:pt>`` elements containing
        the values for this series.
        """
        xml = ''
        for idx, value in enumerate(self._series.values):
            if value is None:
                continue
            xml += (
                '                <c:pt idx="{val_idx:d}">\n'
                '                  <c:v>{value}</c:v>\n'
                '                </c:pt>\n'
            ).format(**{
                'val_idx': idx,
                'value':   value,
            })
        return xml

    @property
    def _val_tmpl(self):
        """
        The template for the ``<c:val>`` element for this series, containing
        the series values and their spreadsheet range reference.
        """
        return (
            '          <c:val{nsdecls}>\n'
            '            <c:numRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>General</c:formatCode>\n'
            '                <c:ptCount val="{val_count}"/>\n'
            '{val_pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
            '          </c:val>\n'
        )


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


class _BubbleSeriesXmlWriter(_XySeriesXmlWriter):
    """
    Generates XML snippets particular to a bubble chart series.
    """
    @property
    def bubbleSize_xml(self):
        """
        Return the ``<c:bubbleSize>`` element for this series as unicode
        text. This element contains the bubble size values for all the
        data points in the chart.
        """
        return (
            '          <c:bubbleSize>\n'
            '{numRef_xml}'
            '          </c:bubbleSize>\n'
        ).format(
            numRef_xml=self.numRef_xml(
                self._series.bubble_sizes_ref, self._series.bubble_sizes
            )
        )
