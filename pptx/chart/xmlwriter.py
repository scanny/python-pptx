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


class _BarChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:barChart>`` element.
    """


class _LineChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:lineChart>`` element.
    """


class _PieChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:pieChart>`` element.
    """
