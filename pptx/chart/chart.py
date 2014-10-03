# encoding: utf-8

"""
Chart shape-related objects such as Chart.
"""

from __future__ import absolute_import, print_function, unicode_literals

from collections import Sequence
from copy import deepcopy

from .axis import CategoryAxis, ValueAxis
from .legend import Legend
from .plot import PlotFactory, PlotTypeInspector
from .series import SeriesCollection
from ..util import lazyproperty


class Chart(object):
    """
    A chart object.
    """
    def __init__(self, chartSpace, chart_part):
        super(Chart, self).__init__()
        self._chartSpace = chartSpace
        self._chart_part = chart_part

    @property
    def category_axis(self):
        """
        The category axis of this chart. Raises |ValueError| if no category
        axis is defined.
        """
        catAx = self._chartSpace.catAx
        if catAx is None:
            raise ValueError('chart has no category axis')
        return CategoryAxis(catAx)

    @property
    def chart_style(self):
        """
        Read/write integer index of chart style used to format this chart.
        Range is from 1 to 48. Value is |None| if no explicit style has been
        assigned, in which case the default chart style is used. Assigning
        |None| causes any explicit setting to be removed. The integer index
        corresponds to the style's position in the chart style gallery in the
        PowerPoint UI.
        """
        style = self._chartSpace.style
        if style is None:
            return None
        return style.val

    @chart_style.setter
    def chart_style(self, value):
        self._chartSpace._remove_style()
        if value is None:
            return
        self._chartSpace._add_style(val=value)

    @property
    def chart_type(self):
        """
        Read-only :ref:`XlChartType` enumeration value specifying the type of
        this chart. If the chart has two plots, for example, a line plot
        overlayed on a bar plot, the type reported is for the first
        (back-most) plot.
        """
        first_plot = self.plots[0]
        return PlotTypeInspector.chart_type(first_plot)

    @property
    def has_legend(self):
        """
        Read/write boolean, |True| if the chart has a legend. Assigning
        |True| causes a legend to be added to the chart if it doesn't already
        have one. Assigning False removes any existing legend definition
        along with any existing legend settings.
        """
        return self._chartSpace.chart.has_legend

    @has_legend.setter
    def has_legend(self, value):
        self._chartSpace.chart.has_legend = bool(value)

    @property
    def legend(self):
        """
        A |Legend| object providing access to the properties of the legend
        for this chart.
        """
        legend_elm = self._chartSpace.chart.legend
        if legend_elm is None:
            return None
        return Legend(legend_elm)

    @lazyproperty
    def plots(self):
        """
        The sequence of plots in this chart. A plot, called a *chart group*
        in the Microsoft API, is a distinct sequence of one or more series
        depicted in a particular charting type. For example, a chart having
        a series plotted as a line overlaid on three series plotted as
        columns would have two plots; the first corresponding to the three
        column series and the second to the line series. Plots are sequenced
        in the order drawn, i.e. back-most to front-most. Supports *len()*,
        membership (e.g. ``p in plots``), iteration, slicing, and indexed
        access (e.g. ``plot = plots[i]``).
        """
        plotArea = self._chartSpace.chart.plotArea
        return _Plots(plotArea, self)

    def replace_data(self, chart_data):
        """
        Use the categories and series values in the |ChartData| object
        *chart_data* to replace those in the XML and Excel worksheet for this
        chart.
        """
        _SeriesRewriter.replace_series_data(self._chartSpace, chart_data)
        self._workbook.update_from_xlsx_blob(chart_data.xlsx_blob)

    @lazyproperty
    def series(self):
        """
        The |SeriesCollection| object containing all the series in this
        chart.
        """
        return SeriesCollection(self._chartSpace)

    @property
    def value_axis(self):
        """
        The |ValueAxis| object providing access to properties of the value
        axis of this chart. Raises |ValueError| if the chart has no value
        axis.
        """
        valAx = self._chartSpace.valAx
        if valAx is None:
            raise ValueError('chart has no value axis')
        return ValueAxis(valAx)

    @property
    def _workbook(self):
        """
        The |ChartWorkbook| object providing access to the Excel source data
        for this chart.
        """
        return self._chart_part.chart_workbook


class _Plots(Sequence):
    """
    The sequence of plots in a chart, such as a bar plot or a line plot. Most
    charts have only a single plot. The concept is necessary when two chart
    types are displayed in a single set of axes, like a bar plot with
    a superimposed line plot.
    """
    def __init__(self, plotArea, chart):
        super(_Plots, self).__init__()
        self._plotArea = plotArea
        self._chart = chart

    def __getitem__(self, index):
        xCharts = list(self._plotArea.iter_plots())
        if isinstance(index, slice):
            plots = [PlotFactory(xChart, self._chart) for xChart in xCharts]
            return plots[index]
        else:
            xChart = xCharts[index]
            return PlotFactory(xChart, self._chart)

    def __len__(self):
        plot_elms = [p for p in self._plotArea.iter_plots()]
        return len(plot_elms)


class _SeriesRewriter(object):
    """
    Pure functional service class that knows how to rewrite the data in
    a chart while disturbing the current formatting as little as possible.
    All the data for a chart is located under its ``<c:ser>`` elements, so
    updating the data for a chart requires changes to series element subtrees
    only.
    """
    @classmethod
    def replace_series_data(cls, chartSpace, chart_data):
        """
        Use the category and series data in *chart_data* to rewrite the
        series name, category labels, and point values of the series
        contained in the *chartSpace* element. All series-level formatting is
        left undisturbed. If *chart_data* contains fewer series than
        *chartSpace*, the excess series are deleted. If *chart_data* contains
        more series than the *chartSpace* element, new series are added to
        the last plot in the chart and series formatting is copied from the
        last series in that plot.
        """
        sers = cls._adjust_ser_count(chartSpace, len(chart_data.series))
        for ser, series_data in zip(sers, chart_data.series):
            cls._rewrite_ser_data(ser, series_data)

    @classmethod
    def _add_cloned_sers(cls, chartSpace, count):
        """
        Add `c:ser` elements to the last xChart element in *chartSpace*,
        cloned from the last `c:ser` child of that xChart.
        """
        def clone_ser(ser, idx):
            new_ser = deepcopy(ser)
            new_ser.idx.val = idx
            new_ser.order.val = idx
            ser.addnext(new_ser)
            return new_ser

        last_ser = chartSpace.last_doc_order_ser
        starting_idx = len(chartSpace.sers)
        for idx in range(starting_idx, starting_idx+count):
            last_ser = clone_ser(last_ser, idx)

    @classmethod
    def _adjust_ser_count(cls, chartSpace, new_ser_count):
        """
        Return the ser elements in *chartSpace* after adjusting their number
        to *new_ser_count*. The ser elements returned are sorted in
        increasing order of the c:ser/c:idx value, starting with 0 and with
        any gaps in numbering collapsed.
        """
        ser_count_diff = new_ser_count - len(chartSpace.sers)
        if ser_count_diff > 0:
            cls._add_cloned_sers(chartSpace, ser_count_diff)
        elif ser_count_diff < 0:
            cls._trim_ser_count_by(chartSpace, abs(ser_count_diff))
        return chartSpace.sers

    @classmethod
    def _rewrite_ser_data(cls, ser, series_data):
        """
        Rewrite the ``<c:tx>``, ``<c:cat>`` and ``<c:val>`` child elements
        of *ser* based on the values in *series_data*.
        """
        ser._remove_tx()
        ser._remove_cat()
        ser._remove_val()
        ser._insert_tx(series_data.tx)
        ser._insert_cat(series_data.cat)
        ser._insert_val(series_data.val)

    @classmethod
    def _trim_ser_count_by(cls, chartSpace, count):
        """
        Remove the last *count* ser elements from *chartSpace*. Any xChart
        elements having no ser child elements after trimming are also
        removed.
        """
        extra_sers = chartSpace.sers[-count:]
        for ser in extra_sers:
            parent = ser.getparent()
            parent.remove(ser)
        extra_xCharts = [
            xChart for xChart in chartSpace.chart.plotArea.iter_plots()
            if len(list(xChart.iter_sers())) == 0
        ]
        for xChart in extra_xCharts:
            parent = xChart.getparent()
            parent.remove(xChart)
