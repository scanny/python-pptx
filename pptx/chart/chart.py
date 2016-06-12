# encoding: utf-8

"""
Chart shape-related objects such as Chart.
"""

from __future__ import absolute_import, print_function, unicode_literals

from collections import Sequence

from .axis import CategoryAxis, ValueAxis
from .legend import Legend
from .plot import PlotFactory, PlotTypeInspector
from .series import SeriesCollection
from ..shared import PartElementProxy
from ..util import lazyproperty
from .xmlwriter import SeriesXmlRewriterFactory


class Chart(PartElementProxy):
    """
    A chart object.
    """
    def __init__(self, chartSpace, chart_part):
        super(Chart, self).__init__(chartSpace, chart_part)
        self._chartSpace = chartSpace

    @property
    def category_axis(self):
        """
        The category axis of this chart. In the case of an XY or Bubble
        chart, this is the X axis. Raises |ValueError| if no category
        axis is defined (as is the case for a pie chart, for example).
        """
        catAx_lst = self._chartSpace.catAx_lst
        if catAx_lst:
            return CategoryAxis(catAx_lst[0])

        valAx_lst = self._chartSpace.valAx_lst
        if valAx_lst:
            return ValueAxis(valAx_lst[0])

        raise ValueError('chart has no category axis')

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
        rewriter = SeriesXmlRewriterFactory(self.chart_type, chart_data)
        rewriter.replace_series_data(self._chartSpace)
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
        valAx_lst = self._chartSpace.valAx_lst
        if not valAx_lst:
            raise ValueError('chart has no value axis')

        idx = 1 if len(valAx_lst) > 1 else 0
        return ValueAxis(valAx_lst[idx])

    @property
    def _workbook(self):
        """
        The |ChartWorkbook| object providing access to the Excel source data
        for this chart.
        """
        return self.part.chart_workbook


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
