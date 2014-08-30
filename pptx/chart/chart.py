# encoding: utf-8

"""
Chart shape-related objects such as Chart.
"""

from __future__ import absolute_import, print_function, unicode_literals

from collections import Sequence

from .axis import CategoryAxis, ValueAxis
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
        :ref:`XlChartType` enumeration value specifying the type of this
        chart. If the chart has two plots, for example, a line plot overlayed
        on a bar plot, the type is reported for the first (back-most) plot.
        Read-only.
        """
        first_plot = self.plots[0]
        return PlotTypeInspector.chart_type(first_plot)

    @lazyproperty
    def plots(self):
        """
        The |Plots| instance containing the sequence of chart groups in
        this chart.
        """
        plotArea = self._chartSpace.chart.plotArea
        return Plots(plotArea)

    @lazyproperty
    def series(self):
        """
        The |SeriesCollection| object containing all the series in this
        chart.
        """
        return SeriesCollection(self._chartSpace)

    @property
    def value_axis(self):
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


class Plots(Sequence):
    """
    The sequence of plots in a chart, such as a bar plot or a line plot. Most
    charts have only a single plot. The concept is necessary when two chart
    types are displayed in a single set of axes, like a bar plot with
    a superimposed line plot.
    """
    def __init__(self, plotArea):
        super(Plots, self).__init__()
        self._plotArea = plotArea

    def __getitem__(self, index):
        plot_elms = [p for p in self._plotArea.iter_plots()]
        if isinstance(index, slice):
            return [PlotFactory(plot_elm) for plot_elm in plot_elms]
        else:
            plot_elm = plot_elms[index]
            return PlotFactory(plot_elm)

    def __len__(self):
        plot_elms = [p for p in self._plotArea.iter_plots()]
        return len(plot_elms)
