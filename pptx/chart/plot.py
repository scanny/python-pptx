# encoding: utf-8

"""
Plot-related objects. A plot is known as a chart group in the MS API. A chart
can have more than one plot overlayed on each other, such as a line plot
layered over a bar plot.
"""

from __future__ import absolute_import, print_function, unicode_literals

from collections import Sequence

from ..enum.chart import XL_CHART_TYPE as XL
from ..oxml.ns import qn
from ..oxml.simpletypes import ST_BarDir, ST_Grouping
from .series import SeriesFactory
from ..text import Font
from ..util import lazyproperty


class Plot(object):
    """
    A distinct plot that appears in the plot area of a chart. A chart may
    have more than one plot, in which case they appear as superimposed
    layers, such as a line plot appearing on top of a bar chart.
    """
    def __init__(self, plot_elm):
        super(Plot, self).__init__()
        self._element = plot_elm

    @property
    def data_labels(self):
        """
        |DataLabels| instance providing properties and methods on the
        collection of data labels associated with this plot.
        """
        dLbls = self._element.dLbls
        if dLbls is None:
            raise ValueError(
                'plot has no data labels, set has_data_labels = True first'
            )
        return DataLabels(dLbls)

    @property
    def has_data_labels(self):
        """
        Read/write boolean, |True| if the series has data labels. Assigning
        |True| causes data labels to be added to the plot. Assigning False
        removes any existing data labels.
        """
        return self._element.dLbls is not None

    @has_data_labels.setter
    def has_data_labels(self, value):
        """
        Add, remove, or leave alone the ``<c:dLbls>`` child element depending
        on current state and assigned *value*. If *value* is |True| and no
        ``<c:dLbls>`` element is present, a new default element is added with
        default child elements and settings. When |False|, any existing dLbls
        element is removed.
        """
        if bool(value) is False:
            self._element._remove_dLbls()
        else:
            if self._element.dLbls is None:
                self._element._add_dLbls()

    @lazyproperty
    def series(self):
        """
        The |SeriesCollection| instance containing the sequence of data
        series in this plot.
        """
        return SeriesCollection(self._element)


class AreaPlot(Plot):
    """
    An area plot.
    """


class Area3DPlot(Plot):
    """
    A 3-dimensional area plot.
    """


class BarPlot(Plot):
    """
    A bar chart-style plot.
    """
    @property
    def gap_width(self):
        """
        Width of gap between bar(s) of each category, as an integer
        percentage of the bar width. The default width for a new bar chart is
        150%.
        """
        gapWidth = self._element.gapWidth
        if gapWidth is None:
            return 150
        return gapWidth.val

    @gap_width.setter
    def gap_width(self, value):
        gapWidth = self._element.get_or_add_gapWidth()
        gapWidth.val = value


class DataLabels(object):
    """
    Collection of data labels associated with a plot, and perhaps with
    a series, not sure about that yet.
    """
    def __init__(self, dLbls):
        super(DataLabels, self).__init__()
        self._element = dLbls

    @lazyproperty
    def font(self):
        """
        The |Font| object that provides access to the text properties for
        these data labels, such as bold, italic, etc. Accessing this property
        Causes a ``<c:txPr>`` child element to be added if not already
        present.
        """
        defRPr = self._element.defRPr
        font = Font(defRPr)
        return font

    @property
    def number_format(self):
        """
        Read/write string specifying the format for the numbers on this set
        of data labels. Returns 'General' if no number format has been set.
        Note that this format string has no effect on rendered data labels
        when :meth:`number_format_is_linked` is |True|. Assigning a format
        string to this property automatically sets
        :meth:`number_format_is_linked` to |False|.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return 'General'
        return numFmt.formatCode

    @number_format.setter
    def number_format(self, value):
        self._element.get_or_add_numFmt().formatCode = value
        self.number_format_is_linked = False

    @property
    def number_format_is_linked(self):
        """
        Read/write boolean specifying whether number formatting should be
        taken from the source spreadsheet rather than the value of
        :meth:`number_format`.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return True
        souceLinked = numFmt.sourceLinked
        if souceLinked is None:
            return True
        return numFmt.sourceLinked

    @number_format_is_linked.setter
    def number_format_is_linked(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.sourceLinked = value


class LinePlot(Plot):
    """
    A line chart-style plot.
    """


class PiePlot(Plot):
    """
    A pie chart-style plot.
    """


def PlotFactory(plot_elm):
    """
    Return an instance of the appropriate subclass of Plot based on the
    tagname of *plot_elm*.
    """
    try:
        PlotCls = {
            qn('c:areaChart'):   AreaPlot,
            qn('c:area3DChart'): Area3DPlot,
            qn('c:barChart'):    BarPlot,
            qn('c:lineChart'):   LinePlot,
            qn('c:pieChart'):    PiePlot,
        }[plot_elm.tag]
    except KeyError:
        raise ValueError('unsupported plot type %s' % plot_elm.tag)

    return PlotCls(plot_elm)


class PlotTypeInspector(object):
    """
    "One-shot" service object that knows how to identify the type of a plot
    as a member of the XL_CHART_TYPE enumeration.
    """
    @classmethod
    def chart_type(cls, plot):
        """
        Return the member of :ref:`XlChartType` that corresponds to the chart
        type of *plot*.
        """
        try:
            chart_type_method = {
                'AreaPlot':   cls._differentiate_area_chart_type,
                'Area3DPlot': cls._differentiate_area_3d_chart_type,
                'BarPlot':    cls._differentiate_bar_chart_type,
                'LinePlot':   cls._differentiate_line_chart_type,
            }[plot.__class__.__name__]
        except KeyError:
            raise NotImplementedError(
                "chart_type() not implemented for %s" %
                plot.__class__.__name__
            )
        return chart_type_method(plot)

    @classmethod
    def _differentiate_area_3d_chart_type(cls, plot):
        return {
            ST_Grouping.STANDARD:        XL.THREE_D_AREA,
            ST_Grouping.STACKED:         XL.THREE_D_AREA_STACKED,
            ST_Grouping.PERCENT_STACKED: XL.THREE_D_AREA_STACKED_100,
        }[plot._element.grouping_val]

    @classmethod
    def _differentiate_area_chart_type(cls, plot):
        return {
            ST_Grouping.STANDARD:        XL.AREA,
            ST_Grouping.STACKED:         XL.AREA_STACKED,
            ST_Grouping.PERCENT_STACKED: XL.AREA_STACKED_100,
        }[plot._element.grouping_val]

    @classmethod
    def _differentiate_bar_chart_type(cls, plot):
        barChart = plot._element
        if barChart.barDir.val == ST_BarDir.BAR:
            return {
                ST_Grouping.CLUSTERED:       XL.BAR_CLUSTERED,
                ST_Grouping.STACKED:         XL.BAR_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.BAR_STACKED_100,
            }[barChart.grouping_val]
        if barChart.barDir.val == ST_BarDir.COL:
            return {
                ST_Grouping.CLUSTERED:       XL.COLUMN_CLUSTERED,
                ST_Grouping.STACKED:         XL.COLUMN_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.COLUMN_STACKED_100,
            }[barChart.grouping_val]
        raise ValueError(
            "invalid barChart.barDir value '%s'" % barChart.barDir.val
        )

    @classmethod
    def _differentiate_line_chart_type(cls, plot):
        lineChart = plot._element
        if cls._has_line_markers(lineChart):
            return {
                ST_Grouping.STANDARD:        XL.LINE_MARKERS,
                ST_Grouping.STACKED:         XL.LINE_MARKERS_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.LINE_MARKERS_STACKED_100,
            }[plot._element.grouping_val]
        else:
            return {
                ST_Grouping.STANDARD:        XL.LINE,
                ST_Grouping.STACKED:         XL.LINE_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.LINE_STACKED_100,
            }[plot._element.grouping_val]

    @classmethod
    def _has_line_markers(cls, lineChart):
        none_marker_symbols = lineChart.xpath(
            './c:ser/c:marker/c:symbol[@val="none"]'
        )
        if none_marker_symbols:
            return False
        return True


class SeriesCollection(Sequence):
    """
    The sequence of series in a plot.
    """
    def __init__(self, plot_elm):
        super(SeriesCollection, self).__init__()
        self._plot_elm = plot_elm

    def __getitem__(self, index):
        ser = self._ser_lst[index]
        return SeriesFactory(self._plot_elm, ser)

    def __len__(self):
        return len(self._ser_lst)

    @property
    def _ser_lst(self):
        return [ser for ser in self._plot_elm.iter_sers()]
