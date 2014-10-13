# encoding: utf-8

"""
Plot-related objects. A plot is known as a chart group in the MS API. A chart
can have more than one plot overlayed on each other, such as a line plot
layered over a bar plot.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..enum.chart import XL_CHART_TYPE as XL
from ..oxml.ns import qn
from ..oxml.simpletypes import ST_BarDir, ST_Grouping
from .series import SeriesCollection
from ..text.text import Font
from ..util import lazyproperty


class Plot(object):
    """
    A distinct plot that appears in the plot area of a chart. A chart may
    have more than one plot, in which case they appear as superimposed
    layers, such as a line plot appearing on top of a bar chart.
    """
    def __init__(self, xChart, chart):
        super(Plot, self).__init__()
        self._element = xChart
        self._chart = chart

    @property
    def categories(self):
        """
        A tuple containing the category strings for this plot, in the order
        they appear on the category axis.
        """
        xChart = self._element
        category_pt_elms = xChart.cat_pts
        return tuple(pt.v.text for pt in category_pt_elms)

    @property
    def chart(self):
        """
        The |Chart| object containing this plot.
        """
        return self._chart

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
        A sequence of |Series| objects representing the series in this plot,
        in the order they appear in the plot.
        """
        return SeriesCollection(self._element)

    @property
    def vary_by_categories(self):
        """
        Read/write boolean value specifying whether to use a different color
        for each of the points in this plot. Only effective when there is
        a single series; PowerPoint automatically varies color by series when
        more than one series is present.
        """
        varyColors = self._element.varyColors
        if varyColors is None:
            return True
        return varyColors.val

    @vary_by_categories.setter
    def vary_by_categories(self, value):
        self._element.get_or_add_varyColors().val = bool(value)


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
        percentage of the bar width. The default value for a new bar chart is
        150, representing 150% or 1.5 times the width of a single bar.
        """
        gapWidth = self._element.gapWidth
        if gapWidth is None:
            return 150
        return gapWidth.val

    @gap_width.setter
    def gap_width(self, value):
        gapWidth = self._element.get_or_add_gapWidth()
        gapWidth.val = value

    @property
    def overlap(self):
        """
        Read/write int value in range -100..100 specifying a percentage of
        the bar width by which to overlap adjacent bars in a multi-series bar
        chart. Default is 0. A setting of -100 creates a gap of a full bar
        width and a setting of 100 causes all the bars in a category to be
        superimposed. A stacked bar plot has overlap of 100 by default.
        """
        overlap = self._element.overlap
        if overlap is None:
            return 0
        return overlap.val

    @overlap.setter
    def overlap(self, value):
        """
        Set the value of the ``<c:overlap>`` child element to *int_value*,
        or remove the overlap element if *int_value* is 0.
        """
        if value == 0:
            self._element._remove_overlap()
            return
        self._element.get_or_add_overlap().val = value


class DataLabels(object):
    """
    Collection of data labels associated with a plot, and perhaps with
    a series or data point, although the latter two are not yet implemented.
    """
    def __init__(self, dLbls):
        super(DataLabels, self).__init__()
        self._element = dLbls

    @lazyproperty
    def font(self):
        """
        The |Font| object that provides access to the text properties for
        these data labels, such as bold, italic, etc.
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

    @property
    def position(self):
        """
        Read/write :ref:`XlDataLabelPosition` enumeration value specifying
        the position of the data labels with respect to their data point, or
        |None| if no position is specified. Assigning |None| causes
        PowerPoint to choose the default position, which varies by chart
        type.
        """
        dLblPos = self._element.dLblPos
        if dLblPos is None:
            return None
        return dLblPos.val

    @position.setter
    def position(self, value):
        if value is None:
            self._element._remove_dLblPos()
            return
        self._element.get_or_add_dLblPos().val = value


class LinePlot(Plot):
    """
    A line chart-style plot.
    """


class PiePlot(Plot):
    """
    A pie chart-style plot.
    """


def PlotFactory(xChart, chart):
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
        }[xChart.tag]
    except KeyError:
        raise ValueError('unsupported plot type %s' % xChart.tag)

    return PlotCls(xChart, chart)


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
                'PiePlot':    cls._differentiate_pie_chart_type,
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

    @classmethod
    def _differentiate_pie_chart_type(cls, plot):
        pieChart = plot._element
        explosion = pieChart.xpath('./c:ser/c:explosion')
        return XL.PIE_EXPLODED if explosion else XL.PIE
