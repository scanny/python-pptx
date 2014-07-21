# encoding: utf-8

"""
Plot-related objects. A plot is known as a chart group in the MS API. A chart
can have more than one plot overlayed on each other, such as a line plot
layered over a bar plot.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..oxml.ns import qn


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


class BarPlot(Plot):
    """
    A bar chart-style plot.
    """


class DataLabels(object):
    """
    Collection of data labels associated with a plot, and perhaps with
    a series, not sure about that yet.
    """
    def __init__(self, dLbls):
        super(DataLabels, self).__init__()
        self._element = dLbls

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
    if plot_elm.tag == qn('c:barChart'):
        return BarPlot(plot_elm)
    if plot_elm.tag == qn('c:lineChart'):
        return LinePlot(plot_elm)
    if plot_elm.tag == qn('c:pieChart'):
        return PiePlot(plot_elm)
    raise ValueError('unsupported plot type %s' % plot_elm.tag)
