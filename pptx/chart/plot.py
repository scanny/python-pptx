# encoding: utf-8

"""
Plot-related objects. A plot is known as a chart group in the MS API. A chart
can have more than one plot overlayed on each other, such as a line plot
layered over a bar plot.
"""

from __future__ import absolute_import, print_function, unicode_literals


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
        Note that setting this format string has no effect on rendered data
        labels when :meth:`number_format_is_linked` is |True|.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return 'General'
        return numFmt.formatCode

    @number_format.setter
    def number_format(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.formatCode = value


def PlotFactory(plot_elm):
    """
    Return an instance of the appropriate subclass of Plot based on the
    tagname of *plot_elm*.
    """
