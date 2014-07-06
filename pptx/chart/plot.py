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


def PlotFactory(plot_elm):
    """
    Return an instance of the appropriate subclass of Plot based on the
    tagname of *plot_elm*.
    """
