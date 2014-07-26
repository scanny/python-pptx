# encoding: utf-8

"""
Series-related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals


def SeriesFactory(plot_elm, ser):
    """
    Return an instance of the appropriate subclass of BaseSeries based on the
    tagname of *plot_elm*.
    """
    raise ValueError('unsupported series type %s' % plot_elm.tag)
