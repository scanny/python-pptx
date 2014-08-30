# encoding: utf-8

"""
Composers for default chart XML for various chart types.
"""

from __future__ import absolute_import, print_function, unicode_literals


def ChartXmlWriter(chart_type, series_seq):
    """
    Factory function returning appropriate XML writer object for
    *chart_type*.
    """
    raise NotImplementedError
