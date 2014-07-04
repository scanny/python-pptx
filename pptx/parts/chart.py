# encoding: utf-8

"""
Chart part objects, including Chart and Charts
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..opc.package import XmlPart
from ..util import lazyproperty


class ChartPart(XmlPart):
    """
    A chart part; corresponds to parts having partnames matching
    ppt/charts/chart[1-9][0-9]*.xml
    """
    @lazyproperty
    def chart(self):
        """
        The |Chart| object representing the chart in this part.
        """
