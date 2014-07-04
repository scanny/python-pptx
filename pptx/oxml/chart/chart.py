# encoding: utf-8

"""
lxml custom element classes for chart-related XML elements.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..simpletypes import XsdString
from ..xmlchemy import BaseOxmlElement, RequiredAttribute


class CT_Chart(BaseOxmlElement):
    """
    ``<c:chart>`` custom element class
    """
    rId = RequiredAttribute('r:id', XsdString)


class CT_ChartSpace(BaseOxmlElement):
    """
    ``<c:chartSpace>`` element class, the root element of a chart part.
    """
