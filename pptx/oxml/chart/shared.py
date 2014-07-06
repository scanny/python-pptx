# encoding: utf-8

"""
Shared oxml objects for charts.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..simpletypes import XsdBoolean, XsdDouble, XsdString
from ..xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute


class CT_Boolean(BaseOxmlElement):
    """
    Common complex type used for elements having a True/False value.
    """
    val = OptionalAttribute('val', XsdBoolean, default=True)


class CT_Double(BaseOxmlElement):
    """
    Used for floating point values.
    """
    val = RequiredAttribute('val', XsdDouble)


class CT_NumFmt(BaseOxmlElement):
    """
    ``<c:numFmt>`` element specifying the formatting for number labels on a
    tick mark or data point.
    """
    formatCode = RequiredAttribute('formatCode', XsdString)
    sourceLinked = OptionalAttribute('sourceLinked', XsdBoolean)
