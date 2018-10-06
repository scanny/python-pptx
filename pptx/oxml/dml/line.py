# encoding: utf-8

"""lxml custom element classes for DrawingML line-related XML elements."""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.enum.dml import MSO_ARROWHEAD_STYLE, MSO_LINE_DASH_STYLE
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute


class CT_HeadEndProperties(BaseOxmlElement):
    """`a:headEnd` custom element class"""
    type = OptionalAttribute('type', MSO_ARROWHEAD_STYLE)


class CT_TailEndProperties(BaseOxmlElement):
    """`a:tailEnd` custom element class"""
    type = OptionalAttribute('type', MSO_ARROWHEAD_STYLE)


class CT_PresetLineDashProperties(BaseOxmlElement):
    """`a:prstDash` custom element class"""

    val = OptionalAttribute("val", MSO_LINE_DASH_STYLE)
