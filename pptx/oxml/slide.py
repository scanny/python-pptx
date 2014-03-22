# encoding: utf-8

"""
lxml custom element classes for slide master-related XML elements.
"""

from __future__ import absolute_import

from pptx.oxml.core import BaseOxmlElement


class CT_Slide(BaseOxmlElement):
    """
    ``<p:sld>`` element, root of a slide part
    """
