# encoding: utf-8

"""
lxml custom element classes for slide master-related XML elements.
"""

from __future__ import absolute_import

from ..xmlchemy import BaseOxmlElement, OneAndOnlyOne


class CT_SlideLayout(BaseOxmlElement):
    """
    ``<p:sldLayout>`` element, root of a slide layout part
    """
    cSld = OneAndOnlyOne('p:cSld')
