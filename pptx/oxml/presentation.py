# encoding: utf-8

"""
Custom element classes for presentation-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml import register_custom_element_class


class CT_Presentation(objectify.ObjectifiedElement):
    """
    ``<p:presentation>`` element, root of the Presentation part stored as
    ``/ppt/presentation.xml``.
    """


register_custom_element_class('p:presentation', CT_Presentation)
