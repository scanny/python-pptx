# encoding: utf-8

"""
Objects shared by pptx modules.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)


class ElementProxy(object):
    """
    Base class for lxml element proxy classes. An element proxy class is one
    whose primary responsibilities are fulfilled by manipulating the
    attributes and child elements of an XML element. They are the most common
    type of class in python-pptx other than custom element (oxml) classes.
    """

    __slots__ = ('_element',)

    def __init__(self, element):
        self._element = element
