# encoding: utf-8

"""
lxml custom element classes for chart axis-related XML elements.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..xmlchemy import BaseOxmlElement, ZeroOrOne


class BaseAxisElement(BaseOxmlElement):
    """
    Base class for catAx, valAx, and perhaps other axis elements.
    """
    delete = ZeroOrOne('c:delete', successors=('c:axPos',))


class CT_CatAx(BaseAxisElement):
    """
    ``<c:catAx>`` element, defining a category axis.
    """


class CT_ValAx(BaseAxisElement):
    """
    ``<c:valAx>`` element, defining a category axis.
    """
