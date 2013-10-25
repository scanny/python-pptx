# encoding: utf-8

"""
Placeholder object.
"""

from pptx.spec import namespaces
from pptx.spec import PH_ORIENT_HORZ, PH_SZ_FULL, PH_TYPE_OBJ


# default namespace map for use in lxml calls
_nsmap = namespaces('a', 'r', 'p')


class _Placeholder(object):
    """
    Decorator (pattern) class for adding placeholder properties to a shape
    that contains a placeholder element, e.g. ``<p:ph>``.
    """
    def __new__(cls, shape):
        cls = type('PlaceholderDecorator', (_Placeholder, shape.__class__), {})
        return object.__new__(cls)

    def __init__(self, shape):
        self.__decorated = shape
        xpath = './*[1]/p:nvPr/p:ph'
        self.__ph = self._element.xpath(xpath, namespaces=_nsmap)[0]

    def __getattr__(self, name):
        """
        Called when *name* is not found in ``self`` or in class tree. In this
        case, delegate attribute lookup to decorated (it's probably in its
        instance namespace).
        """
        return getattr(self.__decorated, name)

    @property
    def type(self):
        """Placeholder type, e.g. PH_TYPE_CTRTITLE"""
        return self.__ph.get('type', PH_TYPE_OBJ)

    @property
    def orient(self):
        """Placeholder 'orient' attribute, e.g. PH_ORIENT_HORZ"""
        return self.__ph.get('orient', PH_ORIENT_HORZ)

    @property
    def sz(self):
        """Placeholder 'sz' attribute, e.g. PH_SZ_FULL"""
        return self.__ph.get('sz', PH_SZ_FULL)

    @property
    def idx(self):
        """Placeholder 'idx' attribute, e.g. '0'"""
        return int(self.__ph.get('idx', 0))
