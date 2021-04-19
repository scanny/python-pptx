# encoding: utf-8

"""
lxml custom element classes for ext-related XML elements.
"""

from __future__ import absolute_import

from .xmlchemy import ZeroOrMore, BaseOxmlElement


class CT_OfficeArtExtensionList(BaseOxmlElement):
    """
    Custom element class for <a:CT_OfficeArtExtensionList> elements.
    """
    ext = ZeroOrMore("a:ext")
    
    def add_extension(self):
        ext = self._add_ext()
        return ext
