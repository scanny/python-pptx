# encoding: utf-8

"""
lxml custom element classes for ext-related XML elements.
"""

from __future__ import absolute_import

from .xmlchemy import ZeroOrMore


class CT_OfficeArtExtensionList(BaseOxmlElement):
    """
    Custom element class for <a:CT_OfficeArtExtensionList> elements.
    """
    ext = ZeroOrMore("a:ext")
    
    def add_extension(self, uri):
        ext = self._add_ext()
        ext.uri = uri
        hyperlinkColor = ext.get_or_add_hyperlinkColor()
        return ext
