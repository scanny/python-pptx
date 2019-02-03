# encoding: utf-8

"""
Custom properties part, corresponds to ``/docProps/custom.xml`` part in package.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from datetime import datetime

from ..opc.constants import CONTENT_TYPE as CT
from ..opc.package import XmlPart
from ..opc.packuri import PackURI
from ..oxml import parse_xml
from ..oxml.customprops import CT_CustomProperties

from lxml import etree

NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


class CustomPropertiesPart(XmlPart):
    """
    Corresponds to part named ``/docProps/custom.xml``, containing the custom
    document properties for this document package.
    """
    @classmethod
    def default(cls, package):
        """
        Return a new |CustomPropertiesPart| object initialized with default
        values for its base properties.
        """
        custom_properties_part = cls._new(package)
        print(custom_properties_part)
        return custom_properties_part

    @classmethod
    def load(cls, partname, content_type, blob, package):
        element = parse_xml(blob)
        return cls(partname, content_type, element, package)

    @classmethod
    def _new(cls, package):
        partname = PackURI('/docProps/custom.xml')
        content_type = CT.OPC_CUSTOM_PROPERTIES
        customProperties = CT_CustomProperties.new()
        return CustomPropertiesPart(
            partname, content_type, customProperties, package
        )
    

    def __getitem__( self, item ):
        # print(etree.tostring(self._element, pretty_print=True))
        prop = self.lookup(item)
        if prop is not None :
            return prop[0].text

    def __setitem__( self, key, value ):
        prop = self.lookup(key)
        if prop is None :
            prop = etree.SubElement( self._element, "property" )
            elm = etree.SubElement(prop, '{%s}lpwstr' % NS_VT, nsmap = {'vt':NS_VT} )
            prop.set("name", key)
            prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            prop.set("pid", "%s" % str(len(self._element) + 1))
        else:
            elm = prop[0]
        elm.text = value
        # etree.tostring(prop, pretty_print=True)

    def lookup(self, item):
        for child in self._element :
            if child.get("name") == item :
                return child
        return None

    @property
    def keys(self):
        _keys = []
        for child in self._element:
            _keys.append(child.get("name"))
        return _keys
        
    def remove(self, item):
        for child in self._element :
            if child.get("name") == item:
                self._element.remove(child)
        
    def pop(self, item):
        prop = self.lookup(item)
        return_text = None
        if prop is not None:
            return prop[0].text
        self.remove(item)
