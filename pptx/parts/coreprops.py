# encoding: utf-8

"""
Core properties part, corresponds to ``/docProps/core.xml`` part in package.
"""

from __future__ import absolute_import

from datetime import datetime

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml_bytes
from pptx.oxml.coreprops import CT_CoreProperties


class CoreProperties(Part):
    """
    Corresponds to part named ``/docProps/core.xml``, containing the core
    document properties for this document package.
    """
    _propnames = (
        'author', 'category', 'comments', 'content_status', 'created',
        'identifier', 'keywords', 'language', 'last_modified_by',
        'last_printed', 'modified', 'revision', 'subject', 'title', 'version'
    )

    def __init__(self, partname, content_type, core_props_elm):
        super(CoreProperties, self).__init__(
            partname, content_type, element=core_props_elm
        )

    @classmethod
    def default(cls):
        core_props = cls._new()
        core_props.title = 'PowerPoint Presentation'
        core_props.last_modified_by = 'python-pptx'
        core_props.revision = 1
        core_props.modified = datetime.utcnow()
        return core_props

    @classmethod
    def load(cls, partname, content_type, blob, package):
        core_props_elm = parse_xml_bytes(blob)
        core_props = cls(partname, content_type, core_props_elm)
        return core_props

    @classmethod
    def _new(cls):
        partname = PackURI('/docProps/core.xml')
        content_type = CT.OPC_CORE_PROPERTIES
        core_props_elm = CT_CoreProperties.new_coreProperties()
        return CoreProperties(partname, content_type, core_props_elm)

    def __getattribute__(self, name):
        """
        Intercept attribute access to generalize property getters.
        """
        if name in CoreProperties._propnames:
            return getattr(self._element, name)
        else:
            return super(CoreProperties, self).__getattribute__(name)

    def __setattr__(self, name, value):
        """
        Intercept attribute assignment to generalize assignment to properties
        """
        if name in CoreProperties._propnames:
            setattr(self._element, name, value)
        else:
            super(CoreProperties, self).__setattr__(name, value)
