# encoding: utf-8

"""
Core properties part, corresponds to ``/docProps/core.xml`` part in package.
"""

from datetime import datetime

from pptx.opc_constants import CONTENT_TYPE as CT
from pptx.oxml import CT_CoreProperties
from pptx.parts.part import _BasePart


class _CoreProperties(_BasePart):
    """
    Corresponds to part named ``/docProps/core.xml``, containing the core
    document properties for this document package.
    """
    _propnames = (
        'author', 'category', 'comments', 'content_status', 'created',
        'identifier', 'keywords', 'language', 'last_modified_by',
        'last_printed', 'modified', 'revision', 'subject', 'title', 'version'
    )

    def __init__(self, partname=None):
        super(_CoreProperties, self).__init__(CT.OPC_CORE_PROPERTIES, partname)

    @classmethod
    def _default(cls):
        core_props = _CoreProperties('/docProps/core.xml')
        core_props._element = CT_CoreProperties.new_coreProperties()
        core_props.title = 'PowerPoint Presentation'
        core_props.last_modified_by = 'python-pptx'
        core_props.revision = 1
        core_props.modified = datetime.utcnow()
        return core_props

    def __getattribute__(self, name):
        """
        Intercept attribute access to generalize property getters.
        """
        if name in _CoreProperties._propnames:
            return getattr(self._element, name)
        else:
            return super(_CoreProperties, self).__getattribute__(name)

    def __setattr__(self, name, value):
        """
        Intercept attribute assignment to generalize assignment to properties
        """
        if name in _CoreProperties._propnames:
            setattr(self._element, name, value)
        else:
            super(_CoreProperties, self).__setattr__(name, value)
