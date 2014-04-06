# encoding: utf-8

"""
Enumerations used by shapes and related objects
"""

from __future__ import absolute_import

from . import (
    alias, Enumeration, EnumMember, ReturnValueOnlyEnumMember
)


@alias('MSO')
class MSO_SHAPE_TYPE(Enumeration):
    """
    Specifies the type of a shape

    Alias: ``MSO``

    Example::

        from pptx.enum.shapes import MSO_SHAPE_TYPE

        assert shape.type == MSO_SHAPE_TYPE.PICTURE
    """

    __ms_name__ = 'MsoShapeType'

    __url__ = (
        'http://msdn.microsoft.com/en-us/library/office/ff860759(v=office.15'
        ').aspx'
    )

    __members__ = (
        EnumMember(
            'AUTO_SHAPE', 1, 'AutoShape'
        ),
        EnumMember(
            'CALLOUT', 2, 'Callout shape'
        ),
        EnumMember(
            'CANVAS', 20, 'Drawing canvas'
        ),
        EnumMember(
            'CHART', 3, 'Chart, e.g. pie chart, bar chart'
        ),
        EnumMember(
            'COMMENT', 4, 'Comment'
        ),
        EnumMember(
            'DIAGRAM', 21, 'Diagram'
        ),
        EnumMember(
            'EMBEDDED_OLE_OBJECT', 7, 'Embedded OLE object'
        ),
        EnumMember(
            'FORM_CONTROL', 8, 'Form control'
        ),
        EnumMember(
            'FREEFORM', 5, 'Freeform'
        ),
        EnumMember(
            'GROUP', 6, 'Group shape'
        ),
        EnumMember(
            'IGX_GRAPHIC', 24, 'SmartArt graphic'
        ),
        EnumMember(
            'INK', 22, 'Ink'
        ),
        EnumMember(
            'INK_COMMENT', 23, 'Ink Comment'
        ),
        EnumMember(
            'LINE', 9, 'Line'
        ),
        EnumMember(
            'LINKED_OLE_OBJECT', 10, 'Linked OLE object'
        ),
        EnumMember(
            'LINKED_PICTURE', 11, 'Linked picture'
        ),
        EnumMember(
            'MEDIA', 16, 'Media'
        ),
        EnumMember(
            'OLE_CONTROL_OBJECT', 12, 'OLE control object'
        ),
        EnumMember(
            'PICTURE', 13, 'Picture'
        ),
        EnumMember(
            'PLACEHOLDER', 14, 'Placeholder'
        ),
        EnumMember(
            'SCRIPT_ANCHOR', 18, 'Script anchor'
        ),
        EnumMember(
            'TABLE', 19, 'Table'
        ),
        EnumMember(
            'TEXT_BOX', 17, 'Text box'
        ),
        EnumMember(
            'TEXT_EFFECT', 15, 'Text effect'
        ),
        EnumMember(
            'WEB_VIDEO', 26, 'Web video'
        ),
        ReturnValueOnlyEnumMember(
            'MIXED', -2, 'Mixed shape types'
        ),
    )
