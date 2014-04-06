.. _MsoShapeType:

``MSO_SHAPE_TYPE``
==================

Specifies the type of a shape

Alias: ``MSO``

Example::

    from pptx.enum.shapes import MSO_SHAPE_TYPE

    assert shape.type == MSO_SHAPE_TYPE.PICTURE

----

AUTO_SHAPE
    AutoShape

CALLOUT
    Callout shape

CANVAS
    Drawing canvas

CHART
    Chart, e.g. pie chart, bar chart

COMMENT
    Comment

DIAGRAM
    Diagram

EMBEDDED_OLE_OBJECT
    Embedded OLE object

FORM_CONTROL
    Form control

FREEFORM
    Freeform

GROUP
    Group shape

IGX_GRAPHIC
    SmartArt graphic

INK
    Ink

INK_COMMENT
    Ink Comment

LINE
    Line

LINKED_OLE_OBJECT
    Linked OLE object

LINKED_PICTURE
    Linked picture

MEDIA
    Media

OLE_CONTROL_OBJECT
    OLE control object

PICTURE
    Picture

PLACEHOLDER
    Placeholder

SCRIPT_ANCHOR
    Script anchor

TABLE
    Table

TEXT_BOX
    Text box

TEXT_EFFECT
    Text effect

WEB_VIDEO
    Web video

MIXED
    Mixed shape types
