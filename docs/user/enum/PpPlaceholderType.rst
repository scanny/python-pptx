.. _PpPlaceholderType:

``PP_PLACEHOLDER_TYPE``
=======================

Specifies one of the 18 distinct types of placeholder.

Alias: ``PP_PLACEHOLDER``

Example::

    from pptx.enum.shapes import PP_PLACEHOLDER

    placeholder = slide.placeholders[0]
    assert placeholder.type == PP_PLACEHOLDER.TITLE

----

BITMAP
    Bitmap

BODY
    Body

CENTER_TITLE
    Center Title

CHART
    Chart

DATE
    Date

FOOTER
    Footer

HEADER
    Header

MEDIA_CLIP
    Media Clip

OBJECT
    Object

ORG_CHART
    Organization Chart

PICTURE
    Picture

SLIDE_NUMBER
    Slide Number

SUBTITLE
    Subtitle

TABLE
    Table

TITLE
    Title

VERTICAL_BODY
    Vertical Body

VERTICAL_OBJECT
    Vertical Object

VERTICAL_TITLE
    Vertical Title

MIXED
    Return value only; multiple placeholders of differing types.
