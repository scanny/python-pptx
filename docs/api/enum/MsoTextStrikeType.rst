.. _MsoTextStrikeType:

``MSO_TEXT_STRIKE_TYPE``
========================

Indicates the type of strikethrough for text. Used with
:attr:`.Font.strikethrough` to specify the style of text strikethrough.

Alias: ``MSO_STRIKE``

Example::

    from pptx.enum.text import MSO_STRIKE

    run.font.strikethrough = MSO_STRIKE.DOUBLE_LINE

----

NONE
    Specifies no strikethrough.

SINGLE_LINE
    Specifies a single line strikethrough.

DOUBLE_LINE
    Specifies a double line strikethrough.

MIXED
    Specifies a mix of strikethrough types (read-only).
