.. _MsoTextUnderlineType:

``MSO_TEXT_UNDERLINE_TYPE``
===========================

Indicates the type of underline for text. Used with :attr:`.Font.underline`
to specify the style of text underlining.

Alias: ``MSO_UNDERLINE``

Example::

    from pptx.enum.text import MSO_UNDERLINE

    run.font.underline = MSO_UNDERLINE.DOUBLE_LINE

----

NONE
    Specifies no underline.

DASH_HEAVY_LINE
    Specifies a dash underline.

DASH_LINE
    Specifies a dash line underline.

DASH_LONG_HEAVY_LINE
    Specifies a long heavy line underline.

DASH_LONG_LINE
    Specifies a dashed long line underline.

DOT_DASH_HEAVY_LINE
    Specifies a dot dash heavy line underline.

DOT_DASH_LINE
    Specifies a dot dash line underline.

DOT_DOT_DASH_HEAVY_LINE
    Specifies a dot dot dash heavy line underline.

DOT_DOT_DASH_LINE
    Specifies a dot dot dash line underline.

DOTTED_HEAVY_LINE
    Specifies a dotted heavy line underline.

DOTTED_LINE
    Specifies a dotted line underline.

DOUBLE_LINE
    Specifies a double line underline.

HEAVY_LINE
    Specifies a heavy line underline.

SINGLE_LINE
    Specifies a single line underline.

WAVY_DOUBLE_LINE
    Specifies a wavy double line underline.

WAVY_HEAVY_LINE
    Specifies a wavy heavy line underline.

WAVY_LINE
    Specifies a wavy line underline.

WORDS
    Specifies underlining words.

MIXED
    Specifies a mixed of underline types.
