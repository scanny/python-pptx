.. _MsoLineDashStyle:

``MSO_LINE_DASH_STYLE``
=======================

Specifies the dash style for a line.

    Alias: ``MSO_LINE``

    Example::

        from pptx.enum.dml import MSO_LINE

        shape.line.dash_style == MSO_LINE.DASH_DOT_DOT

----

DASH
    Line consists of dashes only.

DASH_DOT
    Line is a dash-dot pattern.

DASH_DOT_DOT
    Line is a dash-dot-dot pattern.

LONG_DASH
    Line consists of long dashes.

LONG_DASH_DOT
    Line is a long dash-dot pattern.

ROUND_DOT
    Line is made up of round dots.

SOLID
    Line is solid.

SQUARE_DOT
    Line is made up of square dots.

DASH_STYLE_MIXED
    Not supported.
