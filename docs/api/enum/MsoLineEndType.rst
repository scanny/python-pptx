.. _MsoLineEndType:

``MSO_LINE_END_TYPE``
=====================

Specifies the shape of a line-end decoration (arrowhead).

    Example::

        from pptx.enum.dml import MSO_LINE_END_TYPE

        shape.line.end_arrow.type = MSO_LINE_END_TYPE.TRIANGLE

----

NONE
    No line-end decoration.

TRIANGLE
    Triangular line end.

STEALTH
    Stealth line end (a filled arrow shape).

DIAMOND
    Diamond line end.

OVAL
    Oval line end.

ARROW
    Open arrow (two line strokes) line end.
