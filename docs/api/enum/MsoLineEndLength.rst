.. _MsoLineEndLength:

``MSO_LINE_END_LENGTH``
=======================

Specifies the length of a line-end decoration (arrowhead).

    Example::

        from pptx.enum.dml import MSO_LINE_END_LENGTH

        shape.line.end_arrow.length = MSO_LINE_END_LENGTH.LARGE

----

SMALL
    Small arrowhead length.

MEDIUM
    Medium arrowhead length.

LARGE
    Large arrowhead length.
