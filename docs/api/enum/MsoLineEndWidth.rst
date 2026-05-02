.. _MsoLineEndWidth:

``MSO_LINE_END_WIDTH``
======================

Specifies the width of a line-end decoration (arrowhead).

    Example::

        from pptx.enum.dml import MSO_LINE_END_WIDTH

        shape.line.end_arrow.width = MSO_LINE_END_WIDTH.MEDIUM

----

SMALL
    Small arrowhead width.

MEDIUM
    Medium arrowhead width.

LARGE
    Large arrowhead width.
