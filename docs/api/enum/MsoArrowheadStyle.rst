.. _MsoArrowheadStyle:

``MSO_ARROWHEAD_STYLE``
=======================

Specifies the style of the arrowhead at the end of a line.

    Alias: ``MSO_ARROWHEAD``

    Example::

        from pptx.enum.dml import MSO_ARROWHEAD

        shape.line.head_end == MSO_ARROWHEAD.TRIANGLE

----

DIAMOND
    Diamond-shaped.

NONE
    No arrowhead.

OPEN
    Open.

OVAL
    Oval-shaped.

STEALTH
    Stealth-shaped.

TRIANGLE
    Triangular.

MIXED
    Indicates a combination of the other states.
