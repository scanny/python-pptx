.. _PpParagraphAlignment:

``PP_PARAGRAPH_ALIGNMENT``
==========================

Specifies the horizontal alignment for one or more paragraphs.

Alias: ``PP_ALIGN``

Example::

    from pptx.enum.text import PP_ALIGN

    shape.paragraphs[0].alignment = PP_ALIGN.CENTER

----

CENTER
    Center align

DISTRIBUTE
    Evenly distributes e.g. Japanese characters from left to right within a
    line

JUSTIFY
    Justified, i.e. each line both begins and ends at the margin with spacing
    between words adjusted such that the line exactly fills the width of the
    paragraph.

JUSTIFY_LOW
    Justify using a small amount of space between words.

LEFT
    Left aligned

RIGHT
    Right aligned

THAI_DISTRIBUTE
    Thai distributed

MIXED
    Return value only; indicates multiple paragraph alignments are present in
    a set of paragraphs.
