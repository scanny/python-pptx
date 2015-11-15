.. _MsoVerticalAnchor:

``MSO_VERTICAL_ANCHOR``
=======================

Specifies the vertical alignment of text in a text frame. Used with the
``.vertical_anchor`` property of the |TextFrame| object. Note that the
``vertical_anchor`` property can also have the value None, indicating
there is no directly specified vertical anchor setting and its effective
value is inherited from its placeholder if it has one or from the theme.
None may also be assigned to remove an explicitly specified vertical
anchor setting.

----

TOP
    Aligns text to top of text frame and inherits its value from its layout
    placeholder or theme.

MIDDLE
    Centers text vertically

BOTTOM
    Aligns text to bottom of text frame

MIXED
    Return value only; indicates a combination of the other states.
