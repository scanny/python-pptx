.. _MsoFillType:

``MSO_FILL_TYPE``
=================

Specifies the type of bitmap used for the fill of a shape.

Alias: ``MSO_FILL``

Example::

    from pptx.enum.dml import MSO_FILL

    assert shape.fill.type == MSO_FILL.SOLID

----

BACKGROUND
    The shape is transparent, such that whatever is behind the shape shows
    through. Often this is the slide background, but if a visible shape is
    behind, that will show through.

GRADIENT
    Shape is filled with a gradient

GROUP
    Shape is part of a group and should inherit the fill properties of the
    group.

PATTERNED
    Shape is filled with a pattern

PICTURE
    Shape is filled with a bitmapped image

SOLID
    Shape is filled with a solid color

TEXTURED
    Shape is filled with a texture
