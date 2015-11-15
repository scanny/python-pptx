.. _MsoColorType:

``MSO_COLOR_TYPE``
==================

Specifies the color specification scheme

Example::

    from pptx.enum.dml import MSO_COLOR_TYPE

    assert shape.fill.fore_color.type == MSO_COLOR_TYPE.SCHEME

----

RGB
    Color is specified by an |RGBColor| value

SCHEME
    Color is one of the preset theme colors

HSL
    Color is specified using Hue, Saturation, and Luminosity values

PRESET
    Color is specified using a named built-in color

SCRGB
    Color is an scRGB color, a wide color gamut RGB color space

SYSTEM
    Color is one specified by the operating system, such as the window
    background color.
