
Working with AutoShapes
=======================

Auto shapes are regular shape shapes. Squares, circles, triangles, stars,
that sort of thing. There are 182 different auto shapes to choose from. 120
of these have adjustment "handles" you can use to change the shape, sometimes
dramatically.

Many shape types share a common set of properties. We'll introduce many of
them here because several of those shapes are just a specialized form of
AutoShape.


Adding an auto shape
--------------------

The following code adds a rounded rectangle shape, one inch square, and
positioned one inch from the top-left corner of the slide::

    from pptx.enum.shapes import MSO_SHAPE

    left = top = width = height = Inches(1.0)
    shape = shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )

See the :ref:`MsoAutoShapeType` enumeration page for a list of all 182 auto
shape types.


.. _`EMU`:

Understanding English Metric Units
----------------------------------

In the prior example we set the position and dimension values to the
expression ``Inches(1.0)``. What's that about?

Internally, PowerPoint stores length values in *English Metric Units* (EMU).
This term might be worth a quick Googling, but the short story is EMU is an
integer unit of length, 914400 to the inch. Most lengths in Office documents
are stored in EMU. 914400 has the great virtue that it is evenly divisible by
a great many common factors, allowing exact conversion between inches and
centimeters, for example. Being an integer, it can be represented exactly
across serializations and across platforms.

As you might imagine, working directly in EMU is inconvenient. To make it
easier, python-pptx provides a collection of value types to allow easy
specification and conversion into convenient units::

    >>> from pptx.util import Inches, Pt
    >>> length = Inches(1)
    >>> length
    914400
    >>> length.inches
    1.0
    >>> length.cm
    2.54
    >>> length.pt
    72.0
    >>> length = Pt(72)
    >>> length
    914400

More details are available in the :ref:`API documentation for pptx.util
<util>`


Shape position and dimensions
-----------------------------

All shapes have a position on their slide and have a size. In general,
position and size are specified when the shape is created. Position and size
can also be read from existing shapes and changed::

    >>> from pptx.enum.shapes import MSO_SHAPE
    >>> left = top = width = height = Inches(1.0)
    >>> shape = shapes.add_shape(
    >>>     MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    >>> )
    >>> shape.left, shape.top, shape.width, shape.height
    (914400, 914400, 914400, 914400)
    >>> shape.left.inches
    1.0
    >>> shape.left = Inches(2.0)
    >>> shape.left.inches
    2.0


Fill
----

AutoShapes have an outline around their outside edge. What appears within
that outline is called the shape's *fill*.

The most common type of fill is a solid color. A shape may also be filled
with a gradient, a picture, a pattern (like cross-hatching for example), or
may have no fill (transparent).

When a color is used, it may be specified as a specific RGB value or a color
from the theme palette.

Because there are so many options, the API for fill is a bit complex. This
code sets the fill of a shape to red::

    >>> fill = shape.fill
    >>> fill.solid()
    >>> fill.fore_color.rgb = RGBColor(255, 0, 0)

This sets it to the theme color that appears as 'Accent 1 - 25% Darker' in
the toolbar palette::

    >>> from pptx.enum.dml import MSO_THEME_COLOR
    >>> fill = shape.fill
    >>> fill.solid()
    >>> fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    >>> fill.fore_color.brightness = -0.25

This sets the shape fill to transparent, or 'No Fill' as it's called in the
PowerPoint UI::

    >>> shape.fill.background()

As you can see, the first step is to specify the desired fill type by calling
the corresponding method on fill. Doing so actually changes the properties
available on the fill object. For example, referencing ``.fore_color`` on a
fill object after calling its ``.background()`` method will raise an
exception::

    >>> fill = shape.fill
    >>> fill.solid()
    >>> fill.fore_color
    <pptx.dml.color.ColorFormat object at 0x10ce20910>
    >>> fill.background()
    >>> fill.fore_color
    Traceback (most recent call last):
      ...
    TypeError: a transparent (background) fill has no foreground color


Line
----

The outline of an AutoShape can also be formatted, including setting its
color, width, dash (solid, dashed, dotted, etc.), line style (single, double,
thick-thin, etc.), end cap, join type, and others. At the time of writing,
color and width can be set using python-pptx::

    >>> line = shape.line
    >>> line.color.rgb = RGBColor(255, 0, 0)
    >>> line.color.brightness = 0.5  # 50% lighter
    >>> line.width = Pt(2.5)

Theme colors can be used on lines too::

    >>> line.color.theme_color = MSO_THEME_COLOR.ACCENT_6

``Shape.line`` has the attribute ``.color``. This is essentially a shortcut
for::

    >>> line.fill.solid()
    >>> line.fill.fore_color

This makes sense for line formatting because a shape outline is most
frequently set to a solid color. Accessing the fill directly is required, for
example, to set the line to transparent::

    >>> line.fill.background()


Line width
~~~~~~~~~~

The shape outline also has a read/write width property::

    >>> line.width
    9525
    >>> line.width.pt
    0.75
    >>> line.width = Pt(2.0)
    >>> line.width.pt
    2.0


Adjusting an autoshape
----------------------

Many auto shapes have adjustments. In PowerPoint, these show up as little
yellow diamonds you can drag to change the look of the shape. They're a little
fiddly to work with via a program, but if you have the patience to get them
right, you can achieve some remarkable effects with great precision.


Shape Adjustment Concepts
~~~~~~~~~~~~~~~~~~~~~~~~~

There are a few concepts it's worthwhile to grasp before trying to do serious
work with adjustments.

First, adjustments are particular to a specific auto shape type. Each auto
shape has between zero and eight adjustments. What each of them does is
arbitrary and depends on the shape design.

Conceptually, adjustments are guides, in many ways like the light blue ones you
can align to in the PowerPoint UI and other drawing apps. These don't show, but
they operate in a similar way, each defining an x or y value that part of the
shape will align to, changing the proportions of the shape.

Adjustment values are large integers, each based on a nominal value of 100,000.
The effective value of an adjustment is proportional to the width or height of
the shape. So a value of 50,000 for an x-coordinate adjustment corresponds to
half the width of the shape; a value of 75,000 for a y-coordinate adjustment
corresponds to 3/4 of the shape height.

Adjustment values can be negative, generally indicating the coordinate is to
the left or above the top left corner (origin) of the shape. Values can also be
subject to limits, meaning their effective value cannot be outside a prescribed
range. In practice this corresponds to a point not being able to extend beyond
the left side of the shape, for example.

Spending some time fooling around with shape adjustments in PowerPoint is time
well spent to build an intuitive sense of how they behave. You also might want
to have ``opc-diag`` installed so you can look at the XML values that are
generated by different adjustments as a head start on developing your
adjustment code.


The following code formats a callout shape using its adjustments::

    callout_sp = shapes.add_shape(
        MSO_SHAPE.LINE_CALLOUT_2_ACCENT_BAR, left, top, width, height
    )

    # get the callout line coming out of the right place
    adjs = callout_sp.adjustments
    adjs[0] = 0.5   # vert pos of junction in margin line, 0 is top
    adjs[1] = 0.0   # horz pos of margin ln wrt shape width, 0 is left side
    adjs[2] = 0.5   # vert pos of elbow wrt margin line, 0 is top
    adjs[3] = -0.1  # horz pos of elbow wrt shape width, 0 is margin line
    adjs[4] = 3.0   # vert pos of line end wrt shape height, 0 is top
    a5 = adjs[3] - (adjs[4] - adjs[0]) * height/width
    adjs[5] = a5    # horz pos of elbow wrt shape width, 0 is margin line

    # rotate 45 degrees counter-clockwise
    callout_sp.rotation = -45.0
