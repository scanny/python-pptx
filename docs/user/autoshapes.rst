
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

    shapes = slide.shapes
    left = top = width = height = Inches(1.0)
    shape = shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )

See the :ref:`MsoAutoShapeType` enumeration page for a list of all 182 auto
shape types.

Reading ``shape.auto_shape_type`` returns the member of
:ref:`MsoAutoShapeType` corresponding to the shape's ``a:prstGeom/@prst``
value. For a preset that has no matching enum member (for example, a ``prst``
value written by a newer version of PowerPoint), ``shape.auto_shape_type``
returns ``None`` instead of raising.


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


Line-end arrows
~~~~~~~~~~~~~~~

The outline of a shape (or a connector line) can carry an arrowhead decoration
at either or both ends. The decoration at the head (begin) of the line is
exposed as ``line.begin_arrow`` and the decoration at the tail (end) of the
line is exposed as ``line.end_arrow``. Each exposes three read/write
sub-properties — ``type``, ``width``, and ``length``::

    >>> from pptx.enum.dml import (
    ...     MSO_LINE_END_TYPE, MSO_LINE_END_WIDTH, MSO_LINE_END_LENGTH,
    ... )
    >>> line = shape.line
    >>> line.end_arrow.type = MSO_LINE_END_TYPE.TRIANGLE
    >>> line.end_arrow.width = MSO_LINE_END_WIDTH.LARGE
    >>> line.end_arrow.length = MSO_LINE_END_LENGTH.SMALL
    >>> line.begin_arrow.type = MSO_LINE_END_TYPE.OVAL

Assigning ``None`` to any sub-property removes that attribute from the line-end
element. When all three attributes are cleared, the line-end element itself is
removed.

To author the PowerPoint Insert-Shapes variants *Connector with Arrow*,
*Connector with Double Arrow*, and their curved/elbow siblings, combine any
``MSO_CONNECTOR`` geometry with ``begin_arrow`` / ``end_arrow`` on the
connector's line. The three ``MSO_CONNECTOR`` members
(``STRAIGHT`` / ``ELBOW`` / ``CURVE``) supply the geometry; the arrow-head
decoration is an outline property, not a separate connector type::

    >>> from pptx.enum.shapes import MSO_CONNECTOR
    >>> from pptx.enum.dml import MSO_LINE_END_TYPE, MSO_LINE_END_WIDTH
    >>> conn = slide.shapes.add_connector(
    ...     MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(1), Inches(4), Inches(1)
    ... )
    >>> conn.line.end_arrow.type = MSO_LINE_END_TYPE.TRIANGLE  # "Connector with Arrow"
    >>> conn.line.begin_arrow.type = MSO_LINE_END_TYPE.TRIANGLE  # "Double Arrow"
    >>> conn.line.end_arrow.width = MSO_LINE_END_WIDTH.LARGE


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


Reading path geometry from a shape
----------------------------------

The ``Shape.path_geometry`` property returns a ``PathGeometry`` — an ordered sequence of
``Path`` objects, one per ``<a:path>`` contour in the shape's rendered outline. Each
``Path`` is an ordered sequence of drawing-operation value objects (``MoveTo``, ``LineTo``,
``CubicBezierTo``, ``QuadBezierTo``, ``ArcTo``, ``Close``) whose ``(x, y)`` coordinates
are ``Length`` instances expressed in shape-local EMU. The shape's bounding box runs from
``(0, 0)`` in the top-left to ``(shape.width, shape.height)`` in the bottom-right::

    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(2), Inches(1))
    geom = shape.path_geometry
    assert len(geom) == 1           # one <a:path> contour
    path = geom[0]
    for op in path:
        print(op)                    # MoveTo, LineTo, LineTo, LineTo, Close

Two geometry kinds are supported:

* **Custom (freeform) geometry.** Shapes created with
  :meth:`~pptx.shapes.shapetree.BaseGroupShapes.build_freeform` have the coordinates of each
  drawing op reported verbatim from the stored ``a:custGeom/a:pathLst`` subtree.
* **Static preset geometry.** A subset of the ECMA-376 preset shapes have path definitions
  that do not depend on any adjustment values. For these, the path is resolved against the
  shape's current width and height and returned as a ``PathGeometry``. The supported preset
  ``prst`` identifiers are ``actionButtonBlank``, ``chartPlus``, ``chartStar``, ``chartX``,
  ``flowChartInternalStorage``, ``flowChartManualInput``, ``flowChartProcess``,
  ``flowChartPunchedCard``, ``lineInv``, and ``rect``. In ``MSO_SHAPE`` enumeration terms
  these correspond to members such as ``RECTANGLE``, ``FLOWCHART_PROCESS``,
  ``FLOWCHART_CARD``, ``CHART_PLUS``, ``CHART_STAR``, ``CHART_X`` and the various
  ``ACTION_BUTTON_*`` types.

For dynamic presets — any preset whose path depends on one or more adjustment values (e.g.
``ROUNDED_RECTANGLE``, ``CHEVRON``) — ``Shape.path_geometry`` returns ``None``. The
DrawingML formula evaluator needed to resolve these is not yet implemented.


Reading and writing a shape's theme-style refs
----------------------------------------------

PowerPoint's ribbon "Shape Styles" gallery is encoded in the shape's
``<p:style>`` subtree — four references into the slide master's
``<a:fmtScheme>`` / ``<a:fontScheme>``:

* ``<a:lnRef idx="…">`` — line style index.
* ``<a:fillRef idx="…">`` — fill style index.
* ``<a:effectRef idx="…">`` — effect style index.
* ``<a:fontRef idx="major|minor|none">`` — font-collection key.

``BaseShape.theme_style_refs`` exposes these as a
:class:`~pptx.shapes.base.ThemeStyleRefs` named-tuple. A newly-added
auto-shape carries the defaults PowerPoint emits for a subtle-effect
preset::

    from pptx import Presentation
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.shapes.base import ThemeStyleRefs

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
    )

    print(shape.theme_style_refs)
    # ThemeStyleRefs(line_ref=1, fill_ref=3, effect_ref=2, font_ref='minor')

Assign a :class:`~pptx.shapes.base.ThemeStyleRefs` (or any matching 4-tuple)
to select a different preset from the theme's format-scheme, or assign
|None| to drop the ``<p:style>`` entirely (leaving the shape with only
its explicit ``<p:spPr>`` properties)::

    shape.theme_style_refs = ThemeStyleRefs(2, 4, 1, "major")
    shape.theme_style_refs = None          # -- clear the preset --

Index ``0`` means "no style from the matrix" (the "No Fill" / "No
Outline" choice in PowerPoint's UI); indices ``1..N`` select the Nth
line / fill / effect style from the slide master theme. The ``font_ref``
string must be one of ``"major"``, ``"minor"``, or ``"none"``. The
actual visual result of each index depends on the theme.

Read-only inspection works for every simple shape that has a ``<p:style>``
(autoshapes, text-boxes, connectors, pictures); writing is also supported
on these four shape kinds. Graphic-frame wrappers (charts, tables,
SmartArt) and group shapes do not carry a ``<p:style>`` — reading returns
|None| and assigning raises :class:`ValueError`.

Preserving theme colors with ``BaseShape.theme_style``
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Each of the four refs on ``<p:style>`` also carries a nested
``<a:schemeClr val="…"/>`` that selects *which* theme color the referenced
style is tinted with (``"accent1"``..``"accent6"``, ``"bg1"``/``"bg2"``,
``"tx1"``/``"tx2"``, etc. — the full ``ST_SchemeColorVal`` set). PowerPoint
writes ``accent1`` for every ref when you apply a Shape-Styles gallery
preset, but a user can re-color a shape via the ribbon's Color menu and
the resulting XML will then carry something like ``<a:schemeClr
val="accent4"/>`` on each ref.

``BaseShape.theme_style`` exposes an eight-field
:class:`~pptx.shapes.base.ThemeStyle` view combining the four ``idx``
values and the four ``schemeClr`` colors — use it instead of
:attr:`~pptx.shapes.base.BaseShape.theme_style_refs` when you need to
preserve or change the color assignments during a round-trip::

    from pptx.shapes.base import ThemeStyle

    ts = shape.theme_style
    # ThemeStyle(line_idx=1, line_color='accent1', fill_idx=3,
    #            fill_color='accent1', effect_idx=2, effect_color='accent1',
    #            font_idx='minor', font_color='lt1')

    shape.theme_style = ThemeStyle(
        line_idx=1, line_color="accent2",
        fill_idx=3, fill_color="accent2",
        effect_idx=2, effect_color="accent2",
        font_idx="minor", font_color="lt1",
    )

Writing via ``theme_style_refs`` is still supported and preserves the
per-ref ``schemeClr`` values if a ``<p:style>`` already exists on the
shape. Assigning ``None`` via either accessor removes the ``<p:style>``
entirely.
