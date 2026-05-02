
Working with text
=================

Auto shapes and table cells can contain text. Other shapes can't. Text is
always manipulated the same way, regardless of its container.

Text exists in a hierarchy of three levels:

* :attr:`.Shape.text_frame`
* :attr:`.TextFrame.paragraphs`
* :attr:`._Paragraph.runs`

All the text in a shape is contained in its *text frame*. A text frame has
vertical alignment, margins, wrapping and auto-fit behavior, a rotation angle,
some possible 3D visual features, and can be set to format its text into
multiple columns. It also contains a sequence of paragraphs, which always
contains at least one paragraph, even when empty.

A paragraph has line spacing, space before, space after, available bullet
formatting, tabs, outline/indentation level, and horizontal alignment.
A paragraph can be empty, but if it contains any text, that text is contained
in one or more runs.

A run exists to provide character level formatting, including font typeface,
size, and color, an optional hyperlink target URL, bold, italic, and underline
styles, strikethrough, kerning, and a few capitalization styles like all caps.

Let's run through these one by one. Only features available in the current
release are shown.


Accessing the text frame
------------------------

As mentioned, not all shapes have a text frame. So if you're not sure and you
don't want to catch the possible exception, you'll want to check before
attempting to access it::

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame
        # do things with the text frame
        ...


Accessing paragraphs
--------------------

A text frame always contains at least one paragraph. This causes the process
of getting multiple paragraphs into a shape to be a little clunkier than one
might like. Say for example you want a shape with three paragraphs::

    paragraph_strs = [
        'Egg, bacon, sausage and spam.',
        'Spam, bacon, sausage and spam.',
        'Spam, egg, spam, spam, bacon and spam.'
    ]

    text_frame = shape.text_frame
    text_frame.clear()  # remove any existing paragraphs, leaving one empty one

    p = text_frame.paragraphs[0]
    p.text = paragraph_strs[0]

    for para_str in paragraph_strs[1:]:
        p = text_frame.add_paragraph()
        p.text = para_str


Adding text
-----------

Only runs can actually contain text. Assigning a string to the ``.text``
attribute on a shape, text frame, or paragraph is a shortcut method for placing
text in a run contained by those objects. The following two snippets produce
the same result::

    shape.text = 'foobar'

    # is equivalent to ...

    text_frame = shape.text_frame
    text_frame.clear()
    p = text_frame.paragraphs[0]
    run = p.add_run()
    run.text = 'foobar'


Applying text frame-level formatting
------------------------------------

The following produces a shape with a single paragraph, a slightly wider bottom
than top margin (these default to 0.05"), no left margin, text aligned top, and
word wrapping turned off. In addition, the auto-size behavior is set to
adjust the width and height of the shape to fit its text. Note that vertical
alignment is set on the text frame. Horizontal alignment is set on each
paragraph::

    from pptx.util import Inches
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE

    text_frame = shape.text_frame
    text_frame.text = 'Spam, eggs, and spam'
    text_frame.margin_bottom = Inches(0.08)
    text_frame.margin_left = 0
    text_frame.vertical_anchor = MSO_ANCHOR.TOP
    text_frame.word_wrap = False
    text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

The possible values for ``TextFrame.auto_size`` and
``TextFrame.vertical_anchor`` are specified by the enumeration
:ref:`MsoAutoSize` and :ref:`MsoVerticalAnchor` respectively.


Applying paragraph formatting
-----------------------------

The following produces a shape containing three left-aligned paragraphs, the
second and third indented (like sub-bullets) under the first::

    from pptx.enum.text import PP_ALIGN

    paragraph_strs = [
        'Egg, bacon, sausage and spam.',
        'Spam, bacon, sausage and spam.',
        'Spam, egg, spam, spam, bacon and spam.'
    ]

    text_frame = shape.text_frame
    text_frame.clear()

    p = text_frame.paragraphs[0]
    p.text = paragraph_strs[0]
    p.alignment = PP_ALIGN.LEFT

    for para_str in paragraph_strs[1:]:
        p = text_frame.add_paragraph()
        p.text = para_str
        p.alignment = PP_ALIGN.LEFT
        p.level = 1


Applying character formatting
-----------------------------

Character level formatting is applied at the run level, using the ``.font``
attribute. The following formats a sentence in 18pt Calibri Bold and applies
the theme color Accent 1.

::

    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_THEME_COLOR
    from pptx.util import Pt

    text_frame = shape.text_frame
    text_frame.clear()  # not necessary for newly-created shape

    p = text_frame.paragraphs[0]
    run = p.add_run()
    run.text = 'Spam, eggs, and spam'

    font = run.font
    font.name = 'Calibri'
    font.size = Pt(18)
    font.bold = True
    font.italic = None  # cause value to be inherited from theme
    font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

If you prefer, you can set the font color to an absolute RGB value. Note that
this will not change color when the theme is changed::

    font.color.rgb = RGBColor(0xFF, 0x7F, 0x50)

A run can also be made into a hyperlink by providing a target URL::

    run.hyperlink.address = 'https://github.com/scanny/python-pptx'


.. _text-hyperlink-color-guide:

Coloring a hyperlinked run
--------------------------

By default, PowerPoint ignores the explicit color of a run that carries a
hyperlink and renders the run in the theme's hyperlink color (the color scheme
entry stored as ``a:hlink`` in the theme part). This is the behavior reported
in `issue #940`_.

Starting with python-pptx 1.1, there are two API paths that help you work
around this:

* ``Font.use_theme_hyperlink_color`` -- a tri-state flag on each run's font.
  Setting it to ``False`` records the intent that the run's explicit
  ``font.color`` be preferred over the theme color. python-pptx stores that
  intent as a marker extension under the run's ``a:hlinkClick`` element; the
  marker round-trips cleanly through save/reload and is ignored by PowerPoint,
  so it will not confuse the file. Setting the flag to ``True`` (or ``None``)
  clears the marker.

* ``ThemePart.theme.hlink_color`` / ``folHlink_color`` -- read and write the
  ``srgbClr`` entry of the theme's ``a:hlink`` and ``a:folHlink`` color
  scheme members. Reassigning these is what actually changes the color
  PowerPoint uses when rendering hyperlinked runs. Reach the theme via the
  relationship on a slide master::

      from pptx.dml.color import RGBColor
      from pptx.opc.constants import RELATIONSHIP_TYPE as RT

      theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
      theme = theme_part.theme
      theme.hlink_color = RGBColor(0xFF, 0x00, 0x00)

  The unvisited-hyperlink color is stored on each theme part. A presentation
  typically has one theme per slide master, so multi-master decks need to
  update each theme separately.

The recommended pattern for fully overriding a hyperlink's color is to
combine both: annotate the run (so the intent is visible and survives
round-trips) and update the theme so PowerPoint honors it::

    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    run.hyperlink.address = 'https://example.com'
    run.font.use_theme_hyperlink_color = False  # record the intent

    theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    theme_part.theme.hlink_color = RGBColor(0xFF, 0x00, 0x00)

.. _`issue #940`: https://github.com/scanny/python-pptx/issues/940
