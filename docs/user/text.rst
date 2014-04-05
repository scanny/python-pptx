
Working with text
=================

Auto shapes and table cells can contain text. Other shapes can't. Text is
always manipulated the same way, regardless of its container.

Text exists in a hierarchy of three levels:

* shape.textframe
* textframe.paragraphs
* paragraph.runs

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


Accessing the textframe
-----------------------

As mentioned, not all shapes have a textframe. So if you're not sure and you
don't want to catch the possible exception, you'll want to check before
attempting to access it::

    for shape in slide.shapes:
        if not shape.has_textframe:
            continue
        textframe = shape.textframe
        # do things with the textframe
        ...


Accessing paragraphs
--------------------

A textframe always contains at least one paragraph. This causes the process of
getting multiple paragraphs into a shape to be a little clunkier than one might
like. Say for example you want a shape with three paragraphs::

    paragraph_strs = [
        'Egg, bacon, sausage and spam.',
        'Spam, bacon, sausage and spam.',
        'Spam, egg, spam, spam, bacon and spam.'
    ]

    textframe = shape.textframe
    textframe.clear()  # remove any existing paragraphs, leaving one empty one

    p = textframe.paragraphs[0]
    p.text = paragraph_strs[0]

    for para_str in paragraph_strs[1:]:
        p = textframe.add_paragraph()
        p.text = para_str


Adding text
-----------

Only runs can actually contain text. Assigning a string to the ``.text``
attribute on a shape, text frame, or paragraph is a shortcut method for placing
text in a run contained by those objects. The following two snippets produce
the same result::

    shape.text = 'foobar'

    # is equivalent to ...

    textframe = shape.textframe
    textframe.clear()
    p = textframe.paragraphs[0]
    run = p.add_run()
    run.text = 'foobar'


Applying textframe-level formatting
-----------------------------------

The following produces a shape with a single paragraph, a slightly wider bottom
than top margin (these default to 0.05"), no left margin, text aligned top, and
word wrapping turned off. In addition, the auto-size behavior is set to
adjust the width and height of the shape to fit its text. Note that vertical
alignment is set on the textframe. Horizontal alignment is set on each
paragraph::

    from pptx.util import Inches
    from pptx.constants import MSO
    from pptx.enum.text import MSO_AUTO_SIZE

    textframe = shape.textframe
    textframe.text = 'Spam, eggs, and spam'
    textframe.margin_bottom = Inches(0.08)
    textframe.margin_left = 0
    textframe.vertical_anchor = MSO.ANCHOR_TOP
    textframe.word_wrap = False
    textframe.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

The possible values for ``TextFrame.auto_size`` are specified by the
enumeration :ref:`MsoAutoSize`.


Applying paragraph formatting
-----------------------------

The following produces a shape containing three left-aligned paragraphs, the
second and third indented (like sub-bullets) under the first::

    from pptx.constants import PP

    paragraph_strs = [
        'Egg, bacon, sausage and spam.',
        'Spam, bacon, sausage and spam.',
        'Spam, egg, spam, spam, bacon and spam.'
    ]

    textframe = shape.textframe
    textframe.clear()

    p = textframe.paragraphs[0]
    p.text = paragraph_strs[0]
    p.alignment = PP.ALIGN_LEFT

    for para_str in paragraph_strs[1:]:
        p = textframe.add_paragraph()
        p.text = para_str
        p.alignment = PP.ALIGN_LEFT
        p.level = 1


Applying character formatting
-----------------------------

Character level formatting is applied at the run level, using the ``.font``
attribute. The following formats a sentence in 18pt Calibri Bold and applies
the theme color Accent 1.

::

    from pptx.dml.color import RGBColor
    from pptx.enum import MSO_THEME_COLOR
    from pptx.util import Pt

    textframe = shape.textframe
    textframe.clear()  # not necessary for newly-created shape

    p = textframe.paragraphs[0]
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
