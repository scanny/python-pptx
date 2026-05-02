
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


Paragraph bullets
-----------------

Each paragraph has a |_BulletFormat| proxy, obtained via ``paragraph.bullet``,
that controls the paragraph-level bullet. A paragraph may use a literal
character bullet (e.g. ``•``, ``-``, ``*``), an auto-numbered bullet
(Arabic, Roman, alphabetical, etc.), or it may explicitly suppress a bullet
that would otherwise be inherited from a layout or master.

With no explicit setting, the bullet is inherited from the paragraph's style
hierarchy (layout / master / theme)::

    paragraph = text_frame.paragraphs[0]
    paragraph.bullet.type  # -> None, the paragraph's bullet is inherited

To use a literal character as the bullet::

    paragraph.bullet.character('•')
    paragraph.bullet.type  # -> 'char'
    paragraph.bullet.char  # -> '•'

To use automatic numbering, choose a scheme from :ref:`PpAutoNumberScheme`::

    from pptx.enum.text import PP_AUTO_NUMBER

    paragraph.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)

An optional ``start_at`` argument controls the ordinal at which numbering
starts (1 by default, must be in range 1..32767)::

    paragraph.bullet.auto_number(PP_AUTO_NUMBER.ROMAN_UC_PERIOD, start_at=3)

To explicitly suppress any inherited bullet on a paragraph (produces an
``<a:buNone/>`` child)::

    paragraph.bullet.none()
    paragraph.bullet.type  # -> 'none'

To remove any explicit bullet setting so the paragraph once again inherits
from its style hierarchy::

    paragraph.bullet.clear()
    paragraph.bullet.type  # -> None

Each of ``.character()``, ``.auto_number()``, ``.none()``, and ``.clear()``
returns the |_BulletFormat| itself to support chaining, e.g.
``paragraph.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD, 2)``.


Bullet font, color, and size
----------------------------

The |_BulletFormat| proxy also exposes the bullet's font, color, and size.
These overrides are written alongside the bullet itself (e.g. an ``<a:buChar>``
or ``<a:buAutoNum>`` child) and, when not set, are inherited from the
paragraph's style hierarchy.

To set the bullet typeface — this is often required when the bullet character
is a glyph from a symbol font such as Wingdings::

    paragraph.bullet.character('§')  # section sign in Wingdings
    paragraph.bullet.font = 'Wingdings'
    paragraph.bullet.font  # -> 'Wingdings'

Assigning |None| clears any bullet-font override::

    paragraph.bullet.font = None
    paragraph.bullet.font  # -> None

To set the bullet size as a percentage of the surrounding text
(``<a:buSzPct>``)::

    paragraph.bullet.size_pct = 0.75  # 75% of text size
    paragraph.bullet.size_pct         # -> 0.75

The valid range is 0.25..4.0 (25% to 400% of the text size).

To set the bullet size as an absolute point value (``<a:buSzPts>``)::

    from pptx.util import Pt

    paragraph.bullet.size_points = Pt(14)
    paragraph.bullet.size_points.pt  # -> 14.0

Only one of ``size_pct`` or ``size_points`` may be in effect at a time;
assigning one automatically clears the other. Assigning |None| (or calling
:meth:`._BulletFormat.clear_size`) removes any bullet-size setting::

    paragraph.bullet.clear_size()

To set the bullet color, use the |ColorFormat| exposed as
``paragraph.bullet.color``. This mirrors the ``Font.color`` API::

    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_THEME_COLOR

    paragraph.bullet.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    # or
    paragraph.bullet.color.theme_color = MSO_THEME_COLOR.ACCENT_1

To remove an explicit bullet color, causing it to inherit from the style
hierarchy, call :meth:`._BulletFormat.clear_color`::

    paragraph.bullet.clear_color()

.. note::

    The "follow-text" variants of the bullet-font, bullet-color, and
    bullet-size overrides (``<a:buFontTx>``, ``<a:buClrTx>``, ``<a:buSzTx>``)
    are not yet exposed on |_BulletFormat|. These elements indicate that the
    bullet property should follow the text rather than inherit from the style
    hierarchy, and are relatively uncommon. Use the underlying XML via
    ``paragraph._pPr`` if you need them in the interim.
Setting a typeface for East-Asian or complex-script text
--------------------------------------------------------

A run in PowerPoint has three independent font slots: one for Latin-script
text, one for East-Asian (CJK: Chinese, Japanese, Korean) text, and one for
complex-script text (e.g. Arabic, Hebrew, Thai, Devanagari). PowerPoint
chooses the appropriate slot per-character based on the Unicode range of the
glyph being rendered, so mixed-script text in a single run can display
correctly.

|Font| exposes each of these slots as a separate property:

* :attr:`.Font.name` — Latin typeface (``a:latin``)
* :attr:`.Font.name_ea` — East-Asian typeface (``a:ea``)
* :attr:`.Font.name_cs` — complex-script typeface (``a:cs``)

::

    font = run.font
    font.name = 'Calibri'       # Latin script
    font.name_ea = 'MS Gothic'  # East-Asian / CJK
    font.name_cs = 'Arial'      # complex scripts

Each property reads and writes independently; assigning |None| removes only
that slot's override and restores inheritance from the theme for that
script::

    font.name_ea = None  # clears a:ea; a:latin and a:cs are preserved

Setting :attr:`.Font.name` alone has no effect on how CJK text is displayed,
which is the most common reason ``font.name`` appeared not to work for
East-Asian text — the typeface used for CJK characters is controlled by
:attr:`.Font.name_ea`, not :attr:`.Font.name`.
Strikethrough text
~~~~~~~~~~~~~~~~~~

The ``Font.strikethrough`` property mirrors ``Font.bold`` and ``Font.italic``:
it accepts ``True``, ``False``, or ``None`` (inherit). For a double-line
strikethrough, assign ``MSO_STRIKE.DOUBLE_LINE``::

    from pptx.enum.text import MSO_STRIKE

    run.font.strikethrough = True                    # single-line strikethrough
    run.font.strikethrough = MSO_STRIKE.DOUBLE_LINE  # double-line strikethrough
    run.font.strikethrough = False                   # explicitly no strikethrough
    run.font.strikethrough = None                    # inherit from style hierarchy
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
