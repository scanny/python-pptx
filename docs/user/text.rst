
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
        text_frame.add_paragraph(para_str)

:meth:`.TextFrame.add_paragraph` accepts the paragraph text as its first
argument, saving a second ``.text = ...`` assignment (previously the
one-line-per-paragraph idiom was ``p = tf.add_paragraph(); p.text =
...``). Keyword-only ``bold``, ``italic``, ``size``, ``color``, and
``font_name`` arguments apply run-level formatting to the single run
``add_paragraph`` spawns for that text. The same shorthand is available
on :meth:`._Paragraph.add_run`. See issue #134::

    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    text_frame.add_paragraph(
        "Key insight",
        bold=True,
        size=Pt(18),
        color=RGBColor(0xC0, 0x00, 0x00),
        font_name="Calibri",
    )

``color`` accepts either an |RGBColor| or a member of
:class:`~pptx.enum.dml.MSO_THEME_COLOR`. Passing any run-level kwarg to
``add_paragraph`` without also supplying ``text`` raises ``ValueError`` —
there is no run to apply it to.


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
    p.add_run('foobar')

:meth:`._Paragraph.add_run` accepts the run's text as its first argument
and the same keyword-only ``bold`` / ``italic`` / ``size`` / ``color`` /
``font_name`` shortcuts as :meth:`.TextFrame.add_paragraph`, so a single
call can add a formatted run in one line instead of three. See issue #134.


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


Controlling text overflow (wrap, single-line, clip)
---------------------------------------------------

PowerPoint offers three distinct strategies for what happens when the
text authored into a shape is longer than fits. All three are exposed
through the existing :attr:`TextFrame.word_wrap` and
:attr:`TextFrame.auto_size` properties (see issue #623):

* **Wrap at the shape boundary** — ``text_frame.word_wrap = True``
  (emits ``a:bodyPr/@wrap="square"``). Text flows onto additional
  lines at word boundaries so everything fits horizontally.

* **Single line, overflow horizontally** —
  ``text_frame.word_wrap = False`` (emits ``a:bodyPr/@wrap="none"``).
  Text stays on one line and extends past the shape's right edge.

* **Clip anything that doesn't fit** — set
  ``text_frame.auto_size = MSO_AUTO_SIZE.NONE`` (emits
  ``<a:noAutofit/>``). PowerPoint neither resizes the shape nor scales
  the font; content that falls outside the shape is clipped at the
  shape boundary. This is the correct recipe when you want text to be
  cropped, as opposed to wrapped or overflowed.

Assigning ``None`` to ``word_wrap`` removes the attribute so the
effective value is inherited from the style hierarchy.


Rotating text inside a shape
----------------------------

Two text-frame properties control how text is oriented relative to the
shape that contains it:

* :attr:`.TextFrame.rotation` — a clockwise rotation in degrees applied to
  the text *inside* the shape (corresponds to ``a:bodyPr/@rot``). This is
  distinct from :attr:`.Shape.rotation`, which rotates the whole shape.
  The setter accepts an ``int`` or ``float``; negative values are
  normalized to the equivalent positive rotation in ``[0, 360)``.
  Assigning |None| (or ``0``) removes the attribute::

      from pptx.util import Inches

      tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
      tb.text_frame.text = "Rotated"
      tb.text_frame.rotation = 45   # text renders rotated 45° clockwise

* :attr:`.TextFrame.upright` — a boolean that keeps the text visually
  upright even when the enclosing shape is rotated (corresponds to
  ``a:bodyPr/@upright``). By default rotating a shape rotates its text
  along with it; setting ``upright = True`` is PowerPoint's "keep text
  upright" option on a rotated shape. Useful for callouts, labels on
  rotated diagrams, and legends that must remain readable::

      shape.rotation = 30            # rotate the whole shape 30°
      shape.text_frame.upright = True  # but keep the text itself upright

The two flags are independent — ``rotation`` controls explicit
text-inside-shape rotation, while ``upright`` opts the text out of the
surrounding shape's transform. See issues #133 and #485.


Inspecting the text-rendering rectangle
---------------------------------------

A shape's bounding box and its text-rendering rectangle are not the same.
PowerPoint renders text inside an "inset box" smaller than the shape by the
four ``TextFrame.margin_left`` / ``margin_top`` / ``margin_right`` /
``margin_bottom`` insets. If you need to pick a font size that will not
overflow a shape, or to lay out a diagram on top of the text region,
reading ``shape.text_frame_rect`` is the most direct way to get that
rectangle::

    from pptx.util import Inches

    shape = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
    rect = shape.text_frame_rect
    rect.left.inches    # 1.1   (left + margin_left)
    rect.top.inches     # 2.05  (top + margin_top)
    rect.width.inches   # 3.8   (width - margin_left - margin_right)
    rect.height.inches  # 0.9   (height - margin_top - margin_bottom)

The return value is a :class:`~pptx.text.text.TextFrameRect`, a
four-element ``(left, top, width, height)`` namedtuple of |Length|
values in English Metric Units, so the usual ``.inches`` / ``.cm`` /
``.pt`` / ``.emu`` readbacks are available and the tuple unpacks
directly::

    left, top, width, height = shape.text_frame_rect

Why the rectangle differs from ``(shape.left, shape.top, shape.width,
shape.height)``:

* Default ``TextFrame`` insets are 0.1" on the left and right and 0.05"
  on the top and bottom, so even a freshly-created text-box has a text
  rectangle 0.2" narrower and 0.1" shorter than its bounding box.
* Setting one of the margins (e.g. ``tf.margin_left = Inches(0.25)``)
  shrinks the rectangle on that side accordingly.
* The rectangle represents the *authored* inset box — it does not
  compensate for auto-fit (``auto_size`` / ``normAutofit``) which
  PowerPoint applies at render time to shrink the font, nor for text
  rotation (``TextFrame.rotation``), nor for group-transform compositing
  on a shape nested inside a group; use ``shape.effective_*`` if you
  need the post-composite slide-relative geometry.

``shape.text_frame_rect`` raises ``ValueError`` on a shape that has no
text frame (e.g. a connector or a picture without text), mirroring
``shape.placeholder_format`` — guard on ``shape.has_text_frame`` when
you are not sure.


Fitting text to a placeholder
-----------------------------

PowerPoint placeholders commonly have ``auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE``
set so that when text overflows the shape, PowerPoint reduces the rendered font-size
(and optionally line-spacing) to make it fit. PowerPoint records the reduction it
applied as the ``fontScale`` and ``lnSpcReduction`` attributes on the
``<a:normAutofit/>`` child of the text frame's ``a:bodyPr`` element.

PowerPoint re-computes these attributes when the user edits the text box, but will
not re-compute them when the file is saved programmatically and then opened for
display only. If you are generating a slide whose placeholder text you know will
overflow, you can emit the autofit hints directly so PowerPoint renders the text at
the scaled size on first display::

    text_frame = slide.placeholders[1].text_frame
    text_frame.text = "Long text that will not fit at the authored size..."
    text_frame.font_scale = 85.0          # 85% of authored font size
    text_frame.line_space_reduction = 10.0  # reduce line spacing by 10%

Both ``font_scale`` (range 1.0..100.0, default 100.0) and
``line_space_reduction`` (range 0.0..100.0, default 0.0) are percent values.
Assigning either one ensures the text frame has an ``a:normAutofit`` child.

For text boxes whose font is available locally, :meth:`.TextFrame.fit_text`
computes a best-fit point size and applies it directly to each run (so no
``normAutofit`` hints are required). :meth:`.TextFrame.fit_text` also works on
placeholder text frames; the placeholder's effective width and height are
obtained from its slide layout when not overridden on the slide.

.. note::
   ``fit_text`` needs to measure rendered glyph widths, which requires a
   real TrueType / OpenType file on disk. By default it searches the
   operating system's well-known font directories (macOS and Windows
   only) for a file matching ``font_family`` / ``bold`` / ``italic``. On
   Linux — or on any system where the requested typeface is not
   installed — pass an explicit ``font_file`` argument pointing at a
   ``.ttf`` or ``.otf`` file bundled with your application, or set
   ``text_frame.auto_size = MSO_AUTO_SIZE.NONE`` (or
   ``MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE``) and let PowerPoint handle the
   layout. All font-discovery failures surface as
   :class:`pptx.exc.TextLayoutError`, as does an attempt to fit text
   into a shape whose margins exceed its width or height. See
   issue #168.


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


Copying font formatting
-----------------------

Duplicating the look of one run on another — for example, applying the
title run's bold 24-point coloured formatting to a newly-added run —
used to require re-reading each attribute in turn and re-assigning it.
:meth:`.Font.copy_from` folds that into a single call: every explicit
character property of the source |Font| is copied onto the destination,
and any explicit setting on the destination that the source does *not*
carry is cleared so the two runs end up matching at the XML level.

Only values *explicitly* set on the source are transferred — inherited
("effective") values are **not** resolved first and then written out.
If the source has ``bold = None`` (inheriting from the master style),
the destination ends up with ``bold = None`` too, not with the
master's effective bold setting.

Properties copied:

* :attr:`.Font.bold`, :attr:`.Font.italic`, :attr:`.Font.underline`,
  :attr:`.Font.strikethrough`, :attr:`.Font.size`
* :attr:`.Font.name`, :attr:`.Font.name_ea`, :attr:`.Font.name_cs`
* :attr:`.Font.language_id`
* :attr:`.Font.color` — RGB or theme color, including any
  ``lumMod`` / ``lumOff`` brightness adjustment
* :attr:`.Font.highlight_color`
* :attr:`.Font.use_theme_hyperlink_color` — only when the destination
  run already carries an ``a:hlinkClick``; otherwise the flag has
  nowhere to attach and the copy is a no-op for this property.

The method returns ``self`` so calls chain naturally::

    # --- apply the title's formatting to a new run on another paragraph ---
    title_run = slide.shapes.title.text_frame.paragraphs[0].runs[0]
    body_p = slide.placeholders[1].text_frame.add_paragraph()
    new_run = body_p.add_run()
    new_run.text = "Matches the title"
    new_run.font.copy_from(title_run.font)

Run-level hyperlink relationships (the ``a:hlinkClick`` itself,
``Font.color`` fills beyond a simple ``a:solidFill``, and effects such
as shadow / glow) are **not** copied — :meth:`.Font.copy_from` stays
focused on the character-property surface; use
:attr:`._Run.hyperlink`, :attr:`.Font.shadow`, and
:attr:`.Font.effect_format` directly for those.


Rich text in one call
---------------------

Authoring a paragraph that mixes bold, italic, coloured, and plain runs with
:meth:`._Paragraph.add_run` is accurate but verbose — every run takes four or
five lines to build up. :meth:`._Paragraph.write_rich` folds that loop into a
single call. Each positional argument is a *part* describing one run; the
method appends the runs to the paragraph in order and returns the paragraph so
calls chain naturally.

A part is one of:

* a plain ``str`` — appended as a run with no explicit character formatting
  (it inherits from the paragraph, layout, and master);
* a ``(text, formatting)`` 2-tuple — the formatting mapping is applied to
  the run's ``font``;
* a mapping with a ``"text"`` key — equivalent to the tuple form, with the
  remaining keys treated as the formatting.

Recognised formatting keys are ``bold`` and ``italic`` (tri-state
``True`` / ``False`` / ``None``), ``underline`` (``True`` / ``False`` /
``None`` or a :ref:`MsoTextUnderlineType` member), ``size`` (a |Length|,
most commonly ``Pt(n)``), ``color`` (an |RGBColor|), and ``font_name`` (a
string — the Latin typeface slot). Any other key raises ``ValueError``; use
:meth:`._Paragraph.add_run` directly when you need finer-grained control
(theme colour, East-Asian / complex-script font slots, hyperlink, etc.).

::

    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    p = tb.text_frame.paragraphs[0]

    p.write_rich(
        ("Bold ",   {"bold": True}),
        ("italic ", {"italic": True}),
        "regular ",
        ("red", {"color": RGBColor(0xC0, 0x00, 0x00), "size": Pt(18)}),
    )

``write_rich`` is a thin loop over :meth:`._Paragraph.add_run` — one ``a:r``
element is emitted per part — so runs authored this way round-trip through
PowerPoint, :attr:`._Paragraph.replace_text`, and :meth:`._Paragraph.runs`
exactly as if each was built by hand.


Jumping to another slide from a word in a paragraph
---------------------------------------------------

PowerPoint's *Insert → Hyperlink → Place in This Document* creates an
``a:hlinkClick`` with ``action="ppaction://hlinksldjump"`` on the run of text
the user selected. Assigning a |Slide| to ``run.hyperlink.target_slide``
produces the same XML, so a single word in a paragraph can act as a
click-to-jump into another slide::

    prs = Presentation("deck.pptx")
    overview_slide = prs.slides[0]
    detail_slide = prs.slides[3]

    paragraph = overview_slide.shapes[0].text_frame.paragraphs[0]
    paragraph.text = "Click here for details"

    # the last run in the paragraph is "details"; give it a slide-jump
    details_run = paragraph.add_run()
    details_run.text = " (jump)"
    details_run.hyperlink.target_slide = detail_slide

Reading :attr:`._Hyperlink.target_slide` returns the target |Slide| when the
run carries a slide-jump, and |None| otherwise. To remove a slide-jump on a
run, assign |None| or ``del run.hyperlink.target_slide``. The URL-based
:attr:`._Hyperlink.address` and the slide-jump :attr:`._Hyperlink.target_slide`
are mutually exclusive -- setting one clears the other.
Reading effective font properties
---------------------------------

Most runs do not carry explicit character properties — instead they inherit
size, typeface, bold, italic, and color from the enclosing paragraph, the
text body's list style, the slide layout / master's ``p:txStyles``, or the
presentation's ``p:defaultTextStyle``. Plain ``run.font.size``,
``run.font.bold``, ``run.font.italic``, ``run.font.name``, and
``run.font.color.rgb`` therefore frequently return |None| even though
PowerPoint is rendering something concrete.

The ``effective_*`` read-only properties on |Font| walk the inheritance
chain in the order PowerPoint uses and return the resolved value::

    for para in text_frame.paragraphs:
        for run in para.runs:
            print(run.text,
                  run.font.effective_name,   # e.g. '+mn-lt' for theme minor
                  run.font.effective_size,   # Length (EMU), or None
                  run.font.effective_bold,   # True / False / None
                  run.font.effective_italic, # True / False / None
                  run.font.effective_color)  # RGBColor or None

Each ``effective_*`` property returns |None| when no ancestor in the chain
declares an explicit value (i.e. the run would render using a downstream
fallback, such as PowerPoint's built-in 18pt default). See issue #378.


Text highlight (background) color
---------------------------------

|Font| exposes the text-highlight color — the swatch PowerPoint presents
on the Home ribbon as the text background-color marker — as the
``.highlight_color`` property. It corresponds to the ``a:rPr/a:highlight``
element and behaves like any other |ColorFormat|: both ``.rgb`` and
``.theme_color`` can be assigned.

::

    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_THEME_COLOR

    run.font.highlight_color.rgb = RGBColor(0xFF, 0xFF, 0x00)
    # -- or, theme-driven --
    run.font.highlight_color.theme_color = MSO_THEME_COLOR.ACCENT_1

To remove an explicit highlight entirely (so the run renders with no
background-color marker), call :meth:`.Font.clear_highlight_color`::

    run.font.clear_highlight_color()


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


Numbered lists
--------------

The paragraph-level ``bullet.auto_number()`` API is the building block for
numbered lists: call it on each paragraph you want numbered, passing a scheme
from :class:`PP_AUTO_NUMBER`. PowerPoint auto-numbers consecutive paragraphs
that share a scheme — the second ``ARABIC_PERIOD`` paragraph below renders as
``2.``, the third as ``3.``, etc.::

    from pptx import Presentation
    from pptx.util import Inches
    from pptx.enum.text import PP_AUTO_NUMBER

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(
        Inches(1), Inches(1), Inches(5), Inches(3)
    )
    tf = tb.text_frame

    tf.text = "First item"
    tf.add_paragraph().text = "Second item"
    tf.add_paragraph().text = "Third item"

    for paragraph in tf.paragraphs:
        paragraph.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)

A one-liner equivalent for existing text is to loop over
``text_frame.paragraphs`` after populating text::

    for p in text_frame.paragraphs:
        p.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)

The loop idiom is deliberately the recommended spelling — it reads cleanly,
mirrors the per-paragraph bullet API exactly, and makes per-paragraph
exceptions trivial. To skip a heading paragraph, for example, just branch on
the index::

    for i, p in enumerate(text_frame.paragraphs):
        if i == 0:
            continue  # leave paragraph 0 un-numbered as a heading
        p.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)

To start numbering at a value other than 1, pass ``start_at`` on the first
numbered paragraph only. PowerPoint treats ``start_at`` as the *first*
numbered paragraph's ordinal and increments from there across subsequent
paragraphs using the same scheme — so every paragraph can pass the same
``start_at`` without conflict::

    for p in text_frame.paragraphs:
        p.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD, start_at=5)

Interleaving numbered and unnumbered paragraphs works by clearing the bullet
on the paragraphs that should carry none — PowerPoint resumes counting on the
next numbered paragraph::

    from pptx.enum.text import PP_AUTO_NUMBER

    tf.paragraphs[0].bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)
    tf.paragraphs[1].bullet.none()        # blank spacer paragraph
    tf.paragraphs[2].bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)


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

Text shadow
~~~~~~~~~~~

|Font| exposes the text shadow via :attr:`.Font.shadow`, which returns a
|ShadowFormat|. This is the PowerPoint "Text Effects -> Shadow" checkbox
surfaced as a read/write API on a run. The same four knobs available on
shape / group / chart shadows are available here: :attr:`~.ShadowFormat.blur_radius`,
:attr:`~.ShadowFormat.distance`, :attr:`~.ShadowFormat.direction`, and
:attr:`~.ShadowFormat.color` (issue #546)::

    from pptx.dml.color import RGBColor
    from pptx.util import Emu

    run.font.shadow.blur_radius = Emu(50800)     # 4 pt blur
    run.font.shadow.distance = Emu(38100)        # 3 pt offset
    run.font.shadow.direction = 45.0             # degrees (down-right)
    run.font.shadow.color.rgb = RGBColor(0x80, 0x80, 0x80)

Setting ``run.font.shadow.inherit = True`` removes the explicit effect list
and restores shadow inheritance from the style hierarchy. The full family
of run-level visual effects (``shadow``, ``glow``, ``reflection``,
``soft_edge``) is also available via :attr:`.Font.effect_format`, which
returns an |EffectFormat| object — useful when you need a glow or
reflection in addition to (or instead of) a shadow.

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
Attaching a click-action sound
------------------------------

A shape can be configured to play a WAV-format sound when the user clicks
it (or hovers over it) during a slideshow. The sound is an ``a:snd``
child of the shape's ``a:hlinkClick`` (or ``a:hlinkHover``) element and
is embedded in the package as an audio part.

Use :meth:`~pptx.action.ActionSetting.set_sound` to attach one::

    shape.click_action.set_sound('applause.wav')

The method accepts a filesystem path, a binary file-like object, or a
pre-built :class:`~pptx.media.Audio` instance. A :class:`~pptx.action.Sound`
object is returned and is also accessible via
:attr:`ActionSetting.sound <pptx.action.ActionSetting.sound>`::

    sound = shape.click_action.sound
    sound.name            # 'applause.wav'
    sound.blob            # raw WAV bytestream

Remove an attached sound with
:meth:`~pptx.action.ActionSetting.remove_sound`::

    shape.click_action.remove_sound()


Mouse-hover actions
-------------------

Alongside :attr:`~pptx.shapes.base.BaseShape.click_action`, every shape
also exposes :attr:`~pptx.shapes.base.BaseShape.hover_action` — an
:class:`~pptx.action.ActionSetting` proxy that reads and writes the
shape's ``a:hlinkMouseOver`` element instead of ``a:hlinkClick``. A
hover action fires as the slideshow viewer's mouse pointer passes over
the shape, without a click being required.

The proxy exposes the same API surface as ``click_action``::

    # -- attach a URL that opens on mouse-over --
    shape.hover_action.hyperlink.address = "https://example.com/"
    shape.hover_action.screen_tip = "Hover to open"

    # -- jump to another slide when the pointer passes over the shape --
    shape.hover_action.target_slide = slides[3]

    # -- play a WAV sound on mouse-over --
    shape.hover_action.set_sound("chime.wav")
    shape.hover_action.remove_sound()

Click and hover actions are independent: setting one does not affect the
other, so a shape can carry separate behaviors for click and mouse-over.

.. note::
   PowerPoint only renders a ScreenTip (tooltip) on hover when the
   hyperlink element carries an actionable target — a URL, slide jump,
   embedded sound, or ``ppaction://`` action verb. Setting
   ``click_action.screen_tip`` (or ``hover_action.screen_tip``) on a
   shape that has no other hyperlink target writes a spec-valid but
   behaviorally inert element that PowerPoint will ignore at display
   time. This matches PowerPoint's own Insert Hyperlink dialog, which
   will not accept a ScreenTip without a target. To make a tooltip
   visible, pair it with a URL, slide jump, or sound on the same
   action. This caveat is reported in `issue #1022`_.

.. _`issue #1022`: https://github.com/scanny/python-pptx/issues/1022


.. _templating-text:

Templating text
---------------

A common request is "fill in the blanks" on a deck — load a `.pptx` that
contains placeholders like ``{{ customer_name }}`` or ``{{ order.total }}``
and replace them at render time with values computed in Python. python-pptx
does not ship a templating engine; rendering Jinja2 (or any other template
syntax) is the *caller's* responsibility. What python-pptx provides is the
plumbing:

* :attr:`.TextFrame.text` / :attr:`._Paragraph.text` — read the current text
  out of a shape so you can feed it to your template engine.
* :meth:`.TextFrame.replace_text` / :meth:`._Paragraph.replace_text`
  (added in python-pptx 1.1, see `issue #836`_) — write a replacement back
  into the shape. ``replace_text`` matches across runs, so a token that
  PowerPoint has split across multiple ``a:r`` elements (a common result of
  editing a placeholder in the PowerPoint UI) is still replaced; the run
  containing the start of the match keeps its formatting and absorbs the
  replacement.

The recommended pattern is therefore a three-step loop — **read placeholder
text, apply Jinja2, write back with** ``replace_text`` — run across every
text-bearing shape in the deck.


Rendering an entire deck with Jinja2
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Given a `.pptx` authored with ``{{ ... }}`` / ``{% ... %}`` Jinja2 tokens
inside its text placeholders, the following renders the whole deck in
place. Because :meth:`.TextFrame.replace_text` (and its paragraph
counterpart) scopes each match to a single run-group — matches do not
cross paragraph, line-break (``a:br``), or auto-refresh field (``a:fld``)
boundaries — render **paragraph by paragraph**::

    from jinja2 import Environment
    from pptx import Presentation

    env = Environment(autoescape=False)
    context = {
        "customer_name": "Acme Corp",
        "order": {"total": "$12,450.00", "date": "2025-03-14"},
        "items": ["Widget A", "Widget B", "Widget C"],
    }

    prs = Presentation("template.pptx")

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                original = paragraph.text
                rendered = env.from_string(original).render(**context)
                if rendered != original:
                    paragraph.replace_text(original, rendered)

    prs.save("rendered.pptx")

A few notes on this pattern:

* The run in which the match starts keeps its formatting and absorbs the
  replacement, so per-paragraph "mail merge" preserves fonts, colors, and
  sizes even when PowerPoint has previously split a token across runs.
* Keep each template token on a single paragraph and avoid putting a token
  across a ``Shift+Enter`` soft line-break — ``replace_text`` will not
  match across those boundaries.
* Iterate ``shape.shapes`` recursively to descend into group shapes, and
  use :meth:`.Table.iter_cells` (or nested ``for row in table.rows: for
  cell in row.cells:``) to reach text frames inside table cells, since
  ``slide.shapes`` does not yield them directly.


Per-token replacement
~~~~~~~~~~~~~~~~~~~~~

If you'd rather keep token discovery and rendering separate — for example
to log which tokens were filled, or to handle unresolved tokens explicitly
— you can drive ``replace_text`` one token at a time::

    import re
    from jinja2 import Environment
    from pptx import Presentation

    TOKEN_RE = re.compile(r"\{\{\s*([\w.]+)\s*\}\}")

    env = Environment(autoescape=False)
    context = {"customer_name": "Acme Corp", "year": 2025}

    prs = Presentation("template.pptx")

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text_frame = shape.text_frame
            for match in set(TOKEN_RE.findall(text_frame.text)):
                token = "{{ %s }}" % match
                value = env.from_string(token).render(**context)
                text_frame.replace_text(token, value)

    prs.save("rendered.pptx")

This form is also handy when you want to restrict rendering to a subset
of shapes (for example by shape name) or when the template tokens live
inside table cells or chart text that ``slide.shapes`` does not walk for
you.


Third-party templating libraries
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If you need more than the thin pattern shown above — looped slides,
conditional slide inclusion, image substitution, chart-data substitution,
or block-level tags — a dedicated templating library built on top of
python-pptx will save you a lot of code. Some options to evaluate:

* `python-pptx-templater`_ — Jinja2-style text substitution that walks
  shapes, placeholders, and tables for you.
* `pptx-template`_ — template DSL with looped slides and data-driven
  chart updates.
* `python-pptx-interface`_ — higher-level authoring helpers that compose
  cleanly with the pattern above.

None of these are maintained by the python-pptx project; vet each for
maintenance status, fitness, and license before adopting. The core
``replace_text`` primitive is stable and sufficient for most "mail
merge" style use cases, which is why templating itself is out of scope
for python-pptx (see `issue #398`_).

.. _`issue #398`: https://github.com/scanny/python-pptx/issues/398
.. _`issue #836`: https://github.com/scanny/python-pptx/issues/836
.. _`python-pptx-templater`: https://pypi.org/project/python-pptx-templater/
.. _`pptx-template`: https://pypi.org/project/pptx-template/
.. _`python-pptx-interface`: https://pypi.org/project/python-pptx-interface/
