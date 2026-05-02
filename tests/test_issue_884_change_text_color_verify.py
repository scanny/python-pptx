# pyright: reportPrivateUsage=false

"""Regression test for issue #884 — change text without losing formatting,
and reading a run's color without tripping ``_NoneColor``.

Issue #884 (https://github.com/scanny/python-pptx/issues/884) bundled two
complaints that every templating caller hits simultaneously:

1. **Re-writing the text of a text-frame clobbers the run formatting.**
   The reporter's code::

       shp[1].text_frame.text = "new text"

   replaced the authored run (with its bold / italic / size / color /
   font-family) with a brand-new default run. They wanted a way to
   rewrite just the text, keeping the run-level ``a:rPr`` intact.

2. **Reading the color of an inherited-color run raises.**
   When trying to read back the current color to re-apply after the
   text rewrite, ``font.color.rgb`` raised::

       AttributeError: no .rgb property on color type '_NoneColor'

   because the run's ``a:rPr`` has no explicit ``a:solidFill`` — the
   color is inherited from the theme / master. ``_NoneColor`` has no
   ``.rgb`` to return.

Both sub-issues are resolved by changes shipped in waves 1 / 3 / 6 / 8 /
9 / 10 on this fork:

  * Wave 6 ``feat/issue-836-replace-text-across-runs`` added
    :meth:`TextFrame.replace_text` and
    :meth:`_Paragraph.replace_text`. They flatten each paragraph's
    runs before matching (so a PowerPoint-split token like
    ``"{NA"``/``"ME}"`` is still replaced) and keep the origin run's
    full ``a:rPr`` for the replacement, so callers never have to
    reapply bold / italic / size / font / color after a text rewrite.
    Wave 8 ``feat/issue-285-preserve-format-verify`` pinned the
    preserve-format contract end-to-end via
    ``tests/test_issue_285_replace_text_preserve_format.py``.

  * Wave 1 ``feat/issue-420-universal-to-rgb`` and Wave 3
    ``feat/issue-308-theme-rgb`` shipped a universal
    :meth:`ColorFormat.to_rgb` plus lumMod/lumOff math across every
    color type. Critically, :meth:`ColorFormat.to_rgb` on a
    ``_NoneColor`` returns ``None`` rather than raising — callers can
    safely probe a run's color without the ``AttributeError`` the
    #884 reporter hit.

  * Wave 9 / Wave 10 ``feat/issue-938-effective-font-color`` added
    :attr:`Font.effective_color`, which walks the OOXML inheritance
    chain (``a:rPr`` → paragraph ``a:defRPr`` → text-body
    ``a:lstStyle`` → slide-master ``p:txStyles``), resolves scheme
    colors against the slide-master's theme, and applies any
    ``a:lumMod``/``a:lumOff`` transforms. This is the direct answer to
    the "get the color" half of the #884 report — callers can now
    read the RGB PowerPoint would actually render for a run whose
    color is only set by the theme.

This regression test pins both halves of #884 at the user-facing
`Presentation` API level, round-tripping through
``Presentation.save`` + reopen so a future regression of *any* of
#836 / #420 / #308 / #938 fails loudly against the reporter's
original wording.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Inches, Pt


class DescribeIssue884ChangeTextAndReadColor:
    """End-to-end regression for both halves of #884."""

    # -- Part 1: rewrite text without touching the formatting ------------------

    def it_replaces_text_without_losing_run_formatting(self):
        """The reporter's core grievance: ``text_frame.text = "new"`` lost
        every run-level style.

        The fix is :meth:`TextFrame.replace_text`: the run where the match
        starts keeps its entire ``a:rPr`` and absorbs the replacement
        text, so bold / italic / size / font / color all survive a text
        rewrite with no work from the caller.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # -- blank layout
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tf = textbox.text_frame

        # -- author a single richly-formatted run, like the #884 reporter
        # -- would get from a hand-edited template placeholder --
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "Hello OLD_NAME!"
        run.font.bold = True
        run.font.italic = True
        run.font.underline = True
        run.font.size = Pt(18)
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(0xC0, 0x10, 0x40)

        # -- act: rewrite the templated token. Unlike the reporter's
        # -- `text_frame.text = ...`, `replace_text` is formatting-preserving --
        count = tf.replace_text("OLD_NAME", "World")

        assert count == 1
        assert tf.text == "Hello World!"

        # -- every run-level style survives; caller does NOT have to
        # -- re-apply anything (the exact pain-point of #884) --
        rewritten = tf.paragraphs[0].runs[0]
        assert rewritten.text == "Hello World!"
        assert rewritten.font.bold is True
        assert rewritten.font.italic is True
        assert rewritten.font.underline is True
        assert rewritten.font.size == Pt(18)
        assert rewritten.font.name == "Calibri"
        assert rewritten.font.color.rgb == RGBColor(0xC0, 0x10, 0x40)

        # -- round-trip through save + reopen; formatting must persist --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]

        assert reloaded.text == "Hello World!"
        assert reloaded.font.bold is True
        assert reloaded.font.italic is True
        assert reloaded.font.underline is True
        assert reloaded.font.size == Pt(18)
        assert reloaded.font.name == "Calibri"
        assert reloaded.font.color.rgb == RGBColor(0xC0, 0x10, 0x40)

    def it_replaces_text_across_split_runs_preserving_origin_formatting(self):
        """The PowerPoint-splits-the-token case the #884 reporter's
        placeholder workflow actually produces.

        After in-editor edits PowerPoint frequently splits a token such
        as ``{NAME}`` across multiple ``a:r`` elements. ``replace_text``
        still replaces it, and the origin run's formatting is what
        survives — so the caller sees a single run carrying the
        original rPr.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tf = textbox.text_frame
        p = tf.paragraphs[0]

        # -- origin run: bold + 20 pt + navy --
        r1 = p.add_run()
        r1.text = "Hello {NA"
        r1.font.bold = True
        r1.font.size = Pt(20)
        r1.font.color.rgb = RGBColor(0x00, 0x20, 0x60)

        # -- trailing run deliberately carries different rPr so we can
        # -- see which one wins after the replace --
        r2 = p.add_run()
        r2.text = "ME}!"
        r2.font.italic = True
        r2.font.size = Pt(10)
        r2.font.color.rgb = RGBColor(0xFF, 0xFF, 0x00)

        count = tf.replace_text("{NAME}", "Claude")
        assert count == 1
        assert tf.text == "Hello Claude!"

        runs = tf.paragraphs[0].runs
        assert len(runs) == 2

        # -- merged origin run keeps r1's rPr for "Hello Claude" --
        assert runs[0].text == "Hello Claude"
        assert runs[0].font.bold is True
        assert runs[0].font.italic is None
        assert runs[0].font.size == Pt(20)
        assert runs[0].font.color.rgb == RGBColor(0x00, 0x20, 0x60)

        # -- r2's suffix "!" keeps r2's own rPr --
        assert runs[1].text == "!"
        assert runs[1].font.bold is None
        assert runs[1].font.italic is True
        assert runs[1].font.size == Pt(10)
        assert runs[1].font.color.rgb == RGBColor(0xFF, 0xFF, 0x00)

        # -- round-trip re-asserts everything --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs
        assert [r.text for r in reloaded] == ["Hello Claude", "!"]
        assert reloaded[0].font.bold is True
        assert reloaded[0].font.size == Pt(20)
        assert reloaded[0].font.color.rgb == RGBColor(0x00, 0x20, 0x60)
        assert reloaded[1].font.italic is True
        assert reloaded[1].font.color.rgb == RGBColor(0xFF, 0xFF, 0x00)

    # -- Part 2: read a run's color without the _NoneColor AttributeError -----

    def it_reads_color_without_raising_on_inherited_theme_color(self):
        """The reporter's ``AttributeError: no .rgb property on color type
        '_NoneColor'`` disappears once the caller uses :meth:`to_rgb` or
        :attr:`Font.effective_color`.

        A freshly-added textbox carries a run whose ``a:rPr`` has no
        explicit ``a:solidFill`` — its color inherits from the theme.
        The pre-#420 path ``font.color.rgb`` raises on that, because
        ``_NoneColor.rgb`` has no RGB value to return. The #420 / #308
        universal ``to_rgb()`` returns |None| for a ``_NoneColor``
        without raising, and the #938 ``effective_color`` walks the
        inheritance chain and resolves the theme color to an actual
        |RGBColor|.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tf = textbox.text_frame

        # -- author a run with no explicit color (inherits from theme) --
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "untouched"
        # -- deliberately set *no* color on the run --
        color = run.font.color
        assert color.type is None  # -- this is the _NoneColor state #884 hit --

        # -- pre-#420 behaviour reproducer: plain `.rgb` still raises on
        # -- an inherited color. That's the caller's original bug. We
        # -- pin it only to make the before/after story explicit. --
        with pytest.raises(AttributeError, match="_NoneColor"):
            _ = run.font.color.rgb

        # -- #420 / #308 workaround: universal `to_rgb()` returns None
        # -- instead of raising --
        assert run.font.color.to_rgb() is None

        # -- #938 workaround: `effective_color` walks the inheritance
        # -- chain and resolves the theme color to an actual RGB --
        resolved = run.font.effective_color
        assert isinstance(resolved, RGBColor)
        # -- the default textbox inherits from `otherStyle` at level 0,
        # -- which in the default theme resolves to the `tx1` scheme
        # -- color (black on the stock theme). We assert only that a
        # -- real RGB value came back — exact shade is theme-dependent
        # -- and not the point of the regression. --

    def it_resolves_theme_color_on_an_explicit_scheme_color_run(self):
        """Once the caller sets an explicit theme color, reading the run
        color back via :meth:`to_rgb` or :attr:`Font.effective_color`
        resolves it against the slide-master's theme — another failure
        mode the #884 reporter would have hit when trying to preserve a
        theme-colored run.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tf = textbox.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "themed"
        run.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

        # -- `.rgb` still raises on a scheme color (no RGB on _SchemeColor) --
        with pytest.raises(AttributeError):
            _ = run.font.color.rgb

        # -- but `to_rgb()` resolves the scheme color against theme_colors --
        master = prs.slide_masters[0]
        theme_colors = master.theme_colors
        resolved = run.font.color.to_rgb(theme_colors)
        assert isinstance(resolved, RGBColor)

        # -- `effective_color` resolves it too, without the caller
        # -- needing to pass theme_colors explicitly --
        effective = run.font.effective_color
        assert effective == resolved

        # -- round-trip: the scheme color persists and still resolves --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded_run = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert reloaded_run.font.color.theme_color == MSO_THEME_COLOR.ACCENT_1
        reloaded_theme = prs2.slide_masters[0].theme_colors
        assert reloaded_run.font.color.to_rgb(reloaded_theme) == resolved
        assert reloaded_run.font.effective_color == resolved

    # -- Part 3: the reporter's combined workflow ------------------------------

    def it_supports_the_884_reporter_end_to_end_workflow(self):
        """The full #884 story: read color, rewrite text, assert both
        halves keep working without the caller having to touch runs.

        This is the exact workflow the #884 reporter described —
        "get the color, format, etc. and apply after changing" — which
        they couldn't complete because of the ``AttributeError`` on the
        color read and the formatting loss on the text write. With
        :meth:`to_rgb` / :attr:`effective_color` + :meth:`replace_text`
        both problems are gone and no re-apply step is needed.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tf = textbox.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "Dear CUSTOMER, welcome!"
        run.font.bold = True
        run.font.size = Pt(22)
        run.font.color.rgb = RGBColor(0x33, 0x66, 0x99)

        # -- step 1 (no longer raises): read the current color via
        # -- the #420 / #938 APIs. Both agree for an explicit RGB run. --
        assert run.font.color.to_rgb() == RGBColor(0x33, 0x66, 0x99)
        assert run.font.effective_color == RGBColor(0x33, 0x66, 0x99)

        # -- step 2: rewrite the templated token via the #836 API;
        # -- formatting survives without any re-apply step --
        count = tf.replace_text("CUSTOMER", "Alice")
        assert count == 1
        assert tf.text == "Dear Alice, welcome!"

        # -- step 3: the reporter's goal — text is new, format intact --
        final = tf.paragraphs[0].runs[0]
        assert final.font.bold is True
        assert final.font.size == Pt(22)
        assert final.font.color.rgb == RGBColor(0x33, 0x66, 0x99)
        assert final.font.effective_color == RGBColor(0x33, 0x66, 0x99)

        # -- and it round-trips --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded_tf = prs2.slides[0].shapes[0].text_frame
        assert reloaded_tf.text == "Dear Alice, welcome!"
        reloaded_run = reloaded_tf.paragraphs[0].runs[0]
        assert reloaded_run.font.bold is True
        assert reloaded_run.font.size == Pt(22)
        assert reloaded_run.font.color.rgb == RGBColor(0x33, 0x66, 0x99)
        assert reloaded_run.font.effective_color == RGBColor(0x33, 0x66, 0x99)
