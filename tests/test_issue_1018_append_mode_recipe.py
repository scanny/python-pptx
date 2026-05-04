"""Regression test for issue #1018 — "open a PPT in append mode".

Issue #1018 (https://github.com/scanny/python-pptx/issues/1018) asked
whether python-pptx supports an "append mode" for adding slides to an
existing presentation. It does not — and cannot — because the
``.pptx`` package format is a ZIP of cross-referencing parts (the
presentation part, the slide parts, the relationships graph, the
content-types manifest) that has to be rewritten as a whole on every
save. Adding a slide mutates the ``p:sldIdLst`` in
``/ppt/presentation.xml``, adds a new ``/ppt/slides/slideN.xml`` part,
and updates ``[Content_Types].xml``; there is no way to splice those
changes into an archive without rewriting it.

The intended "append" idiom is the ordinary read / modify / save
cycle — open the deck, add slides, save either in place or to a new
path. That is the recipe documented in ``docs/user/use-cases.rst`` →
"Frequently asked: why no append mode?". This suite pins that recipe
against the kind of silent regression that would break the documented
answer:

  * Build a "prior" deck that already contains a slide, save it to a
    temporary location, and read it back — no shortcuts that bypass
    the real :class:`.Presentation` constructor.
  * Open with ``Presentation(existing_path)``, add new slides via
    :meth:`.Slides.add_slide` onto a layout drawn from the existing
    deck's :attr:`.Presentation.slide_layouts`.
  * Save to the same path (in-place overwrite) *and* to a different
    path — both call shapes that the docs recipe presents as
    equivalent.
  * Reopen each saved package from scratch and verify the new slides
    are present with their content intact and the prior slides
    unchanged. Slide order must match what the caller constructed:
    pre-existing slides first, appended slides last.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.util import Inches, Pt


def _build_prior_deck() -> io.BytesIO:
    """Return an in-memory ``.pptx`` simulating a pre-existing deck."""
    prs = Presentation()
    # -- one pre-existing slide with a recognisable title --
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # "Title Only"
    slide.shapes.title.text = "Existing slide"
    stream = io.BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream


def _title_of(slide) -> str:
    """Return the title text of *slide*, or an empty string."""
    title = slide.shapes.title
    return title.text if title is not None else ""


def _textbox_strings(slide) -> list[str]:
    """Return the text of every plain textbox on *slide*."""
    out: list[str] = []
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.is_placeholder:
            out.append(shape.text_frame.text)
    return out


class DescribeIssue1018AppendModeRecipe:
    """The documented read / modify / save "append" recipe round-trips."""

    def it_appends_slides_and_saves_to_a_new_path(self):
        prior = _build_prior_deck()

        # -- 1. open the existing deck --
        prs = Presentation(prior)
        assert len(prs.slides) == 1
        assert _title_of(prs.slides[0]) == "Existing slide"

        # -- 2. add two new slides, mirroring the docs snippet --
        layout = prs.slide_layouts[5]  # "Title Only"
        slide_a = prs.slides.add_slide(layout)
        slide_a.shapes.title.text = "Appended slide A"
        slide_b = prs.slides.add_slide(layout)
        slide_b.shapes.title.text = "Appended slide B"
        # -- also drop a plain textbox so we verify more than the title --
        textbox = slide_b.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
        textbox.text_frame.text = "appended body"

        # -- 3. save to a *different* path --
        out = io.BytesIO()
        prs.save(out)

        # -- 4. reopen from scratch and verify round-trip --
        out.seek(0)
        reopened = Presentation(out)
        assert len(reopened.slides) == 3
        titles = [_title_of(s) for s in reopened.slides]
        assert titles == [
            "Existing slide",
            "Appended slide A",
            "Appended slide B",
        ]
        # -- appended textbox survived --
        assert "appended body" in _textbox_strings(reopened.slides[2])

    def it_appends_slides_and_overwrites_in_place(self, tmp_path):
        # -- create a real on-disk .pptx the caller might have authored --
        path = tmp_path / "deck.pptx"
        path.write_bytes(_build_prior_deck().getvalue())

        # -- append + overwrite in place --
        prs = Presentation(str(path))
        layout = prs.slide_layouts[5]
        appended = prs.slides.add_slide(layout)
        appended.shapes.title.text = "In-place append"
        prs.save(str(path))  # same path — docs snippet's in-place form

        # -- reopen from disk; prior slide preserved, new slide present --
        reopened = Presentation(str(path))
        assert len(reopened.slides) == 2
        assert _title_of(reopened.slides[0]) == "Existing slide"
        assert _title_of(reopened.slides[1]) == "In-place append"

    def it_preserves_existing_slide_content_when_appending(self):
        # -- author a prior deck whose content we pin afterwards --
        prior = Presentation()
        s0 = prior.slides.add_slide(prior.slide_layouts[5])
        s0.shapes.title.text = "Prior title"
        tb = s0.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
        tb_tf = tb.text_frame
        tb_tf.text = "prior body"
        # -- stash a specific font size we'll assert survives --
        tb_tf.paragraphs[0].runs[0].font.size = Pt(18)

        stream = io.BytesIO()
        prior.save(stream)
        stream.seek(0)

        # -- append --
        prs = Presentation(stream)
        appended = prs.slides.add_slide(prs.slide_layouts[5])
        appended.shapes.title.text = "Follower"
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)

        # -- reopen and pin prior slide content byte-for-byte --
        reopened = Presentation(out)
        assert len(reopened.slides) == 2
        s0_r = reopened.slides[0]
        assert _title_of(s0_r) == "Prior title"
        textboxes = [
            shape for shape in s0_r.shapes if shape.has_text_frame and not shape.is_placeholder
        ]
        assert len(textboxes) == 1
        tf = textboxes[0].text_frame
        assert tf.text == "prior body"
        assert tf.paragraphs[0].runs[0].font.size == Pt(18)

    def it_supports_multiple_append_save_cycles(self, tmp_path):
        # -- simulate the user running the recipe repeatedly, each cycle
        # -- opening the previously-saved file and adding one more slide --
        path = tmp_path / "rolling.pptx"
        path.write_bytes(_build_prior_deck().getvalue())

        for i in range(3):
            prs = Presentation(str(path))
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = f"Cycle {i}"
            prs.save(str(path))

        final = Presentation(str(path))
        titles = [_title_of(s) for s in final.slides]
        assert titles == [
            "Existing slide",
            "Cycle 0",
            "Cycle 1",
            "Cycle 2",
        ]
