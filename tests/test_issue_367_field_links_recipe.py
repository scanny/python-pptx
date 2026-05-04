# pyright: reportPrivateUsage=false

"""Regression tests for issue #367 — cross-slide text-linking recipe.

Issue #367 (https://github.com/scanny/python-pptx/issues/367) asked for
a way to *link* two text fields across slides so that editing one
automatically propagates to the other — the PowerPoint equivalent of an
Excel cell reference (``=A1``) or a Word ``REF`` cross-reference field.

PowerPoint / OOXML does not expose any such mechanism. The ``<a:fld>``
element has a fixed, closed set of ``@type`` values (slide number,
date/time in thirteen formats, ``datetimeFigureOut``, and ``footer``) —
all of which are computed by PowerPoint itself, not driven by the
content of another text frame. There is no bookmark primitive in
PresentationML, no formula layer in DrawingML, and no relationship type
that binds one run to another.

The canonical python-pptx answer — documented in
``docs/user/linked-content.rst`` — is to treat the generation script as
the single source of truth and re-author the shared string onto every
target slide. This regression suite pins the primitives that recipe
relies on so that a future change to the text API does not silently
break the documented workflow:

1. :meth:`._Paragraph.add_field` still authors only the known PowerPoint
   field types — this is what forces callers to fall back to the
   re-author recipe for user-defined cross-references.
2. The single-source-of-truth re-author loop (add a textbox on every
   slide with a Python variable as the source) round-trips cleanly
   through save + reopen.
3. Updating the Python variable and re-running the loop against an
   existing deck — the "customer renamed the project" variant — yields
   a fully refreshed deck with no stale copies of the old string.
4. The alternative "author a shape on the master with a recognisable
   name, rewrite its text per deck" variant also round-trips.
5. :meth:`.TextFrame.replace_text`, the token-substitution variant of
   the same pattern, swaps every ``{{ token }}`` marker across all
   slides and survives save + reopen.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.util import Inches, Pt


class DescribeIssue367FieldLinksRecipe:
    """Regression suite for the #367 cross-slide text-linking recipe.

    Exercises the primitives the ``docs/user/linked-content.rst`` recipe
    depends on: ``add_field`` for the native-fields callout,
    textbox authoring for the re-author loop, and ``replace_text`` for
    the token-substitution variant.
    """

    # -- 1. Only the known PowerPoint field types are supported ----------

    def it_only_authors_the_known_powerpoint_field_types(self):
        """``add_field`` exposes the closed set of PowerPoint field types.

        The doc's "why can't I author a custom field" argument rests on
        this: the supported ``@type`` values are the hard-coded set
        PowerPoint recognises. This pins the slide-number + datetime +
        footer triad so a regression that (say) stopped emitting
        ``a:fld/@type`` would be caught.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(0.5))
        p = tb.text_frame.paragraphs[0]

        f_num = p.add_field("slidenum", "#")
        f_dt = p.add_field("datetime", "")
        f_ft = p.add_field("footer", "")

        assert f_num.field_type == "slidenum"
        assert f_dt.field_type == "datetime"
        assert f_ft.field_type == "footer"

    # -- 2. The single-source-of-truth re-author loop round-trips --------

    def it_stamps_a_shared_header_onto_every_slide_from_one_variable(self):
        """Author the same project-name header across every slide.

        The core recipe: iterate ``prs.slides`` and add a textbox whose
        text comes from a single Python variable. Save + reopen and
        confirm every slide carries the same string in the expected
        location — the workflow the #367 reporter adopts in lieu of a
        native cross-slide field link.
        """
        project_name = "Apollo - Q3 2026 review"
        prs = Presentation()
        # -- add three slides so the "every slide" assertion is non-trivial --
        for _ in range(3):
            prs.slides.add_slide(prs.slide_layouts[5])

        # -- act: stamp the header onto every slide from one variable --
        for slide in prs.slides:
            tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9), Inches(0.4))
            p = tb.text_frame.paragraphs[0]
            p.add_run(project_name, bold=True, size=Pt(14))

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        assert len(prs2.slides) == 3
        for slide in prs2.slides:
            # -- the project-name run appears on every slide --
            texts = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
            assert project_name in texts

    def it_refreshes_every_slide_when_the_source_variable_changes(self):
        """Re-run the recipe against an existing deck with a new value.

        The "customer renamed the project" scenario: open an already-
        generated deck, find the header shapes (by name, here), and
        rewrite them from the new source string. No stale copies of the
        old name remain after save + reopen — which is the whole point
        of keeping one source of truth in the script.
        """
        prs = Presentation()
        for _ in range(2):
            prs.slides.add_slide(prs.slide_layouts[5])

        # -- round 1: stamp the original name and name the shape so the
        # -- refresh pass can find it later --
        original = "Apollo"
        for slide in prs.slides:
            tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9), Inches(0.4))
            tb.name = "ProjectHeader"
            tb.text_frame.text = original

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        # -- round 2: open the deck, rewrite every ProjectHeader from a
        # -- new source value --
        renamed = "Artemis"
        prs2 = Presentation(buf)
        for slide in prs2.slides:
            for shape in slide.shapes:
                if shape.name == "ProjectHeader":
                    shape.text_frame.text = renamed

        # -- save + reload once more to prove the overwrite persisted --
        buf2 = io.BytesIO()
        prs2.save(buf2)
        buf2.seek(0)
        prs3 = Presentation(buf2)

        # -- every slide now carries the new name and none the old --
        for slide in prs3.slides:
            header_texts = [s.text_frame.text for s in slide.shapes if s.name == "ProjectHeader"]
            assert renamed in header_texts
            assert original not in header_texts

    # -- 3. The master-shape variant also round-trips --------------------

    def it_supports_the_master_shape_lookup_by_name_variant(self):
        """Author a named shape on the master, rewrite its text per deck.

        The alternative recipe from ``linked-content.rst``: when the
        header already exists on the master (authored in the template),
        iterate slide shapes by name and rewrite the text per
        generation. Here we simulate "authored on the master" by adding
        a named shape to a fresh slide and proving the name-based
        lookup + rewrite round-trips.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9), Inches(0.4))
        tb.name = "ProjectHeader"
        tb.text_frame.text = "PLACEHOLDER"

        # -- the recipe: find shapes by name, rewrite from one variable --
        project_name = "Apollo - Q3 2026 review"
        for shape in slide.shapes:
            if shape.name == "ProjectHeader":
                shape.text_frame.text = project_name

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        named = [s for s in slide2.shapes if s.name == "ProjectHeader"]
        assert len(named) == 1
        assert named[0].text_frame.text == project_name

    # -- 4. The token-substitution variant (TextFrame.replace_text) ------

    def it_supports_the_token_substitution_variant_via_replace_text(self):
        """``{{ token }}`` substitution across every slide round-trips.

        The recipe's "token marker" alternative: pre-author
        ``{{ project_name }}`` markers in the template, then substitute
        them from a mapping at generation time. This exercises
        :meth:`.TextFrame.replace_text` — the same primitive the
        templating recipe in ``docs/user/text.rst`` documents — on a
        multi-slide deck to prove the cross-slide reach.
        """
        prs = Presentation()
        for _ in range(3):
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9), Inches(0.4))
            tb.text_frame.text = "Title: {{ project_name }}"

        # -- act: substitute the marker on every slide from one value --
        project_name = "Apollo"
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    shape.text_frame.replace_text("{{ project_name }}", project_name)

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        for slide in prs2.slides:
            texts = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
            assert any(project_name in t for t in texts)
            assert not any("{{ project_name }}" in t for t in texts)
