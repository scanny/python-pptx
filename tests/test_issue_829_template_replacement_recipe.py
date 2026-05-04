# pyright: reportPrivateUsage=false

"""Regression test for issue #829 — template-style text + picture replacement.

Issue #829 (https://github.com/scanny/python-pptx/issues/829) collected
utility functions several reporters had written around a common
workflow: *"open a template deck, replace every ``{{ token }}`` in the
text, and swap identified pictures for per-customer artwork, then
save."* The underlying capabilities have shipped on the public API:

* :meth:`.TextFrame.replace_text` / :meth:`._Paragraph.replace_text`
  — replace literal strings across the runs of a text frame while
  preserving the origin run's formatting (issues #285 / #836).
* :meth:`.Picture.replace_image` — swap an existing picture's pixel
  bytes while preserving its position, size, rotation, cropping,
  masking shape, and identifying ``cNvPr/@name`` / ``cNvPr/@descr``
  attributes (issues #116 / #834 / #819).

The docs recipe in ``docs/user/text.rst`` → "Template-style
replacement: text and pictures" glues these two primitives together
with a ``cNvPr/@name`` lookup (the Selection-pane name PowerPoint
authors use) and a ``shapes`` descent into groups. This suite pins
the end-to-end flow against the kind of silent regression that would
make the recipe stop working:

  * Build a template deck holding per-slide ``{{ customer_name }}`` /
    ``{{ order_total }}`` / ``{{ close_date }}`` tokens across body
    placeholders and plain textboxes, plus two named pictures
    ("CustomerLogo", "AccountPhoto").
  * Apply the documented loop (TOKENS dict + PICTURES dict keyed by
    ``shape.name``).
  * Save + reopen — every text replacement and every picture swap
    must have survived the round-trip.
  * A second slide with a nested group containing a tokenised textbox
    pins the ``iter_shapes`` group-descent helper in the recipe.
  * A variation keying the picture map by ``shape.alt_text`` instead
    of ``shape.name`` pins the alternative predicate called out in
    the docs.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

# -- documented recipe helpers (mirror the snippets in docs/user/text.rst) --


def iter_shapes(shapes):
    """Yield every shape under *shapes*, descending into groups."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)
        else:
            yield shape


def apply_tokens(prs, tokens):
    """Apply *tokens* dict to every text-bearing shape in *prs*."""
    for slide in prs.slides:
        for shape in iter_shapes(slide.shapes):
            if shape.has_text_frame:
                for find, replace in tokens.items():
                    shape.text_frame.replace_text(find, replace)


def apply_pictures_by_name(prs, pictures):
    """Swap pictures in *prs* whose shape name matches a key in *pictures*."""
    for slide in prs.slides:
        for shape in iter_shapes(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                new_image = pictures.get(shape.name)
                if new_image is not None:
                    shape.replace_image(new_image)


def apply_pictures_by_alt_text(prs, pictures):
    """Swap pictures in *prs* whose alt_text matches a key in *pictures*."""
    for slide in prs.slides:
        for shape in iter_shapes(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                new_image = pictures.get(shape.alt_text)
                if new_image is not None:
                    shape.replace_image(new_image)


# -- fixture builders ----------------------------------------------------


def _tagged_picture(slide, image_path, name, alt_text=None, x=Inches(1), y=Inches(1)):
    """Add a picture to *slide* and stamp it with *name* (and optional alt_text)."""
    picture = slide.shapes.add_picture(image_path, x, y, Inches(2), Inches(2))
    picture.name = name
    if alt_text is not None:
        picture.alt_text = alt_text
    return picture


def _template_deck():
    """Return a fresh |Presentation| shaped like the docs-recipe template.

    Two slides:
      * Slide 0 — title + body placeholder + a named logo picture and
        an additional textbox. Each text target carries a distinct
        Jinja2-style token so the recipe has something to replace.
      * Slide 1 — blank layout with a group shape containing a
        tokenised textbox and an "AccountPhoto" picture positioned
        outside the group. Pins the group-descent branch of
        ``iter_shapes``.
    """
    prs = Presentation()

    # -- slide 0 ---------------------------------------------------------
    slide0 = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
    slide0.shapes.title.text_frame.text = "Proposal for {{ customer_name }}"
    body = slide0.placeholders[1]
    body.text_frame.text = "Total: {{ order_total }}"
    tb = slide0.shapes.add_textbox(Inches(1), Inches(5), Inches(5), Inches(0.5))
    tb.text_frame.text = "Closes {{ close_date }}"
    _tagged_picture(
        slide0,
        "tests/test_files/python-powered.png",
        name="CustomerLogo",
        alt_text="customer-logo",
        x=Inches(7),
        y=Inches(0.5),
    )

    # -- slide 1 ---------------------------------------------------------
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    # -- nested group with a tokenised textbox --
    group_tb = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(0.5))
    group_tb.text_frame.text = "Account owner: {{ customer_name }}"
    # -- named account photo (alt_text keyed) --
    _tagged_picture(
        slide1,
        "tests/test_files/python-icon.jpeg",
        name="AccountPhoto",
        alt_text="account-photo",
        x=Inches(6),
        y=Inches(1),
    )

    return prs


def _deck_with_grouped_token():
    """Template deck whose token lives inside a *group* shape.

    Pins the ``iter_shapes`` group-descent branch: a token placed inside
    a group must still be replaced by the documented loop.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # -- build two textboxes, then group them --
    tb1 = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(0.5))
    tb1.text_frame.text = "Hello {{ customer_name }}"
    tb2 = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(0.5))
    tb2.text_frame.text = "Order: {{ order_total }}"

    slide.shapes.add_group_shape([tb1, tb2])
    return prs


# -- fixtures & tests ----------------------------------------------------


class DescribeIssue829TemplateReplacementRecipe:
    """End-to-end coverage of the #829 text + picture template recipe."""

    def it_replaces_every_text_token_across_the_deck(self):
        # -- the documented loop must replace every token it's given,
        # -- regardless of whether the target is a title placeholder,
        # -- a body placeholder, or an authored textbox.
        prs = _template_deck()
        tokens = {
            "{{ customer_name }}": "Acme Corp",
            "{{ order_total }}": "$12,450.00",
            "{{ close_date }}": "2026-03-14",
        }

        apply_tokens(prs, tokens)

        all_text = "\n".join(
            shape.text_frame.text
            for slide in prs.slides
            for shape in iter_shapes(slide.shapes)
            if shape.has_text_frame
        )
        for token in tokens:
            assert token not in all_text, "token %r was not replaced — the recipe regressed" % token
        assert "Acme Corp" in all_text
        assert "$12,450.00" in all_text
        assert "2026-03-14" in all_text

    def it_swaps_named_pictures_by_shape_name(self):
        # -- keying the picture map off ``shape.name`` (the Selection-pane
        # -- name) must update the pixel bytes of the matching picture
        # -- while leaving every identifying attribute intact.
        prs = _template_deck()
        pictures = {
            "CustomerLogo": "tests/test_files/monty-truth.png",
        }

        # -- capture the original logo's identifying attributes --
        logo_before = next(
            s
            for slide in prs.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.name == "CustomerLogo"
        )
        original_descr = logo_before._element.nvPicPr.cNvPr.get("descr")
        original_name = logo_before._element.nvPicPr.cNvPr.get("name")
        original_blob = logo_before.image.blob

        apply_pictures_by_name(prs, pictures)

        with open("tests/test_files/monty-truth.png", "rb") as f:
            expected = f.read()

        logo_after = next(
            s
            for slide in prs.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.name == "CustomerLogo"
        )
        assert logo_after.image.blob == expected
        assert logo_after.image.blob != original_blob
        # -- identifying attributes preserved so the next run still finds --
        # -- the same shape by name ---------------------------------------
        assert logo_after._element.nvPicPr.cNvPr.get("name") == original_name
        assert logo_after._element.nvPicPr.cNvPr.get("descr") == original_descr

    def it_leaves_pictures_not_in_the_map_alone(self):
        # -- the recipe must only replace pictures whose name is in the
        # -- map; unmapped pictures keep their original bytes.
        prs = _template_deck()
        pictures = {
            "CustomerLogo": "tests/test_files/monty-truth.png",
            # -- note: AccountPhoto is intentionally absent --
        }

        with open("tests/test_files/python-icon.jpeg", "rb") as f:
            account_photo_bytes = f.read()

        apply_pictures_by_name(prs, pictures)

        account_photo = next(
            s
            for slide in prs.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.name == "AccountPhoto"
        )
        assert account_photo.image.blob == account_photo_bytes

    def it_round_trips_text_and_pictures_through_save_and_reopen(self):
        # -- the headline recipe: apply tokens, swap pictures, save,
        # -- reopen. Every change must survive the round-trip.
        prs = _template_deck()
        tokens = {
            "{{ customer_name }}": "Acme Corp",
            "{{ order_total }}": "$12,450.00",
            "{{ close_date }}": "2026-03-14",
        }
        pictures = {
            "CustomerLogo": "tests/test_files/monty-truth.png",
            "AccountPhoto": "tests/test_files/python-powered.png",
        }

        apply_tokens(prs, tokens)
        apply_pictures_by_name(prs, pictures)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reloaded = Presentation(buf)

        # -- text: every token is gone and every substitution is present --
        all_text = "\n".join(
            shape.text_frame.text
            for slide in reloaded.slides
            for shape in iter_shapes(slide.shapes)
            if shape.has_text_frame
        )
        for token in tokens:
            assert token not in all_text
        assert "Acme Corp" in all_text
        assert "$12,450.00" in all_text
        assert "2026-03-14" in all_text

        # -- pictures: both were swapped to the new bytes ------------------
        with open("tests/test_files/monty-truth.png", "rb") as f:
            logo_bytes = f.read()
        with open("tests/test_files/python-powered.png", "rb") as f:
            photo_bytes = f.read()

        reloaded_pictures = {
            s.name: s
            for slide in reloaded.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE
        }
        assert reloaded_pictures["CustomerLogo"].image.blob == logo_bytes
        assert reloaded_pictures["AccountPhoto"].image.blob == photo_bytes

    def it_descends_into_group_shapes_for_text_replacement(self):
        # -- the docs ``iter_shapes`` helper descends into groups. A token
        # -- authored inside a group must still be replaced.
        prs = _deck_with_grouped_token()
        tokens = {
            "{{ customer_name }}": "Acme Corp",
            "{{ order_total }}": "$12,450.00",
        }

        apply_tokens(prs, tokens)

        # -- verify both textboxes (now inside a group) were touched --
        group_text = "\n".join(
            shape.text_frame.text
            for slide in prs.slides
            for shape in iter_shapes(slide.shapes)
            if shape.has_text_frame
        )
        assert "{{ customer_name }}" not in group_text
        assert "{{ order_total }}" not in group_text
        assert "Hello Acme Corp" in group_text
        assert "Order: $12,450.00" in group_text

    def it_supports_the_alt_text_keyed_picture_variant(self):
        # -- the docs call out ``shape.alt_text`` as the alternative key
        # -- when the deck's shape names are not reliable. The two helpers
        # -- share a find-then-swap loop so both must work end-to-end.
        prs = _template_deck()
        pictures = {
            "customer-logo": "tests/test_files/monty-truth.png",
        }

        apply_pictures_by_alt_text(prs, pictures)

        with open("tests/test_files/monty-truth.png", "rb") as f:
            expected = f.read()

        logo = next(
            s
            for slide in prs.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.alt_text == "customer-logo"
        )
        assert logo.image.blob == expected

    def it_preserves_picture_geometry_across_the_replace(self):
        # -- the reporter's motivation for a supported swap (over a
        # -- delete + add_picture pair) is that the geometry survives.
        prs = _template_deck()

        original_logo = next(
            s
            for slide in prs.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.name == "CustomerLogo"
        )
        original_left = original_logo.left
        original_top = original_logo.top
        original_width = original_logo.width
        original_height = original_logo.height

        apply_pictures_by_name(prs, {"CustomerLogo": "tests/test_files/monty-truth.png"})

        swapped_logo = next(
            s
            for slide in prs.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.name == "CustomerLogo"
        )
        assert swapped_logo.left == original_left
        assert swapped_logo.top == original_top
        assert swapped_logo.width == original_width
        assert swapped_logo.height == original_height

    def it_runs_the_recipe_idempotently_across_two_cycles(self):
        # -- the reporter's scenario included re-running the program
        # -- against the output of an earlier run. Applying the recipe
        # -- twice must not break anything; the second cycle is a no-op
        # -- for already-replaced tokens and re-applies the picture swap.
        prs = _template_deck()
        tokens = {
            "{{ customer_name }}": "Acme Corp",
            "{{ order_total }}": "$12,450.00",
            "{{ close_date }}": "2026-03-14",
        }
        pictures = {"CustomerLogo": "tests/test_files/monty-truth.png"}

        apply_tokens(prs, tokens)
        apply_pictures_by_name(prs, pictures)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        # -- second cycle: tokens are already gone; applying again is a --
        # -- no-op for text, and re-applies the same picture to confirm --
        # -- the shape is still locatable by name after the first cycle. --
        apply_tokens(prs2, tokens)
        apply_pictures_by_name(prs2, pictures)

        buf2 = io.BytesIO()
        prs2.save(buf2)
        buf2.seek(0)
        prs3 = Presentation(buf2)

        all_text = "\n".join(
            shape.text_frame.text
            for slide in prs3.slides
            for shape in iter_shapes(slide.shapes)
            if shape.has_text_frame
        )
        assert "Acme Corp" in all_text
        assert "{{ customer_name }}" not in all_text

        with open("tests/test_files/monty-truth.png", "rb") as f:
            expected = f.read()
        logo = next(
            s
            for slide in prs3.slides
            for s in iter_shapes(slide.shapes)
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.name == "CustomerLogo"
        )
        assert logo.image.blob == expected
