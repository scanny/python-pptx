"""Gherkin step implementations for presentation-level font embedding."""

from __future__ import annotations

import io
import zipfile

from behave import given, then, when
from behave.runner import Context
from helpers import test_pptx

from pptx import Presentation


@given("a presentation for font embedding")
def given_a_presentation_for_font_embedding(context: Context):
    """Alternative given kept distinct from the shared `given("a presentation")`.

    Currently unused — the shared `given("a presentation")` is re-used instead — but
    retained here so future scenarios can add a dedicated fixture if one becomes
    useful.
    """
    context.presentation = Presentation(test_pptx("prs-properties"))


@when('I embed a font file under the typeface "{typeface}"')
def when_embed_font_under_typeface(context: Context, typeface: str):
    stream = io.BytesIO(b"fake-font-bytes-for-acceptance-test")
    context.presentation.embed_font(stream, typeface)


@when('I embed a regular and a bold font file under the typeface "{typeface}"')
def when_embed_regular_and_bold(context: Context, typeface: str):
    context.presentation.embed_font(io.BytesIO(b"regular-font-bytes"), typeface, style="regular")
    context.presentation.embed_font(io.BytesIO(b"bold-font-bytes"), typeface, style="bold")


@then('Presentation.embedded_fonts contains "{typeface}"')
def then_embedded_fonts_contains(context: Context, typeface: str):
    assert (
        typeface in context.presentation.embedded_fonts
    ), "typeface %r not present in embedded_fonts %r" % (
        typeface,
        context.presentation.embedded_fonts,
    )


@then("the presentation has a ppt/fonts/font1.fntdata part after save")
def then_presentation_has_font_part_after_save(context: Context):
    stream = io.BytesIO()
    context.presentation.save(stream)
    stream.seek(0)
    with zipfile.ZipFile(stream) as z:
        names = z.namelist()
    assert (
        "ppt/fonts/font1.fntdata" in names
    ), "expected 'ppt/fonts/font1.fntdata' in package, got: %r" % [n for n in names if "font" in n]


@then("the p:embeddedFont entry has both p:regular and p:bold child elements")
def then_entry_has_regular_and_bold(context: Context):
    prs_elm = context.presentation._element
    embeddedFontLst = prs_elm.embeddedFontLst
    assert embeddedFontLst is not None
    entry = embeddedFontLst.embeddedFont_lst[0]
    assert entry.rId_for_style("regular") is not None
    assert entry.rId_for_style("bold") is not None
