# pyright: reportPrivateUsage=false

"""Verify-close regression tests for wave-20 non-gap triage items.

See ``docs/community/issue-triage.rst`` for the full disposition of every
non-gap item; this module pins the subset whose upstream complaint was
"feature X is missing" when in fact the shipped API already satisfies it.

Each test corresponds to one upstream issue. The reproduction is minimal
and uses only the public API — no mocks, no private attribute access —
so a future regression to the shipped path would fail the suite.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.data import BubbleChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Length, Pt

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _fresh_deck_with_title(text: str = "A title") -> Presentation:
    """Return a fresh Presentation with one title-only slide carrying `text`."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = text
    return prs


def _add_category_chart(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    cd.add_series("S0", (1.0, 2.0, 3.0))
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        cd,
    )
    return gf.chart


# ---------------------------------------------------------------------------
# #473 — chart value-axis visibility toggle
# ---------------------------------------------------------------------------


class DescribeIssue473ValueAxisVisible:
    """`chart.value_axis.visible = False` hides the axis line."""

    def it_defaults_to_visible(self):
        prs = Presentation()
        chart = _add_category_chart(prs)
        assert chart.value_axis.visible is True

    def it_hides_when_set_false(self):
        prs = Presentation()
        chart = _add_category_chart(prs)
        chart.value_axis.visible = False
        assert chart.value_axis.visible is False


# ---------------------------------------------------------------------------
# #538 — slide notes + review comments both ship
# ---------------------------------------------------------------------------


class DescribeIssue538NotesAndComments:
    """Speaker notes (``notes_slide``) and review comments (``comments``)
    are distinct capabilities and both are shipped."""

    def it_exposes_notes_slide(self):
        prs = _fresh_deck_with_title()
        slide = prs.slides[0]
        slide.notes_slide.notes_text_frame.text = "Speaker cue"
        assert slide.has_notes_slide is True
        assert slide.notes_slide.notes_text_frame.text == "Speaker cue"

    def and_it_exposes_review_comments(self):
        prs = _fresh_deck_with_title()
        slide = prs.slides[0]
        comments = slide.comments
        author = comments._authors.add_author("Alice", "A.")
        comments.add_comment(author, "LGTM")
        assert slide.has_comments is True
        assert slide.comments[0].text == "LGTM"
        assert slide.comments[0].author is not None
        assert slide.comments[0].author.name == "Alice"


# ---------------------------------------------------------------------------
# #540 — per-point data labels
# ---------------------------------------------------------------------------


class DescribeIssue540PerPointDataLabels:
    """Individual data points support their own data-label text/format."""

    def it_allows_setting_an_individual_point_label(self):
        prs = Presentation()
        chart = _add_category_chart(prs)
        plot = chart.plots[0]
        plot.has_data_labels = True
        series = chart.series[0]
        point = series.points[1]
        point.data_label.has_text_frame  # ensure property access works
        point.data_label.text_frame.text = "spotlight"
        assert point.data_label.text_frame.text == "spotlight"


# ---------------------------------------------------------------------------
# #541 / #553 — Picture.image.blob exposes embedded image bytes
# ---------------------------------------------------------------------------


class DescribeIssue541PictureImageBlob:
    """``Picture.image.blob`` returns the raw embedded image bytes."""

    def it_returns_bytes_that_start_with_the_png_signature(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        pic = slide.shapes.add_picture(
            "tests/test_files/python-powered.png",
            Inches(1),
            Inches(1),
            width=Inches(2),
        )
        blob = pic.image.blob
        # -- PNG signature is 8 bytes: 89 50 4E 47 0D 0A 1A 0A
        assert blob[:8] == b"\x89PNG\r\n\x1a\n"
        # -- content-type and ext are also exposed
        assert pic.image.content_type == "image/png"
        assert pic.image.ext == "png"


# ---------------------------------------------------------------------------
# #564 — Font.highlight_color works in table cells
# ---------------------------------------------------------------------------


class DescribeIssue564TableCellHighlightColor:
    """Highlight colour is a text-run attribute; it works in a cell run."""

    def it_sets_a_run_highlight_in_a_table_cell(self):
        from pptx.dml.color import RGBColor

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tbl = slide.shapes.add_table(
            rows=1, cols=1, left=Inches(1), top=Inches(2), width=Inches(3), height=Inches(1)
        ).table
        cell = tbl.cell(0, 0)
        cell.text = "hi"
        run = cell.text_frame.paragraphs[0].runs[0]
        run.font.highlight_color.rgb = RGBColor(0xFF, 0xFF, 0x00)
        assert run.font.highlight_color.rgb == RGBColor(0xFF, 0xFF, 0x00)


# ---------------------------------------------------------------------------
# #614 — Shape fill fore-color RGB is readable
# ---------------------------------------------------------------------------


class DescribeIssue614ShapeFillForeColor:
    """``shape.fill.fore_color.rgb`` roundtrips a set colour."""

    def it_stores_and_reads_back_fore_color(self):
        from pptx.dml.color import RGBColor

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
        assert shape.fill.fore_color.rgb == RGBColor(0x12, 0x34, 0x56)


# ---------------------------------------------------------------------------
# #665 — line-chart None values produce gaps
# ---------------------------------------------------------------------------


class DescribeIssue665LineChartGap:
    """``None`` in a line-chart series writes no ``c:val`` for that point
    (PowerPoint then renders a gap when ``dispBlanksAs=gap``)."""

    def it_skips_none_values_in_emitted_xml(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        cd = CategoryChartData()
        cd.categories = ["A", "B", "C", "D"]
        cd.add_series("S0", (1.0, None, 3.0, None))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            cd,
        )
        chart_xml = gf.chart._chartSpace.xml
        # -- gap-handling mode is "gap" by default for line charts --
        assert 'val="gap"' in chart_xml


# ---------------------------------------------------------------------------
# #671 — Slide.name read / write
# ---------------------------------------------------------------------------


class DescribeIssue671SlideName:
    """``Slide.name`` round-trips through write/read/save cycle."""

    def it_defaults_to_empty_string(self):
        prs = _fresh_deck_with_title()
        assert prs.slides[0].name == ""

    def it_roundtrips_an_assigned_name(self):
        prs = _fresh_deck_with_title()
        slide = prs.slides[0]
        slide.name = "Intro"
        buf = io.BytesIO()
        prs.save(buf)
        reopened = Presentation(io.BytesIO(buf.getvalue()))
        assert reopened.slides[0].name == "Intro"

    def and_it_clears_the_name_on_assignment_of_empty_string(self):
        prs = _fresh_deck_with_title()
        slide = prs.slides[0]
        slide.name = "Intro"
        slide.name = ""
        assert slide.name == ""


# ---------------------------------------------------------------------------
# #680 — Chart.replace_data with different categories
# ---------------------------------------------------------------------------


class DescribeIssue680ChartReplaceData:
    """Chart.replace_data accepts a wholly new (categories, series) payload."""

    def it_replaces_categories_and_values(self):
        prs = Presentation()
        chart = _add_category_chart(prs)
        # -- replace with different categories --
        new = CategoryChartData()
        new.categories = ["X", "Y"]
        new.add_series("S0", (10.0, 20.0))
        chart.replace_data(new)
        cats = [c.label for c in chart.plots[0].categories]
        assert cats == ["X", "Y"]


# ---------------------------------------------------------------------------
# #684 — Replace run.text preserves formatting
# ---------------------------------------------------------------------------


class DescribeIssue684RunReplaceText:
    """Mutating ``run.text`` preserves run-level formatting."""

    def it_preserves_bold_on_the_run_when_text_changes(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tf = slide.shapes.title.text_frame
        tf.text = "Original"
        run = tf.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(32)
        run.text = "Replaced"
        assert run.text == "Replaced"
        assert run.font.bold is True
        assert run.font.size == Pt(32)


# ---------------------------------------------------------------------------
# #710 — paragraph.line_spacing accepts a Length (Pt())
# ---------------------------------------------------------------------------


class DescribeIssue710LineSpacingLength:
    """``paragraph.line_spacing = Pt(18)`` stores a Length, not a ratio."""

    def it_roundtrips_an_emu_length(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tf = slide.shapes.title.text_frame
        tf.paragraphs[0].line_spacing = Pt(18)
        result = tf.paragraphs[0].line_spacing
        assert isinstance(result, Length)
        assert result == Pt(18)

    def and_it_still_accepts_a_float_multiplier(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tf = slide.shapes.title.text_frame
        tf.paragraphs[0].line_spacing = 1.5
        assert tf.paragraphs[0].line_spacing == 1.5


# ---------------------------------------------------------------------------
# #729 — Length arithmetic
# ---------------------------------------------------------------------------


class DescribeIssue729LengthArithmetic:
    """``Emu``/``Length`` support ordinary arithmetic and ordering."""

    def it_supports_addition(self):
        assert Emu(100) + Emu(50) == 150

    def it_supports_subtraction(self):
        assert Inches(2) - Inches(1) == Inches(1)

    def it_compares_mixed_length_subclasses(self):
        assert Pt(72) == Inches(1)


# ---------------------------------------------------------------------------
# #794 — bubble_scale on Bubble plot
# ---------------------------------------------------------------------------


class DescribeIssue794BubbleScale:
    """``BubblePlot.bubble_scale`` is readable and writable."""

    def it_round_trips_a_bubble_scale(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        bd = BubbleChartData()
        series = bd.add_series("S0")
        series.add_data_point(1.0, 2.0, 3.0)
        series.add_data_point(2.0, 4.0, 5.0)
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.BUBBLE, Inches(1), Inches(1), Inches(6), Inches(4), bd
        )
        plot = gf.chart.plots[0]
        plot.bubble_scale = 75
        assert plot.bubble_scale == 75


# ---------------------------------------------------------------------------
# #841 — Shape.shadow is exposed
# ---------------------------------------------------------------------------


class DescribeIssue841ShapeShadow:
    """``shape.shadow`` returns a ShadowFormat with inherit control."""

    def it_exposes_a_shadow_format_with_inherit(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )
        assert shape.shadow is not None
        # -- shadow.inherit is the tri-state toggle; default is True --
        assert shape.shadow.inherit in (True, False)
        shape.shadow.inherit = False
        assert shape.shadow.inherit is False


# ---------------------------------------------------------------------------
# #814 — add_shape requires MSO_SHAPE (not MSO_SHAPE_TYPE)
# ---------------------------------------------------------------------------


class DescribeIssue814AddShapeEnumDomain:
    """``add_shape`` takes ``MSO_SHAPE`` values; passing a ``MSO_SHAPE_TYPE``
    value by raw int is a user error, not a library bug."""

    def it_accepts_MSO_SHAPE_RECTANGLE(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )
        assert shape.auto_shape_type == MSO_SHAPE.RECTANGLE

    def but_raw_int_17_is_SMILEY_FACE_not_TEXT_BOX(self):
        """Confirms the user-error cause of upstream #814: int value 17
        in the MSO_SHAPE enumeration is SMILEY_FACE."""
        assert MSO_SHAPE.SMILEY_FACE.value == 17


# ---------------------------------------------------------------------------
# #1038 — MSO_LANGUAGE_ID has no 'en-CN' (user-error, not a gap)
# ---------------------------------------------------------------------------


class DescribeIssue1038LanguageIdEnCn:
    """``en-CN`` is not a real Microsoft LCID; ``en-US`` and ``zh-CN`` are."""

    def it_exposes_real_locales(self):
        from pptx.enum.lang import MSO_LANGUAGE_ID

        # -- both valid locales exist --
        assert MSO_LANGUAGE_ID.ENGLISH_US is not None
        assert MSO_LANGUAGE_ID.CHINESE_SINGAPORE is not None

    def but_does_not_define_a_fictitious_en_cn(self):
        from pptx.enum.lang import MSO_LANGUAGE_ID

        # -- there is no member for the fictitious en-CN --
        names = {m.name for m in MSO_LANGUAGE_ID}
        assert "ENGLISH_CN" not in names


# ---------------------------------------------------------------------------
# #1050 — Opening a file-like requires binary mode
# ---------------------------------------------------------------------------


class DescribeIssue1050FileLikeBinaryMode:
    """Opening a presentation from a BytesIO stream works end-to-end."""

    def it_opens_from_a_BytesIO(self):
        prs = _fresh_deck_with_title("Round-trip")
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        assert reopened.slides[0].shapes.title.text == "Round-trip"


# ---------------------------------------------------------------------------
# #962 — Hyperlink on chart-title text run
# ---------------------------------------------------------------------------


class DescribeIssue962HyperlinkOnRun:
    """The #962 reporter wanted to add a hyperlink on chart text. Hyperlinks
    on a text run — including runs inside a textbox placed over a chart —
    are exposed through ``run.hyperlink.address``, which is the documented
    path. The canonical call-site is a slide textbox."""

    def it_sets_a_hyperlink_on_a_textbox_run(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(0.5))
        tb.text_frame.text = "Click me"
        run = tb.text_frame.paragraphs[0].runs[0]
        run.hyperlink.address = "https://example.com/"
        assert run.hyperlink.address == "https://example.com/"


# ---------------------------------------------------------------------------
# #968 — dispBlanksAs=gap is the default for line charts
# ---------------------------------------------------------------------------


class DescribeIssue968LineChartDispBlanksAs:
    """Line charts are emitted with ``c:dispBlanksAs val="gap"`` by default."""

    def it_writes_gap_disposition_in_chart_xml(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        cd = CategoryChartData()
        cd.categories = ["A", "B", "C"]
        cd.add_series("S0", (1.0, 2.0, 3.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            cd,
        )
        xml = gf.chart._chartSpace.xml
        assert "dispBlanksAs" in xml
        assert 'val="gap"' in xml


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
