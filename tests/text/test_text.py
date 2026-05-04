# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.text.text` module."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

import pytest

from pptx.dml.color import ColorFormat, RGBColor
from pptx.dml.fill import FillFormat
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import (
    MSO_ANCHOR,
    MSO_AUTO_SIZE,
    MSO_UNDERLINE,
    PP_ALIGN,
    PP_AUTO_NUMBER,
    PP_AUTO_NUMBER_SCHEME,
)
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, MSO_STRIKE, MSO_UNDERLINE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.oxml.ns import qn
from pptx.shapes.autoshape import Shape
from pptx.text.text import (
    Font,
    TextFrame,
    TextFrameRect,
    _BulletFormat,
    _Field,
    _Hyperlink,
    _Paragraph,
    _Run,
)
from pptx.util import Inches, Pt

from ..oxml.unitdata.text import a_p, a_t, an_hlinkClick, an_r, an_rPr
from ..unitutil.cxml import element, xml
from ..unitutil.mock import (
    class_mock,
    instance_mock,
    loose_mock,
    method_mock,
    property_mock,
)

if TYPE_CHECKING:
    from pptx.oxml.text import CT_TextBody, CT_TextParagraph


class DescribeTextFrame(object):
    """Unit-test suite for `pptx.text.text.TextFrame` object."""

    def it_can_add_a_paragraph_to_itself(self, add_paragraph_fixture):
        text_frame, expected_xml = add_paragraph_fixture
        text_frame.add_paragraph()
        assert text_frame._txBody.xml == expected_xml

    def it_can_add_a_paragraph_with_text_in_one_call(self):
        """`TextFrame.add_paragraph(text=...)` spawns an `a:r` carrying `text` — issue #134."""
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)

        paragraph = text_frame.add_paragraph("hello")

        assert isinstance(paragraph, _Paragraph)
        assert paragraph.text == "hello"
        assert len(paragraph.runs) == 1
        assert paragraph.runs[0].text == "hello"

    def it_applies_font_kwargs_to_the_new_run_when_text_is_given(self):
        """Run-level kwargs apply to the newly-spawned run — issue #134."""
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)

        paragraph = text_frame.add_paragraph(
            "styled",
            bold=True,
            italic=True,
            size=Pt(24),
            color=RGBColor(0xAA, 0xBB, 0xCC),
            font_name="Calibri",
        )

        run = paragraph.runs[0]
        assert run.text == "styled"
        assert run.font.bold is True
        assert run.font.italic is True
        assert run.font.size == Pt(24)
        assert run.font.name == "Calibri"
        assert run.font.color.rgb == RGBColor(0xAA, 0xBB, 0xCC)

    def it_accepts_theme_color_as_color_kwarg(self):
        """A MSO_THEME_COLOR member sets theme_color on the new run — issue #134."""
        from pptx.enum.dml import MSO_THEME_COLOR

        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)

        paragraph = text_frame.add_paragraph("x", color=MSO_THEME_COLOR.ACCENT_1)

        assert paragraph.runs[0].font.color.theme_color == MSO_THEME_COLOR.ACCENT_1

    def it_raises_when_run_kwargs_given_without_text(self):
        """Passing a run-level kwarg without `text` raises ValueError — issue #134."""
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)

        with pytest.raises(ValueError, match="require `text`"):
            text_frame.add_paragraph(bold=True)

    def it_raises_when_color_kwarg_is_wrong_type(self):
        """An unsupported `color` type raises TypeError — issue #134."""
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)

        with pytest.raises(TypeError, match="RGBColor or MSO_THEME_COLOR"):
            text_frame.add_paragraph("x", color="red")  # pyright: ignore[reportArgumentType]

    def it_leaves_font_attrs_unset_when_kwargs_default_to_None(self):
        """`None` defaults don't mutate the new run's rPr — issue #134."""
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)

        paragraph = text_frame.add_paragraph("plain")

        run = paragraph.runs[0]
        assert run.font.bold is None
        assert run.font.italic is None
        assert run.font.size is None
        assert run.font.name is None

    def it_knows_its_autosize_setting(self, autosize_get_fixture):
        text_frame, expected_value = autosize_get_fixture
        assert text_frame.auto_size == expected_value

    @pytest.mark.parametrize(
        ("txBody_cxml", "value", "expected_cxml"),
        [
            ("p:txBody/a:bodyPr", MSO_AUTO_SIZE.NONE, "p:txBody/a:bodyPr/a:noAutofit"),
            (
                "p:txBody/a:bodyPr/a:noAutofit",
                MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT,
                "p:txBody/a:bodyPr/a:spAutoFit",
            ),
            (
                "p:txBody/a:bodyPr/a:spAutoFit",
                MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
                "p:txBody/a:bodyPr/a:normAutofit",
            ),
            ("p:txBody/a:bodyPr/a:normAutofit", None, "p:txBody/a:bodyPr"),
        ],
    )
    def it_can_change_its_autosize_setting(
        self, txBody_cxml: str, value: MSO_AUTO_SIZE | None, expected_cxml: str
    ):
        text_frame = TextFrame(element(txBody_cxml), None)
        text_frame.auto_size = value
        assert text_frame._txBody.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        "txBody_cxml",
        (
            "p:txBody/(a:p,a:p,a:p)",
            'p:txBody/a:p/a:r/a:t"foo"',
            'p:txBody/a:p/(a:br,a:r/a:t"foo")',
            'p:txBody/a:p/(a:fld,a:br,a:r/a:t"foo")',
        ),
    )
    def it_can_clear_itself_of_content(self, txBody_cxml):
        text_frame = TextFrame(element(txBody_cxml), None)
        text_frame.clear()
        assert text_frame._element.xml == xml("p:txBody/a:p")

    @pytest.mark.parametrize(
        ("txBody_cxml", "expected_value"),
        [
            ("p:txBody/a:bodyPr", 100.0),
            ("p:txBody/a:bodyPr/a:normAutofit", 100.0),
            ("p:txBody/a:bodyPr/a:normAutofit{fontScale=85000}", 85.0),
            ("p:txBody/a:bodyPr/a:normAutofit{fontScale=50000}", 50.0),
            ("p:txBody/a:bodyPr/a:spAutoFit", 100.0),
        ],
    )
    def it_knows_its_font_scale(self, txBody_cxml: str, expected_value: float):
        text_frame = TextFrame(element(txBody_cxml), None)
        assert text_frame.font_scale == expected_value

    @pytest.mark.parametrize(
        ("txBody_cxml", "value", "expected_cxml"),
        [
            # --adds normAutofit when not present--
            (
                "p:txBody/a:bodyPr",
                85.0,
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=85000}",
            ),
            # --replaces existing normAutofit fontScale--
            (
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=85000}",
                50.0,
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=50000}",
            ),
            # --preserves existing lnSpcReduction on normAutofit--
            (
                "p:txBody/a:bodyPr/a:normAutofit{lnSpcReduction=20000}",
                80.0,
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=80000,lnSpcReduction=20000}",
            ),
            # --replaces an existing spAutoFit choice sibling--
            (
                "p:txBody/a:bodyPr/a:spAutoFit",
                75.0,
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=75000}",
            ),
            # --assigning the default value of 100.0 clears the fontScale attr--
            (
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=85000}",
                100.0,
                "p:txBody/a:bodyPr/a:normAutofit",
            ),
        ],
    )
    def it_can_set_its_font_scale(self, txBody_cxml: str, value: float, expected_cxml: str):
        text_frame = TextFrame(element(txBody_cxml), None)
        text_frame.font_scale = value
        assert text_frame._txBody.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("txBody_cxml", "expected_value"),
        [
            ("p:txBody/a:bodyPr", 0.0),
            ("p:txBody/a:bodyPr/a:normAutofit", 0.0),
            ("p:txBody/a:bodyPr/a:normAutofit{lnSpcReduction=20000}", 20.0),
            ("p:txBody/a:bodyPr/a:noAutofit", 0.0),
        ],
    )
    def it_knows_its_line_space_reduction(self, txBody_cxml: str, expected_value: float):
        text_frame = TextFrame(element(txBody_cxml), None)
        assert text_frame.line_space_reduction == expected_value

    @pytest.mark.parametrize(
        ("txBody_cxml", "value", "expected_cxml"),
        [
            # --adds normAutofit when not present--
            (
                "p:txBody/a:bodyPr",
                20.0,
                "p:txBody/a:bodyPr/a:normAutofit{lnSpcReduction=20000}",
            ),
            # --preserves existing fontScale on normAutofit--
            (
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=80000}",
                10.0,
                "p:txBody/a:bodyPr/a:normAutofit{fontScale=80000,lnSpcReduction=10000}",
            ),
            # --replaces an existing noAutofit choice sibling--
            (
                "p:txBody/a:bodyPr/a:noAutofit",
                15.0,
                "p:txBody/a:bodyPr/a:normAutofit{lnSpcReduction=15000}",
            ),
            # --assigning the default value of 0.0 clears the lnSpcReduction attr--
            (
                "p:txBody/a:bodyPr/a:normAutofit{lnSpcReduction=20000}",
                0.0,
                "p:txBody/a:bodyPr/a:normAutofit",
            ),
        ],
    )
    def it_can_set_its_line_space_reduction(
        self, txBody_cxml: str, value: float, expected_cxml: str
    ):
        text_frame = TextFrame(element(txBody_cxml), None)
        text_frame.line_space_reduction = value
        assert text_frame._txBody.xml == xml(expected_cxml)

    def it_raises_on_font_scale_out_of_range(self):
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)
        with pytest.raises(ValueError):
            text_frame.font_scale = 0.5
        with pytest.raises(ValueError):
            text_frame.font_scale = 150.0

    def it_raises_on_line_space_reduction_out_of_range(self):
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)
        with pytest.raises(ValueError):
            text_frame.line_space_reduction = -1.0
        with pytest.raises(ValueError):
            text_frame.line_space_reduction = 150.0

    def it_knows_its_margin_settings(self, margin_get_fixture):
        text_frame, prop_name, unit, expected_value = margin_get_fixture
        margin_value = getattr(text_frame, prop_name)
        assert getattr(margin_value, unit) == expected_value

    def it_can_change_its_margin_settings(self, margin_set_fixture):
        text_frame, prop_name, new_value, expected_xml = margin_set_fixture
        setattr(text_frame, prop_name, new_value)
        assert text_frame._txBody.xml == expected_xml

    @pytest.mark.parametrize(
        ("txBody_cxml", "expected_value"),
        [
            ("p:txBody/a:bodyPr", None),
            ("p:txBody/a:bodyPr{anchor=t}", MSO_ANCHOR.TOP),
            ("p:txBody/a:bodyPr{anchor=b}", MSO_ANCHOR.BOTTOM),
        ],
    )
    def it_knows_its_vertical_alignment(self, txBody_cxml: str, expected_value: MSO_ANCHOR | None):
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        assert text_frame.vertical_anchor == expected_value

    @pytest.mark.parametrize(
        ("txBody_cxml", "new_value", "expected_cxml"),
        [
            ("p:txBody/a:bodyPr", MSO_ANCHOR.TOP, "p:txBody/a:bodyPr{anchor=t}"),
            (
                "p:txBody/a:bodyPr{anchor=t}",
                MSO_ANCHOR.MIDDLE,
                "p:txBody/a:bodyPr{anchor=ctr}",
            ),
            (
                "p:txBody/a:bodyPr{anchor=ctr}",
                MSO_ANCHOR.BOTTOM,
                "p:txBody/a:bodyPr{anchor=b}",
            ),
            ("p:txBody/a:bodyPr{anchor=b}", None, "p:txBody/a:bodyPr"),
        ],
    )
    def it_can_change_its_vertical_alignment(
        self, txBody_cxml: str, new_value: MSO_ANCHOR | None, expected_cxml: str
    ):
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        text_frame.vertical_anchor = new_value
        assert text_frame._element.xml == xml(expected_cxml)

    def it_knows_its_word_wrap_setting(self, wrap_get_fixture):
        text_frame, expected_value = wrap_get_fixture
        assert text_frame.word_wrap == expected_value

    def it_can_change_its_word_wrap_setting(self, wrap_set_fixture):
        text_frame, new_value, expected_xml = wrap_set_fixture
        text_frame.word_wrap = new_value
        assert text_frame._element.xml == expected_xml

    def it_provides_access_to_its_paragraphs(self, paragraphs_fixture):
        text_frame, ps = paragraphs_fixture
        paragraphs = text_frame.paragraphs
        assert len(paragraphs) == len(ps)
        for idx, paragraph in enumerate(paragraphs):
            assert isinstance(paragraph, _Paragraph)
            assert paragraph._element is ps[idx]

    def it_raises_on_attempt_to_set_margin_to_non_int(self):
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)
        with pytest.raises(TypeError):
            text_frame.margin_bottom = "0.1"

    @pytest.mark.parametrize(
        ("txBody_cxml", "expected_value"),
        [
            # -- default (no rot attribute) is 0.0 --
            ("p:txBody/a:bodyPr", 0.0),
            # -- 45 degrees = 45 * 60000 = 2700000 --
            ("p:txBody/a:bodyPr{rot=2700000}", 45.0),
            # -- 90 degrees = 5400000 --
            ("p:txBody/a:bodyPr{rot=5400000}", 90.0),
            # -- 270 degrees = 16200000 --
            ("p:txBody/a:bodyPr{rot=16200000}", 270.0),
        ],
    )
    def it_knows_its_rotation(self, txBody_cxml: str, expected_value: float):
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        assert text_frame.rotation == expected_value

    @pytest.mark.parametrize(
        ("txBody_cxml", "new_value", "expected_cxml"),
        [
            # -- adds rot attribute when not present --
            ("p:txBody/a:bodyPr", 45, "p:txBody/a:bodyPr{rot=2700000}"),
            # -- updates existing rot attribute --
            (
                "p:txBody/a:bodyPr{rot=2700000}",
                90,
                "p:txBody/a:bodyPr{rot=5400000}",
            ),
            # -- negative rotation normalizes to positive equivalent --
            ("p:txBody/a:bodyPr", -45, "p:txBody/a:bodyPr{rot=18900000}"),
            # -- float values accepted (e.g. 45.5 degrees) --
            ("p:txBody/a:bodyPr", 45.5, "p:txBody/a:bodyPr{rot=2730000}"),
        ],
    )
    def it_can_change_its_rotation(self, txBody_cxml: str, new_value: float, expected_cxml: str):
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        text_frame.rotation = new_value
        assert text_frame._element.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("txBody_cxml", "expected_cxml"),
        [
            # -- assigning None on an element with @rot removes the attribute --
            ("p:txBody/a:bodyPr{rot=2700000}", "p:txBody/a:bodyPr"),
            # -- assigning None on an element without @rot is a no-op --
            ("p:txBody/a:bodyPr", "p:txBody/a:bodyPr"),
        ],
    )
    def it_clears_rotation_when_assigned_None(
        self, txBody_cxml: str, expected_cxml: str
    ):
        """Assigning `None` to `TextFrame.rotation` removes `@rot` — issue #485."""
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        text_frame.rotation = None
        assert text_frame._element.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("txBody_cxml", "expected_value"),
        [
            # -- default (no upright attribute) is False --
            ("p:txBody/a:bodyPr", False),
            # -- explicit upright=1 reads as True --
            ("p:txBody/a:bodyPr{upright=1}", True),
            # -- explicit upright=0 reads as False --
            ("p:txBody/a:bodyPr{upright=0}", False),
        ],
    )
    def it_knows_its_upright_setting(self, txBody_cxml: str, expected_value: bool):
        """`TextFrame.upright` reflects `a:bodyPr/@upright` — issue #485."""
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        assert text_frame.upright is expected_value

    @pytest.mark.parametrize(
        ("txBody_cxml", "new_value", "expected_cxml"),
        [
            # -- assigning True adds @upright="1" --
            ("p:txBody/a:bodyPr", True, "p:txBody/a:bodyPr{upright=1}"),
            # -- assigning False (the default) clears any explicit setting --
            ("p:txBody/a:bodyPr{upright=1}", False, "p:txBody/a:bodyPr"),
            # -- toggling False->True on a bare element adds the attribute --
            ("p:txBody/a:bodyPr", False, "p:txBody/a:bodyPr"),
            # -- toggling True on an element already carrying the attribute
            #    is idempotent (still @upright="1")
            (
                "p:txBody/a:bodyPr{upright=1}",
                True,
                "p:txBody/a:bodyPr{upright=1}",
            ),
        ],
    )
    def it_can_change_its_upright_setting(
        self, txBody_cxml: str, new_value: bool, expected_cxml: str
    ):
        """Assignment updates the `a:bodyPr/@upright` attribute — issue #485."""
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        text_frame.upright = new_value
        assert text_frame._element.xml == xml(expected_cxml)

    def it_knows_the_part_it_belongs_to(self, text_frame_with_parent_):
        text_frame, parent_ = text_frame_with_parent_
        part = text_frame.part
        assert part is parent_.part

    def it_knows_what_text_it_contains(self, request, text_get_fixture, paragraphs_prop_):
        paragraph_texts, expected_value = text_get_fixture
        paragraphs_prop_.return_value = tuple(
            instance_mock(request, _Paragraph, text=text) for text in paragraph_texts
        )
        text_frame = TextFrame(None, None)

        text = text_frame.text

        assert text == expected_value

    def it_can_replace_the_text_it_contains(self, text_set_fixture):
        txBody, text, expected_xml = text_set_fixture
        text_frame = TextFrame(txBody, None)

        text_frame.text = text

        assert text_frame._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("txBody_cxml", "find", "replace", "expected_count", "expected_cxml"),
        [
            # -- match in a single run --
            (
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"Hello {NAME}!")',
                "{NAME}",
                "Alice",
                1,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"Hello Alice!")',
            ),
            # -- match split across two adjacent runs (the issue #836 case) --
            (
                'p:txBody/(a:bodyPr,a:p/(a:r/a:t"{NA",a:r/a:t"ME}"))',
                "{NAME}",
                "Alice",
                1,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"Alice")',
            ),
            # -- match across paragraphs does NOT cross the paragraph boundary --
            (
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"{X}",a:p/(a:r/a:t"{NA",a:r/a:t"ME}"))',
                "{NAME}",
                "Bob",
                1,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"{X}",a:p/a:r/a:t"Bob")',
            ),
            # -- multiple matches, one entirely in one run, one spanning two --
            (
                'p:txBody/(a:bodyPr,a:p/(a:r/a:t"{X} mid {",a:r/a:t"X}"))',
                "{X}",
                "Y",
                2,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"Y mid Y")',
            ),
            # -- no match → no change, count 0 --
            (
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"Hello")',
                "NotFound",
                "X",
                0,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"Hello")',
            ),
        ],
    )
    def it_can_replace_text_across_multiple_runs(
        self,
        txBody_cxml: str,
        find: str,
        replace: str,
        expected_count: int,
        expected_cxml: str,
    ):
        text_frame = TextFrame(cast("CT_TextBody", element(txBody_cxml)), None)
        assert text_frame.replace_text(find, replace) == expected_count
        assert text_frame._element.xml == xml(expected_cxml)

    def it_can_resize_its_text_to_best_fit(self, request, text_prop_):
        family, max_size, bold, italic, font_file, font_size = (
            "Family",
            42,
            "bold",
            "italic",
            "font_file",
            21,
        )
        text_prop_.return_value = "some text"
        _best_fit_font_size_ = method_mock(
            request, TextFrame, "_best_fit_font_size", return_value=font_size
        )
        _apply_fit_ = method_mock(request, TextFrame, "_apply_fit")
        text_frame = TextFrame(None, None)

        text_frame.fit_text(family, max_size, bold, italic, font_file)

        _best_fit_font_size_.assert_called_once_with(
            text_frame, family, max_size, bold, italic, font_file
        )
        _apply_fit_.assert_called_once_with(text_frame, family, font_size, bold, italic)

    def it_calculates_its_best_fit_font_size_to_help_fit_text(self, size_font_fixture):
        text_frame, family, max_size, bold, italic = size_font_fixture[:5]
        FontFiles_, TextFitter_, text, extents = size_font_fixture[5:9]
        font_file_, font_size_ = size_font_fixture[9:]

        font_size = text_frame._best_fit_font_size(family, max_size, bold, italic, None)

        FontFiles_.find.assert_called_once_with(family, bold, italic)
        TextFitter_.best_fit_font_size.assert_called_once_with(text, extents, max_size, font_file_)
        assert font_size is font_size_

    def it_calculates_its_effective_size_to_help_fit_text(self):
        sp_cxml = (
            "p:sp/(p:spPr/a:xfrm/(a:off{x=914400,y=914400},a:ext{cx=914400,c"
            "y=914400}),p:txBody/(a:bodyPr,a:p))"
        )
        text_frame = Shape(element(sp_cxml), None).text_frame
        assert text_frame._extents == (731520, 822960)

    def it_applies_fit_to_help_fit_text(self, request):
        family, font_size, bold, italic = "Family", 42, True, False
        _set_font_ = method_mock(request, TextFrame, "_set_font")
        text_frame = TextFrame(element("p:txBody/a:bodyPr"), None)

        text_frame._apply_fit(family, font_size, bold, italic)

        assert text_frame.auto_size is MSO_AUTO_SIZE.NONE
        assert text_frame.word_wrap is True
        _set_font_.assert_called_once_with(text_frame, family, font_size, bold, italic)

    def it_sets_its_font_to_help_fit_text(self, set_font_fixture):
        text_frame, family, size, bold, italic, expected_xml = set_font_fixture
        text_frame._set_font(family, size, bold, italic)
        assert text_frame._element.xml == expected_xml

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[
            ("p:txBody/a:bodyPr", "p:txBody/(a:bodyPr,a:p)"),
            ("p:txBody/(a:bodyPr,a:p)", "p:txBody/(a:bodyPr,a:p,a:p)"),
        ]
    )
    def add_paragraph_fixture(self, request):
        txBody_cxml, expected_cxml = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        expected_xml = xml(expected_cxml)
        return text_frame, expected_xml

    @pytest.fixture(
        params=[
            ("p:txBody/a:bodyPr", None),
            ("p:txBody/a:bodyPr/a:noAutofit", MSO_AUTO_SIZE.NONE),
            ("p:txBody/a:bodyPr/a:spAutoFit", MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT),
            ("p:txBody/a:bodyPr/a:normAutofit", MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE),
        ]
    )
    def autosize_get_fixture(self, request):
        txBody_cxml, expected_value = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        return text_frame, expected_value

    @pytest.fixture(
        params=[
            ("p:txBody/a:bodyPr", "left", "emu", Inches(0.1)),
            ("p:txBody/a:bodyPr", "top", "emu", Inches(0.05)),
            ("p:txBody/a:bodyPr", "right", "emu", Inches(0.1)),
            ("p:txBody/a:bodyPr", "bottom", "emu", Inches(0.05)),
            ("p:txBody/a:bodyPr{lIns=9144}", "left", "cm", 0.0254),
            ("p:txBody/a:bodyPr{tIns=18288}", "top", "mm", 0.508),
            ("p:txBody/a:bodyPr{rIns=76200}", "right", "pt", 6.0),
            ("p:txBody/a:bodyPr{bIns=36576}", "bottom", "inches", 0.04),
        ]
    )
    def margin_get_fixture(self, request):
        txBody_cxml, side, unit, expected_value = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        prop_name = "margin_%s" % side
        return text_frame, prop_name, unit, expected_value

    @pytest.fixture(
        params=[
            (
                "p:txBody/a:bodyPr",
                "left",
                Inches(0.11),
                "p:txBody/a:bodyPr{lIns=100584}",
            ),
            (
                "p:txBody/a:bodyPr{tIns=1234}",
                "top",
                Inches(0.12),
                "p:txBody/a:bodyPr{tIns=109728}",
            ),
            (
                "p:txBody/a:bodyPr{rIns=2345}",
                "right",
                Inches(0.13),
                "p:txBody/a:bodyPr{rIns=118872}",
            ),
            (
                "p:txBody/a:bodyPr{bIns=3456}",
                "bottom",
                Inches(0.14),
                "p:txBody/a:bodyPr{bIns=128016}",
            ),
            ("p:txBody/a:bodyPr", "left", Inches(0.1), "p:txBody/a:bodyPr"),
            ("p:txBody/a:bodyPr", "top", Inches(0.05), "p:txBody/a:bodyPr"),
            ("p:txBody/a:bodyPr", "right", Inches(0.1), "p:txBody/a:bodyPr"),
            ("p:txBody/a:bodyPr", "bottom", Inches(0.05), "p:txBody/a:bodyPr"),
        ]
    )
    def margin_set_fixture(self, request):
        txBody_cxml, side, new_value, expected_txBody_cxml = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        prop_name = "margin_%s" % side
        expected_xml = xml(expected_txBody_cxml)
        return text_frame, prop_name, new_value, expected_xml

    @pytest.fixture(params=["p:txBody", "p:txBody/a:p", "p:txBody/(a:p,a:p)"])
    def paragraphs_fixture(self, request):
        txBody_cxml = request.param
        txBody = element(txBody_cxml)
        text_frame = TextFrame(txBody, None)
        ps = txBody.xpath(".//a:p")
        return text_frame, ps

    @pytest.fixture(
        params=[
            (
                "p:txBody/(a:bodyPr,a:p/a:r)",
                True,
                False,
                "p:txBody/(a:bodyPr,a:p/(a:r/a:rPr{sz=600,b=1,i=0}/a:latin{typeface"
                "=F},a:endParaRPr{sz=600,b=1,i=0}/a:latin{typeface=F}))",
            ),
            (
                "p:txBody/a:p/a:br",
                True,
                False,
                "p:txBody/a:p/(a:br/a:rPr{sz=600,b=1,i=0}/a:latin{typeface=F},a:end"
                "ParaRPr{sz=600,b=1,i=0}/a:latin{typeface=F})",
            ),
            (
                "p:txBody/a:p/a:fld",
                True,
                False,
                "p:txBody/a:p/(a:fld/a:rPr{sz=600,b=1,i=0}/a:latin{typeface=F},a:en"
                "dParaRPr{sz=600,b=1,i=0}/a:latin{typeface=F})",
            ),
        ]
    )
    def set_font_fixture(self, request):
        txBody_cxml, bold, italic, expected_cxml = request.param
        family, size = "F", 6
        text_frame = TextFrame(element(txBody_cxml), None)
        expected_xml = xml(expected_cxml)
        return text_frame, family, size, bold, italic, expected_xml

    @pytest.fixture
    def size_font_fixture(self, FontFiles_, TextFitter_, text_prop_, _extents_prop_):
        text_frame = TextFrame(None, None)
        family, max_size, bold, italic = "Family", 42, True, False
        text, extents, font_size, font_file = "text", (111, 222), 21, "f.ttf"
        text_prop_.return_value = text
        _extents_prop_.return_value = extents
        FontFiles_.find.return_value = font_file
        TextFitter_.best_fit_font_size.return_value = font_size
        return (
            text_frame,
            family,
            max_size,
            bold,
            italic,
            FontFiles_,
            TextFitter_,
            text,
            extents,
            font_file,
            font_size,
        )

    @pytest.fixture(params=[(["foobar"], "foobar"), (["foo", "bar", "baz"], "foo\nbar\nbaz")])
    def text_get_fixture(self, request):
        paragraph_texts, expected_value = request.param
        return paragraph_texts, expected_value

    @pytest.fixture(
        params=[
            # ---empty to something---
            ("p:txBody/a:p", "foobar", 'p:txBody/a:p/a:r/a:t"foobar"'),
            # ---something to something else---
            ('p:txBody/a:p/a:r/a:t"foobar"', "barfoo", 'p:txBody/a:p/a:r/a:t"barfoo"'),
            # ---single paragraph to multiple---
            (
                'p:txBody/a:p/a:r/a:t"barfoo"',
                "foo\nbar",
                'p:txBody/(a:p/a:r/a:t"foo",a:p/a:r/a:t"bar")',
            ),
            # ---multiple paragraphs to single---
            (
                'p:txBody/(a:p/a:r/a:t"foo",a:p/a:r/a:t"bar")',
                "barfoo",
                'p:txBody/a:p/a:r/a:t"barfoo"',
            ),
            # ---something to empty---
            ('p:txBody/a:p/a:r/a:t"foobar"', "", "p:txBody/a:p"),
            # ---vertical-tab becomes line-break---
            ("p:txBody/a:p", "a\vb", 'p:txBody/a:p/(a:r/a:t"a",a:br,a:r/a:t"b")'),
        ]
    )
    def text_set_fixture(self, request):
        txBody_cxml, text, expected_cxml = request.param
        txBody = element(txBody_cxml)
        expected_xml = xml(expected_cxml)
        return txBody, text, expected_xml

    @pytest.fixture(
        params=[
            ("p:txBody/a:bodyPr", None),
            ("p:txBody/a:bodyPr{wrap=square}", True),
            ("p:txBody/a:bodyPr{wrap=none}", False),
        ]
    )
    def wrap_get_fixture(self, request):
        txBody_cxml, expected_value = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        return text_frame, expected_value

    @pytest.fixture(
        params=[
            ("p:txBody/a:bodyPr", True, "p:txBody/a:bodyPr{wrap=square}"),
            ("p:txBody/a:bodyPr{wrap=square}", False, "p:txBody/a:bodyPr{wrap=none}"),
            ("p:txBody/a:bodyPr{wrap=none}", None, "p:txBody/a:bodyPr"),
        ]
    )
    def wrap_set_fixture(self, request):
        txBody_cxml, new_value, expected_txBody_cxml = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        expected_xml = xml(expected_txBody_cxml)
        return text_frame, new_value, expected_xml

    # fixture components -----------------------------------

    @pytest.fixture
    def _extents_prop_(self, request):
        return property_mock(request, TextFrame, "_extents")

    @pytest.fixture
    def FontFiles_(self, request):
        return class_mock(request, "pptx.text.text.FontFiles")

    @pytest.fixture
    def paragraphs_prop_(self, request):
        return property_mock(request, TextFrame, "paragraphs")

    @pytest.fixture
    def TextFitter_(self, request):
        return class_mock(request, "pptx.text.text.TextFitter")

    @pytest.fixture
    def text_frame_with_parent_(self, request):
        parent_ = loose_mock(request, name="parent_")
        text_frame = TextFrame(None, parent_)
        return text_frame, parent_

    @pytest.fixture
    def text_prop_(self, request):
        return property_mock(request, TextFrame, "text")


class DescribeTextFrameRect(object):
    """Unit-test suite for `pptx.text.text.TextFrameRect` namedtuple."""

    def it_exposes_its_fields_by_name_and_position(self):
        from pptx.util import Emu

        rect = TextFrameRect(left=Emu(100), top=Emu(200), width=Emu(300), height=Emu(400))

        # -- field access
        assert rect.left == 100
        assert rect.top == 200
        assert rect.width == 300
        assert rect.height == 400

        # -- tuple unpacking preserves (left, top, width, height) order
        left, top, width, height = rect
        assert (left, top, width, height) == (100, 200, 300, 400)

    def it_returns_Length_values_that_support_unit_conversion(self):
        from pptx.util import Length

        rect = TextFrameRect(left=Inches(1), top=Inches(2), width=Inches(3), height=Inches(4))

        assert isinstance(rect.left, Length)
        assert rect.left.inches == 1
        assert rect.top.inches == 2
        assert rect.width.inches == 3
        assert rect.height.inches == 4


class DescribeFont(object):
    """Unit-test suite for `pptx.text.text.Font` object."""

    def it_knows_its_bold_setting(self, bold_get_fixture):
        font, expected_value = bold_get_fixture
        assert font.bold == expected_value

    def it_can_change_its_bold_setting(self, bold_set_fixture):
        font, new_value, expected_xml = bold_set_fixture
        font.bold = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_italic_setting(self, italic_get_fixture):
        font, expected_value = italic_get_fixture
        assert font.italic == expected_value

    def it_can_change_its_italic_setting(self, italic_set_fixture):
        font, new_value, expected_xml = italic_set_fixture
        font.italic = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_language_id(self, language_id_get_fixture):
        font, expected_value = language_id_get_fixture
        assert font.language_id == expected_value

    def it_can_change_its_language_id_setting(self, language_id_set_fixture):
        font, new_value, expected_xml = language_id_set_fixture
        font.language_id = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_underline_setting(self, underline_get_fixture):
        font, expected_value = underline_get_fixture
        assert font.underline is expected_value, "got %s" % font.underline

    def it_can_change_its_underline_setting(self, underline_set_fixture):
        font, new_value, expected_xml = underline_set_fixture
        font.underline = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_strikethrough_setting(self, strikethrough_get_fixture):
        font, expected_value = strikethrough_get_fixture
        assert font.strikethrough is expected_value, "got %s" % font.strikethrough

    def it_can_change_its_strikethrough_setting(self, strikethrough_set_fixture):
        font, new_value, expected_xml = strikethrough_set_fixture
        font.strikethrough = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_size(self, size_get_fixture):
        font, expected_value = size_get_fixture
        assert font.size == expected_value

    def it_can_change_its_size(self, size_set_fixture):
        font, new_value, expected_xml = size_set_fixture
        font.size = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_latin_typeface(self, name_get_fixture):
        font, expected_value = name_get_fixture
        assert font.name == expected_value

    def it_can_change_its_latin_typeface(self, name_set_fixture):
        font, new_value, expected_xml = name_set_fixture
        font.name = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_east_asian_typeface(self, name_ea_get_fixture):
        font, expected_value = name_ea_get_fixture
        assert font.name_ea == expected_value

    def it_can_change_its_east_asian_typeface(self, name_ea_set_fixture):
        font, new_value, expected_xml = name_ea_set_fixture
        font.name_ea = new_value
        assert font._element.xml == expected_xml

    def it_knows_its_complex_script_typeface(self, name_cs_get_fixture):
        font, expected_value = name_cs_get_fixture
        assert font.name_cs == expected_value

    def it_can_change_its_complex_script_typeface(self, name_cs_set_fixture):
        font, new_value, expected_xml = name_cs_set_fixture
        font.name_cs = new_value
        assert font._element.xml == expected_xml

    def it_preserves_other_font_slots_when_assigning_name(self):
        cxml = "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic},a:cs{typeface=Arial})"
        font = Font(element(cxml))
        font.name = "Times New Roman"
        assert font.name == "Times New Roman"
        assert font.name_ea == "MS Gothic"
        assert font.name_cs == "Arial"

    def it_preserves_other_font_slots_when_assigning_name_ea(self):
        cxml = "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic},a:cs{typeface=Arial})"
        font = Font(element(cxml))
        font.name_ea = "SimSun"
        assert font.name == "Calibri"
        assert font.name_ea == "SimSun"
        assert font.name_cs == "Arial"

    def it_preserves_other_font_slots_when_assigning_name_cs(self):
        cxml = "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic},a:cs{typeface=Arial})"
        font = Font(element(cxml))
        font.name_cs = "Mangal"
        assert font.name == "Calibri"
        assert font.name_ea == "MS Gothic"
        assert font.name_cs == "Mangal"

    def it_provides_access_to_its_color(self, font):
        assert isinstance(font.color, ColorFormat)

    def it_provides_access_to_its_fill(self, font):
        assert isinstance(font.fill, FillFormat)

    # -- highlight_color (issue #675) -----------------------------------

    def it_provides_a_ColorFormat_for_the_highlight_color(self):
        font = Font(element("a:rPr"))

        color = font.highlight_color

        assert isinstance(color, ColorFormat)
        # -- the accessor is a lazyproperty; same instance is returned --
        assert font.highlight_color is color

    def it_creates_highlight_on_access_and_allows_setting_rgb(self):
        rPr = element("a:rPr")
        font = Font(rPr)

        font.highlight_color.rgb = RGBColor(0xFF, 0xFF, 0x00)

        assert rPr.xml == xml("a:rPr/a:highlight/a:srgbClr{val=FFFF00}")

    def it_allows_setting_highlight_theme_color(self):
        from pptx.enum.dml import MSO_THEME_COLOR

        rPr = element("a:rPr")
        font = Font(rPr)

        font.highlight_color.theme_color = MSO_THEME_COLOR.ACCENT_1

        assert rPr.xml == xml("a:rPr/a:highlight/a:schemeClr{val=accent1}")

    def it_reads_an_existing_highlight_color(self):
        rPr = element("a:rPr/a:highlight/a:srgbClr{val=ABCDEF}")
        font = Font(rPr)

        assert font.highlight_color.rgb == RGBColor(0xAB, 0xCD, 0xEF)

    def it_can_clear_its_highlight_color(self):
        rPr = element("a:rPr/a:highlight/a:srgbClr{val=FF0000}")
        font = Font(rPr)

        return_value = font.clear_highlight_color()

        assert rPr.xml == xml("a:rPr")
        assert return_value is font

    def it_is_a_no_op_to_clear_highlight_when_none_is_present(self):
        rPr = element("a:rPr")
        font = Font(rPr)

        font.clear_highlight_color()

        assert rPr.xml == xml("a:rPr")

    def it_preserves_sibling_rPr_children_when_setting_highlight(self):
        cxml = "a:rPr/(a:solidFill/a:srgbClr{val=112233},a:latin{typeface=Calibri})"
        rPr = element(cxml)
        font = Font(rPr)

        font.highlight_color.rgb = RGBColor(0xFF, 0xFF, 0x00)

        # -- highlight lands between solidFill and latin per spec order --
        expected = xml(
            "a:rPr/(a:solidFill/a:srgbClr{val=112233},"
            "a:highlight/a:srgbClr{val=FFFF00},a:latin{typeface=Calibri})"
        )
        assert rPr.xml == expected

    # -- shadow / effect_format (issue #546) ----------------------------

    def it_provides_access_to_its_shadow(self):
        from pptx.dml.effect import ShadowFormat

        font = Font(element("a:rPr"))
        assert isinstance(font.shadow, ShadowFormat)
        assert font.shadow.inherit is True

    def it_writes_shadow_xml_under_rPr_effectLst_outerShdw(self):
        from pptx.util import Emu

        font = Font(element("a:rPr"))
        font.shadow.blur_radius = Emu(50800)
        font.shadow.distance = Emu(38100)
        font.shadow.direction = 45.0

        assert font.shadow.blur_radius == 50800
        assert font.shadow.distance == 38100
        assert font.shadow.direction == 45.0
        assert font.shadow.inherit is False
        # -- XML shape is a:rPr/a:effectLst/a:outerShdw --
        assert "effectLst" in font._element.xml
        assert "outerShdw" in font._element.xml

    def it_restores_inheritance_when_shadow_inherit_set_True(self):
        from pptx.util import Emu

        font = Font(element("a:rPr"))
        font.shadow.blur_radius = Emu(50800)
        assert font.shadow.inherit is False

        font.shadow.inherit = True

        assert font.shadow.inherit is True
        assert font.shadow.blur_radius is None

    def it_provides_access_to_its_effect_format(self):
        from pptx.dml.effect import EffectFormat

        font = Font(element("a:rPr"))
        assert isinstance(font.effect_format, EffectFormat)

    # -- effective_color (issue #938) -----------------------------------

    def it_returns_None_for_effective_color_without_part_context(self):
        # -- Font constructed with no parent can't walk inheritance --
        font = Font(element("a:rPr"))
        assert font.effective_color is None

    def it_resolves_effective_color_from_run_rPr_solidFill(self):
        from pptx.oxml import parse_xml

        rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:solidFill><a:srgbClr val="FF6600"/></a:solidFill>'
            "</a:rPr>"
        )
        font = Font(parse_xml(rPr_xml))
        assert font.effective_color == RGBColor(0xFF, 0x66, 0x00)

    def it_resolves_effective_color_from_paragraph_defRPr(self):
        from pptx.oxml import parse_xml

        p_xml = (
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<a:pPr><a:defRPr>"
            '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>'
            "</a:defRPr></a:pPr>"
            "<a:r><a:rPr/><a:t>hi</a:t></a:r>"
            "</a:p>"
        )
        p = parse_xml(p_xml)
        rPr = p[1].rPr
        font = Font(rPr)
        assert font.effective_color == RGBColor(0x00, 0xFF, 0x00)

    def it_resolves_effective_color_from_lstStyle_level(self):
        from pptx.oxml import parse_xml

        txBody_xml = (
            '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<a:bodyPr/>"
            "<a:lstStyle>"
            "<a:lvl2pPr><a:defRPr>"
            '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>'
            "</a:defRPr></a:lvl2pPr>"
            "</a:lstStyle>"
            '<a:p><a:pPr lvl="1"/><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>'
            "</p:txBody>"
        )
        txBody = parse_xml(txBody_xml)
        rPr = txBody[2][1].rPr  # -- a:p -> a:r -> a:rPr --
        font = Font(rPr)
        assert font.effective_color == RGBColor(0x11, 0x22, 0x33)

    # -- effective_size / effective_bold / effective_italic / effective_name
    # -- (issue #378) ---------------------------------------------------

    def it_returns_None_for_effective_size_when_no_size_declared(self):
        # -- Font with bare rPr and no containing paragraph => no ancestor --
        font = Font(element("a:rPr"))
        assert font.effective_size is None

    def it_resolves_effective_size_from_run_rPr_sz(self):
        font = Font(element("a:rPr{sz=2400}"))
        assert font.effective_size == Pt(24)

    def it_resolves_effective_size_from_paragraph_defRPr(self):
        from pptx.oxml import parse_xml

        p_xml = (
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:pPr><a:defRPr sz="3600"/></a:pPr>'
            "<a:r><a:rPr/><a:t>hi</a:t></a:r>"
            "</a:p>"
        )
        p = parse_xml(p_xml)
        rPr = p[1].rPr
        font = Font(rPr)
        assert font.effective_size == Pt(36)

    def it_resolves_effective_size_from_lstStyle_level(self):
        from pptx.oxml import parse_xml

        txBody_xml = (
            '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<a:bodyPr/>"
            "<a:lstStyle>"
            '<a:lvl2pPr><a:defRPr sz="2800"/></a:lvl2pPr>'
            "</a:lstStyle>"
            '<a:p><a:pPr lvl="1"/><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>'
            "</p:txBody>"
        )
        txBody = parse_xml(txBody_xml)
        rPr = txBody[2][1].rPr  # -- a:p -> a:r -> a:rPr --
        font = Font(rPr)
        assert font.effective_size == Pt(28)

    def it_prefers_run_rPr_sz_over_paragraph_defRPr_sz(self):
        # -- run-level setting always wins --
        from pptx.oxml import parse_xml

        p_xml = (
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:pPr><a:defRPr sz="3600"/></a:pPr>'
            '<a:r><a:rPr sz="1800"/><a:t>hi</a:t></a:r>'
            "</a:p>"
        )
        p = parse_xml(p_xml)
        font = Font(p[1].rPr)
        assert font.effective_size == Pt(18)

    def it_returns_None_for_effective_bold_when_not_declared(self):
        font = Font(element("a:rPr"))
        assert font.effective_bold is None

    def it_resolves_effective_bold_from_run_rPr(self):
        font_true = Font(element("a:rPr{b=1}"))
        assert font_true.effective_bold is True

        font_false = Font(element("a:rPr{b=0}"))
        assert font_false.effective_bold is False

    def it_resolves_effective_bold_from_paragraph_defRPr(self):
        from pptx.oxml import parse_xml

        p_xml = (
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:pPr><a:defRPr b="1"/></a:pPr>'
            "<a:r><a:rPr/><a:t>hi</a:t></a:r>"
            "</a:p>"
        )
        p = parse_xml(p_xml)
        font = Font(p[1].rPr)
        assert font.effective_bold is True

    def it_returns_None_for_effective_italic_when_not_declared(self):
        font = Font(element("a:rPr"))
        assert font.effective_italic is None

    def it_resolves_effective_italic_from_paragraph_defRPr(self):
        from pptx.oxml import parse_xml

        p_xml = (
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:pPr><a:defRPr i="1"/></a:pPr>'
            "<a:r><a:rPr/><a:t>hi</a:t></a:r>"
            "</a:p>"
        )
        p = parse_xml(p_xml)
        font = Font(p[1].rPr)
        assert font.effective_italic is True

    def it_returns_None_for_effective_name_when_no_latin_declared(self):
        font = Font(element("a:rPr"))
        assert font.effective_name is None

    def it_resolves_effective_name_from_run_rPr(self):
        font = Font(element("a:rPr/a:latin{typeface=Arial}"))
        assert font.effective_name == "Arial"

    def it_resolves_effective_name_from_paragraph_defRPr(self):
        from pptx.oxml import parse_xml

        p_xml = (
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<a:pPr><a:defRPr>"
            '<a:latin typeface="Times New Roman"/>'
            "</a:defRPr></a:pPr>"
            "<a:r><a:rPr/><a:t>hi</a:t></a:r>"
            "</a:p>"
        )
        p = parse_xml(p_xml)
        font = Font(p[1].rPr)
        assert font.effective_name == "Times New Roman"

    def it_resolves_effective_name_from_lstStyle_level(self):
        from pptx.oxml import parse_xml

        txBody_xml = (
            '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<a:bodyPr/>"
            "<a:lstStyle>"
            '<a:lvl2pPr><a:defRPr><a:latin typeface="Verdana"/></a:defRPr></a:lvl2pPr>'
            "</a:lstStyle>"
            '<a:p><a:pPr lvl="1"/><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>'
            "</p:txBody>"
        )
        txBody = parse_xml(txBody_xml)
        rPr = txBody[2][1].rPr
        font = Font(rPr)
        assert font.effective_name == "Verdana"

    def it_resolves_effective_size_end_to_end_through_placeholder(self):
        """Round-trip: a new textbox with an unset run.font.size resolves
        through the inheritance chain to a real |Length|.

        Exercises the full path: run.font.size is None, but effective_size
        walks to the slide master's `p:bodyStyle/a:lvl1pPr/a:defRPr/@sz`.
        """
        from pptx import Presentation
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tf = box.text_frame
        tf.text = "hello"
        run = tf.paragraphs[0].runs[0]

        # -- the run itself carries no explicit size --
        assert run.font.size is None
        # -- but the effective size resolves via the chain (32pt in the
        # -- bundled default master bodyStyle's lvl1pPr/defRPr) --
        size = run.font.effective_size
        assert size is not None
        assert size == Pt(32)
        # -- Latin typeface resolves to the theme placeholder (+mn-lt) --
        assert run.font.effective_name == "+mn-lt"

    # -- use_theme_hyperlink_color (issue #940) -------------------------

    def it_returns_None_for_use_theme_hyperlink_color_when_no_hyperlink(self):
        # -- no `a:hlinkClick` on the run => no override marker possible --
        font = Font(element("a:rPr"))
        assert font.use_theme_hyperlink_color is None

    def it_returns_None_when_hlinkClick_has_no_override_marker(self):
        # -- an `a:hlinkClick` without the python-pptx marker => theme applies --
        font = Font(element("a:rPr/a:hlinkClick"))
        assert font.use_theme_hyperlink_color is None

    def it_returns_False_when_hlinkClick_carries_override_marker(self):
        rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:hlinkClick><a:extLst><a:ext uri="{PY-PPTX-940}"/></a:extLst>'
            "</a:hlinkClick></a:rPr>"
        )
        from pptx.oxml import parse_xml

        font = Font(parse_xml(rPr_xml))
        assert font.use_theme_hyperlink_color is False

    def it_can_opt_out_of_theme_hyperlink_color(self):
        font = Font(element("a:rPr/a:hlinkClick"))
        font.use_theme_hyperlink_color = False
        assert font.use_theme_hyperlink_color is False
        # -- marker is present exactly once --
        from pptx.oxml.ns import qn

        hlinkClick = font._element.find(qn("a:hlinkClick"))
        extLst = hlinkClick.find(qn("a:extLst"))
        assert extLst is not None
        exts = extLst.findall(qn("a:ext"))
        assert len(exts) == 1
        assert exts[0].get("uri") == "{PY-PPTX-940}"

    def it_is_idempotent_when_opting_out_twice(self):
        font = Font(element("a:rPr/a:hlinkClick"))
        font.use_theme_hyperlink_color = False
        font.use_theme_hyperlink_color = False
        from pptx.oxml.ns import qn

        hlinkClick = font._element.find(qn("a:hlinkClick"))
        exts = hlinkClick.find(qn("a:extLst")).findall(qn("a:ext"))
        assert len(exts) == 1

    def it_can_restore_theme_hyperlink_color_by_setting_True(self):
        font = Font(element("a:rPr/a:hlinkClick"))
        font.use_theme_hyperlink_color = False
        font.use_theme_hyperlink_color = True
        assert font.use_theme_hyperlink_color is None
        from pptx.oxml.ns import qn

        hlinkClick = font._element.find(qn("a:hlinkClick"))
        assert hlinkClick.find(qn("a:extLst")) is None

    def it_can_restore_theme_hyperlink_color_by_setting_None(self):
        font = Font(element("a:rPr/a:hlinkClick"))
        font.use_theme_hyperlink_color = False
        font.use_theme_hyperlink_color = None
        assert font.use_theme_hyperlink_color is None

    def it_is_a_no_op_to_opt_out_when_run_has_no_hyperlink(self):
        font = Font(element("a:rPr"))
        font.use_theme_hyperlink_color = False
        # -- still None, and rPr still has no children --
        assert font.use_theme_hyperlink_color is None
        assert len(font._element) == 0

    # -- copy_from (issue #566) -----------------------------------------

    def it_returns_self_from_copy_from_for_chaining(self):
        src = Font(element("a:rPr"))
        dst = Font(element("a:rPr"))

        result = dst.copy_from(src)

        assert result is dst

    def it_copies_bold_italic_underline_strike_and_size(self):
        from pptx.enum.text import MSO_STRIKE, MSO_UNDERLINE

        src = Font(element("a:rPr{b=1,i=1,u=dbl,strike=dblStrike,sz=2800}"))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst.bold is True
        assert dst.italic is True
        assert dst.underline is MSO_UNDERLINE.DOUBLE_LINE
        assert dst.strikethrough is MSO_STRIKE.DOUBLE_LINE
        assert dst.size == Pt(28)

    def it_copies_latin_ea_and_cs_typeface_names(self):
        src_cxml = (
            "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic},"
            "a:cs{typeface=Arial})"
        )
        src = Font(element(src_cxml))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst.name == "Calibri"
        assert dst.name_ea == "MS Gothic"
        assert dst.name_cs == "Arial"

    def it_copies_language_id(self):
        from pptx.enum.lang import MSO_LANGUAGE_ID

        src = Font(element("a:rPr{lang=pl-PL}"))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst.language_id == MSO_LANGUAGE_ID.POLISH

    def it_copies_rgb_color(self):
        src = Font(element("a:rPr/a:solidFill/a:srgbClr{val=FF6600}"))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst.color.rgb == RGBColor(0xFF, 0x66, 0x00)

    def it_copies_theme_color(self):
        from pptx.enum.dml import MSO_THEME_COLOR

        src = Font(element("a:rPr/a:solidFill/a:schemeClr{val=accent1}"))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst.color.theme_color == MSO_THEME_COLOR.ACCENT_1

    def it_preserves_lumMod_lumOff_when_copying_color(self):
        from pptx.oxml import parse_xml

        src_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:solidFill><a:schemeClr val="accent1">'
            '<a:lumMod val="75000"/><a:lumOff val="25000"/>'
            "</a:schemeClr></a:solidFill></a:rPr>"
        )
        src = Font(parse_xml(src_xml))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        # -- nested lumMod / lumOff survived the deep copy --
        assert dst.color.brightness == 0.25

    def it_copies_highlight_color(self):
        src = Font(element("a:rPr/a:highlight/a:srgbClr{val=FFFF00}"))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst.highlight_color.rgb == RGBColor(0xFF, 0xFF, 0x00)

    def it_copies_use_theme_hyperlink_color_marker(self):
        from pptx.oxml import parse_xml

        src_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:hlinkClick><a:extLst><a:ext uri="{PY-PPTX-940}"/></a:extLst>'
            "</a:hlinkClick></a:rPr>"
        )
        src = Font(parse_xml(src_xml))
        # -- dst must already carry an a:hlinkClick for the marker to stick --
        dst = Font(element("a:rPr/a:hlinkClick"))

        dst.copy_from(src)

        assert dst.use_theme_hyperlink_color is False

    def it_clears_dst_bold_when_src_has_none(self):
        src = Font(element("a:rPr"))
        dst = Font(element("a:rPr{b=1}"))

        dst.copy_from(src)

        assert dst.bold is None
        assert dst._element.xml == xml("a:rPr")

    def it_clears_dst_color_when_src_has_none(self):
        src = Font(element("a:rPr"))
        dst = Font(element("a:rPr/a:solidFill/a:srgbClr{val=FF0000}"))

        dst.copy_from(src)

        # -- color.type is None => no explicit color on the run --
        assert dst.color.type is None
        assert dst._element.xml == xml("a:rPr")

    def it_clears_dst_highlight_when_src_has_none(self):
        src = Font(element("a:rPr"))
        dst = Font(element("a:rPr/a:highlight/a:srgbClr{val=FFFF00}"))

        dst.copy_from(src)

        assert dst._element.find(qn("a:highlight")) is None

    def it_clears_dst_latin_ea_cs_when_src_has_none(self):
        src = Font(element("a:rPr"))
        cxml = (
            "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic},"
            "a:cs{typeface=Arial})"
        )
        dst = Font(element(cxml))

        dst.copy_from(src)

        assert dst.name is None
        assert dst.name_ea is None
        assert dst.name_cs is None

    def it_clears_dst_language_id_when_src_has_none(self):
        from pptx.enum.lang import MSO_LANGUAGE_ID

        src = Font(element("a:rPr"))
        dst = Font(element("a:rPr{lang=pl-PL}"))

        dst.copy_from(src)

        # -- public surface sentinels NONE for "no explicit setting" --
        assert dst.language_id == MSO_LANGUAGE_ID.NONE
        assert dst._rPr.lang is None

    def it_is_idempotent_when_copying_an_empty_font_onto_itself(self):
        src = Font(element("a:rPr"))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst._element.xml == xml("a:rPr")

    def it_yields_equivalent_xml_after_copy_from(self):
        from pptx.oxml import parse_xml

        src_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' b="1" i="1" sz="2400" lang="en-US">'
            '<a:solidFill><a:srgbClr val="2E74B5"/></a:solidFill>'
            '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>'
            '<a:latin typeface="Calibri"/>'
            "</a:rPr>"
        )
        src = Font(parse_xml(src_xml))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        # -- the two rPrs should be indistinguishable at the XML level --
        assert dst._element.xml == src._element.xml

    def it_overwrites_existing_color_when_copying(self):
        src = Font(element("a:rPr/a:solidFill/a:srgbClr{val=00FF00}"))
        dst = Font(element("a:rPr/a:solidFill/a:srgbClr{val=FF0000}"))

        dst.copy_from(src)

        assert dst.color.rgb == RGBColor(0x00, 0xFF, 0x00)

    def it_replaces_non_solid_fill_when_src_has_solid(self):
        from pptx.oxml import parse_xml

        src = Font(element("a:rPr/a:solidFill/a:srgbClr{val=00FF00}"))
        # -- dst carries a gradient fill; copy_from should replace it with
        # -- the src's solid fill so the public color surface matches.
        dst_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:gradFill><a:gsLst/></a:gradFill></a:rPr>'
        )
        dst = Font(parse_xml(dst_xml))

        dst.copy_from(src)

        assert dst.color.rgb == RGBColor(0x00, 0xFF, 0x00)
        assert dst._element.find(qn("a:gradFill")) is None

    def it_does_not_mutate_src_when_copying(self):
        src_cxml = (
            "a:rPr{b=1}/(a:solidFill/a:srgbClr{val=FF6600},a:latin{typeface=Calibri})"
        )
        src = Font(element(src_cxml))
        src_snapshot = src._element.xml
        dst = Font(element("a:rPr/a:latin{typeface=Arial}"))

        dst.copy_from(src)

        assert src._element.xml == src_snapshot

    def it_is_a_no_op_for_use_theme_hyperlink_color_when_dst_has_no_hlinkClick(self):
        # -- src carries the override marker but dst has no hyperlink --
        from pptx.oxml import parse_xml

        src_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:hlinkClick><a:extLst><a:ext uri="{PY-PPTX-940}"/></a:extLst>'
            "</a:hlinkClick></a:rPr>"
        )
        src = Font(parse_xml(src_xml))
        dst = Font(element("a:rPr"))

        dst.copy_from(src)

        assert dst.use_theme_hyperlink_color is None
        # -- no hlinkClick was created on dst (copy_from doesn't reach into
        # -- hyperlink relationships)
        assert dst._element.find(qn("a:hlinkClick")) is None

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[("a:rPr", None), ("a:rPr{b=0}", False), ("a:rPr{b=1}", True)])
    def bold_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", True, "a:rPr{b=1}"),
            ("a:rPr{b=1}", False, "a:rPr{b=0}"),
            ("a:rPr{b=0}", None, "a:rPr"),
        ]
    )
    def bold_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(params=[("a:rPr", None), ("a:rPr{i=0}", False), ("a:rPr{i=1}", True)])
    def italic_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", True, "a:rPr{i=1}"),
            ("a:rPr{i=1}", False, "a:rPr{i=0}"),
            ("a:rPr{i=0}", None, "a:rPr"),
        ]
    )
    def italic_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:rPr", MSO_LANGUAGE_ID.NONE),
            ("a:rPr{lang=pl-PL}", MSO_LANGUAGE_ID.POLISH),
            ("a:rPr{lang=de-AT}", MSO_LANGUAGE_ID.GERMAN_AUSTRIA),
            ("a:rPr{lang=fr-FR}", MSO_LANGUAGE_ID.FRENCH),
        ]
    )
    def language_id_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", MSO_LANGUAGE_ID.ZULU, "a:rPr{lang=zu-ZA}"),
            ("a:rPr{lang=zu-ZA}", MSO_LANGUAGE_ID.URDU, "a:rPr{lang=ur-PK}"),
            ("a:rPr{lang=ur-PK}", MSO_LANGUAGE_ID.NONE, "a:rPr"),
            ("a:rPr{lang=ur-PK}", None, "a:rPr"),
            ("a:rPr", MSO_LANGUAGE_ID.NONE, "a:rPr"),
            ("a:rPr", None, "a:rPr"),
        ]
    )
    def language_id_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(params=[("a:rPr", None), ("a:rPr/a:latin{typeface=Foobar}", "Foobar")])
    def name_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", "Foobar", "a:rPr/a:latin{typeface=Foobar}"),
            (
                "a:rPr/a:latin{typeface=Foobar}",
                "Barfoo",
                "a:rPr/a:latin{typeface=Barfoo}",
            ),
            ("a:rPr/a:latin{typeface=Barfoo}", None, "a:rPr"),
        ]
    )
    def name_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:rPr", None),
            ("a:rPr/a:ea{typeface=MS Gothic}", "MS Gothic"),
            ("a:rPr/a:ea{typeface=SimSun}", "SimSun"),
        ]
    )
    def name_ea_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", "MS Gothic", "a:rPr/a:ea{typeface=MS Gothic}"),
            (
                "a:rPr/a:ea{typeface=MS Gothic}",
                "SimSun",
                "a:rPr/a:ea{typeface=SimSun}",
            ),
            ("a:rPr/a:ea{typeface=SimSun}", None, "a:rPr"),
            # -- ea slot is inserted after latin in document order --
            (
                "a:rPr/a:latin{typeface=Calibri}",
                "MS Gothic",
                "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic})",
            ),
        ]
    )
    def name_ea_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:rPr", None),
            ("a:rPr/a:cs{typeface=Arial}", "Arial"),
            ("a:rPr/a:cs{typeface=Mangal}", "Mangal"),
        ]
    )
    def name_cs_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", "Arial", "a:rPr/a:cs{typeface=Arial}"),
            (
                "a:rPr/a:cs{typeface=Arial}",
                "Mangal",
                "a:rPr/a:cs{typeface=Mangal}",
            ),
            ("a:rPr/a:cs{typeface=Mangal}", None, "a:rPr"),
            # -- cs slot is inserted after latin and ea in document order --
            (
                "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic})",
                "Arial",
                "a:rPr/(a:latin{typeface=Calibri},a:ea{typeface=MS Gothic},a:cs{typeface=Arial})",
            ),
        ]
    )
    def name_cs_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(params=[("a:rPr", None), ("a:rPr{sz=2400}", 304800)])
    def size_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(params=[("a:rPr", Pt(24), "a:rPr{sz=2400}"), ("a:rPr{sz=2400}", None, "a:rPr")])
    def size_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:rPr", None),
            ("a:rPr{u=none}", False),
            ("a:rPr{u=sng}", True),
            ("a:rPr{u=dbl}", MSO_UNDERLINE.DOUBLE_LINE),
        ]
    )
    def underline_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", True, "a:rPr{u=sng}"),
            ("a:rPr{u=sng}", False, "a:rPr{u=none}"),
            ("a:rPr{u=none}", MSO_UNDERLINE.WAVY_LINE, "a:rPr{u=wavy}"),
            ("a:rPr{u=wavy}", MSO_UNDERLINE.NONE, "a:rPr{u=none}"),
            ("a:rPr{u=wavy}", None, "a:rPr"),
        ]
    )
    def underline_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:rPr", None),
            ("a:rPr{strike=noStrike}", False),
            ("a:rPr{strike=sngStrike}", True),
            ("a:rPr{strike=dblStrike}", MSO_STRIKE.DOUBLE_LINE),
        ]
    )
    def strikethrough_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[
            ("a:rPr", True, "a:rPr{strike=sngStrike}"),
            ("a:rPr{strike=sngStrike}", False, "a:rPr{strike=noStrike}"),
            (
                "a:rPr{strike=noStrike}",
                MSO_STRIKE.DOUBLE_LINE,
                "a:rPr{strike=dblStrike}",
            ),
            ("a:rPr{strike=dblStrike}", MSO_STRIKE.NONE, "a:rPr{strike=noStrike}"),
            ("a:rPr{strike=dblStrike}", None, "a:rPr"),
        ]
    )
    def strikethrough_set_fixture(self, request):
        rPr_cxml, new_value, expected_rPr_cxml = request.param
        font = Font(element(rPr_cxml))
        expected_xml = xml(expected_rPr_cxml)
        return font, new_value, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def font(self):
        return Font(element("a:rPr"))


class Describe_Hyperlink(object):
    """Unit-test suite for `pptx.text.text._Hyperlink` object."""

    def it_knows_the_target_url_of_the_hyperlink(self, hlink_with_url_):
        hlink, rId, url = hlink_with_url_
        assert hlink.address == url
        hlink.part.target_ref.assert_called_once_with(rId)

    def it_has_None_for_address_when_no_hyperlink_is_present(self, hlink):
        assert hlink.address is None

    def it_can_set_the_target_url(self, hlink, rPr_with_hlinkClick_xml, url):
        hlink.address = url
        # verify -----------------------
        hlink.part.relate_to.assert_called_once_with(url, RT.HYPERLINK, is_external=True)
        assert hlink._rPr.xml == rPr_with_hlinkClick_xml
        assert hlink.address == url

    def it_can_remove_the_hyperlink(self, remove_hlink_fixture_):
        hlink, rPr_xml, rId = remove_hlink_fixture_
        hlink.address = None
        assert hlink._rPr.xml == rPr_xml
        hlink.part.drop_rel.assert_called_once_with(rId)

    def it_should_remove_the_hyperlink_when_url_set_to_empty_string(self, remove_hlink_fixture_):
        hlink, rPr_xml, rId = remove_hlink_fixture_
        hlink.address = ""
        assert hlink._rPr.xml == rPr_xml
        hlink.part.drop_rel.assert_called_once_with(rId)

    def it_can_change_the_target_url(self, change_hlink_fixture_):
        # fixture ----------------------
        hlink, rId_existing, new_url, new_rPr_xml = change_hlink_fixture_
        # exercise ---------------------
        hlink.address = new_url
        # verify -----------------------
        assert hlink._rPr.xml == new_rPr_xml
        hlink.part.drop_rel.assert_called_once_with(rId_existing)
        hlink.part.relate_to.assert_called_once_with(new_url, RT.HYPERLINK, is_external=True)

    # -- target_slide property (issue #1077) ------------------------------

    def it_returns_None_target_slide_when_no_hyperlink_is_present(self, request):
        rPr = element("a:rPr")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part")
        assert hlink.target_slide is None

    def it_returns_None_target_slide_for_an_external_url_hyperlink(self, request):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId3}")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part")
        assert hlink.target_slide is None

    def it_returns_the_target_slide_for_a_slide_jump(self, request):
        from pptx.parts.slide import SlidePart
        from pptx.slide import Slide

        rPr = element("a:rPr/a:hlinkClick{r:id=rId7,action=ppaction://hlinksldjump}")
        hlink = _Hyperlink(rPr, None)
        slide_ = instance_mock(request, Slide)
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.slide = slide_
        part_ = instance_mock(request, XmlPart)
        part_.related_part.return_value = slide_part_
        property_mock(request, _Hyperlink, "part", return_value=part_)

        result = hlink.target_slide

        part_.related_part.assert_called_once_with("rId7")
        assert result is slide_

    def and_address_returns_None_for_a_slide_jump_hyperlink(self, request):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId7,action=ppaction://hlinksldjump}")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part")
        assert hlink.address is None

    def it_can_set_the_target_slide(self, request):
        from pptx.oxml.ns import qn
        from pptx.parts.slide import SlidePart
        from pptx.slide import Slide

        rPr = element("a:rPr/a:hlinkClick{r:foo=x}")  # -- pre-declare r: ns
        rPr.remove(rPr[0])
        hlink = _Hyperlink(rPr, None)
        slide_ = instance_mock(request, Slide)
        slide_part_ = instance_mock(request, SlidePart)
        slide_.part = slide_part_
        part_ = instance_mock(request, XmlPart)
        part_.relate_to.return_value = "rId9"
        property_mock(request, _Hyperlink, "part", return_value=part_)

        hlink.target_slide = slide_

        part_.relate_to.assert_called_once_with(slide_part_, RT.SLIDE)
        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None
        assert hlinkClick.get("action") == "ppaction://hlinksldjump"
        assert hlinkClick.get(qn("r:id")) == "rId9"

    def it_replaces_an_existing_url_when_setting_a_target_slide(self, request):
        from pptx.parts.slide import SlidePart
        from pptx.slide import Slide

        rPr = element("a:rPr/a:hlinkClick{r:id=rId3}")
        hlink = _Hyperlink(rPr, None)
        slide_ = instance_mock(request, Slide)
        slide_part_ = instance_mock(request, SlidePart)
        slide_.part = slide_part_
        part_ = instance_mock(request, XmlPart)
        part_.relate_to.return_value = "rId9"
        property_mock(request, _Hyperlink, "part", return_value=part_)

        hlink.target_slide = slide_

        part_.drop_rel.assert_called_once_with("rId3")
        part_.relate_to.assert_called_once_with(slide_part_, RT.SLIDE)
        assert hlink._rPr.xml == xml("a:rPr/a:hlinkClick{action=ppaction://hlinksldjump,r:id=rId9}")

    def it_clears_the_hlinkClick_when_target_slide_is_None(self, request):
        from pptx.oxml.ns import qn

        rPr = element("a:rPr/a:hlinkClick{r:id=rId7,action=ppaction://hlinksldjump}")
        hlink = _Hyperlink(rPr, None)
        part_ = instance_mock(request, XmlPart)
        property_mock(request, _Hyperlink, "part", return_value=part_)

        hlink.target_slide = None

        part_.drop_rel.assert_called_once_with("rId7")
        assert rPr.find(qn("a:hlinkClick")) is None

    def it_removes_the_target_slide_via_del(self, request):
        from pptx.oxml.ns import qn

        rPr = element("a:rPr/a:hlinkClick{r:id=rId7,action=ppaction://hlinksldjump}")
        hlink = _Hyperlink(rPr, None)
        part_ = instance_mock(request, XmlPart)
        property_mock(request, _Hyperlink, "part", return_value=part_)

        del hlink.target_slide

        part_.drop_rel.assert_called_once_with("rId7")
        assert rPr.find(qn("a:hlinkClick")) is None

    def it_is_a_noop_to_clear_target_slide_when_no_hyperlink_is_present(self, request):
        from pptx.oxml.ns import qn

        rPr = element("a:rPr")
        hlink = _Hyperlink(rPr, None)
        part_ = instance_mock(request, XmlPart)
        property_mock(request, _Hyperlink, "part", return_value=part_)

        hlink.target_slide = None

        part_.drop_rel.assert_not_called()
        assert rPr.find(qn("a:hlinkClick")) is None

    # -- screen_tip property (issue #455) ---------------------------------

    def it_returns_None_screen_tip_when_no_hyperlink_is_present(self):
        hlink = _Hyperlink(element("a:rPr"), None)
        assert hlink.screen_tip is None

    def it_returns_the_tooltip_when_set(self):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId1,tooltip=Click me}")
        hlink = _Hyperlink(rPr, None)
        assert hlink.screen_tip == "Click me"

    def it_returns_None_screen_tip_when_hlinkClick_has_no_tooltip(self):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId1}")
        hlink = _Hyperlink(rPr, None)
        assert hlink.screen_tip is None

    def it_can_set_the_screen_tip_creating_hlinkClick_if_needed(self):
        rPr = element("a:rPr")
        hlink = _Hyperlink(rPr, None)

        hlink.screen_tip = "Hover me"

        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None
        assert hlinkClick.get("tooltip") == "Hover me"

    def it_can_update_an_existing_tooltip_without_disturbing_the_URL(self):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId3,tooltip=old}")
        hlink = _Hyperlink(rPr, None)

        hlink.screen_tip = "new"

        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None
        assert hlinkClick.get("tooltip") == "new"
        assert hlinkClick.get(qn("r:id")) == "rId3"

    def it_removes_only_the_tooltip_attribute_when_screen_tip_set_to_None(self):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId3,tooltip=old}")
        hlink = _Hyperlink(rPr, None)

        hlink.screen_tip = None

        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None  # -- element preserved
        assert hlinkClick.get("tooltip") is None
        assert hlinkClick.get(qn("r:id")) == "rId3"  # -- URL intact

    def it_is_a_noop_to_clear_screen_tip_when_no_hyperlink_is_present(self):
        rPr = element("a:rPr")
        hlink = _Hyperlink(rPr, None)

        hlink.screen_tip = None

        assert rPr.find(qn("a:hlinkClick")) is None

    def it_treats_empty_string_as_clear_of_the_tooltip(self):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId3,tooltip=old}")
        hlink = _Hyperlink(rPr, None)

        hlink.screen_tip = ""

        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None
        assert hlinkClick.get("tooltip") is None

    # -- sound / set_sound / remove_sound (issue #455) --------------------

    def it_returns_None_sound_when_no_hyperlink_is_present(self):
        hlink = _Hyperlink(element("a:rPr"), None)
        assert hlink.sound is None

    def it_returns_None_sound_when_hlinkClick_has_no_snd(self):
        rPr = element("a:rPr/a:hlinkClick{r:id=rId1}")
        hlink = _Hyperlink(rPr, None)
        assert hlink.sound is None

    def it_provides_a_Sound_object_when_snd_is_present(self, request):
        from pptx.action import Sound
        from pptx.parts.slide import SlidePart

        rPr = element("a:rPr/a:hlinkClick{r:id=rId1}" "/a:snd{r:embed=rId2,name=applause.wav}")
        hlink = _Hyperlink(rPr, None)
        slide_part_ = instance_mock(request, SlidePart)
        property_mock(request, _Hyperlink, "part", return_value=slide_part_)

        sound = hlink.sound

        assert isinstance(sound, Sound)
        assert sound.rId == "rId2"
        assert sound.name == "applause.wav"

    def it_can_set_a_sound_from_a_file_like(self, request):
        import io as _io

        from pptx.parts.slide import SlidePart

        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.get_or_add_sound_media_part.return_value = "rId9"
        rPr = element("a:rPr{a:a=a,r:r=r}")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part", return_value=slide_part_)

        sound = hlink.set_sound(_io.BytesIO(b"RIFFWAV"), name="boing.wav")

        assert slide_part_.get_or_add_sound_media_part.call_count == 1
        assert sound.rId == "rId9"
        assert sound.name == "boing.wav"
        assert rPr.xml == xml("a:rPr{a:a=a,r:r=r}/a:hlinkClick/a:snd{r:embed=rId9,name=boing.wav}")

    def it_replaces_an_existing_sound_when_setting_a_new_one(self, request):
        import io as _io

        from pptx.parts.slide import SlidePart

        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.get_or_add_sound_media_part.return_value = "rId99"
        rPr = element("a:rPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId5,name=old.wav}")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part", return_value=slide_part_)

        sound = hlink.set_sound(_io.BytesIO(b"RIFFWAV"), name="new.wav")

        assert slide_part_.drop_rel.call_args_list[0].args == ("rId5",)
        assert sound.rId == "rId99"
        assert rPr.xml == xml("a:rPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId99,name=new.wav}")

    def it_accepts_an_Audio_instance_directly(self, request):
        from pptx.media import Audio
        from pptx.parts.slide import SlidePart

        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.get_or_add_sound_media_part.return_value = "rId8"
        audio = Audio.from_blob(b"RIFFWAV", None, "zap.wav")
        rPr = element("a:rPr")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part", return_value=slide_part_)

        hlink.set_sound(audio)

        args, _ = slide_part_.get_or_add_sound_media_part.call_args
        assert args[0] is audio

    def it_removes_a_sound_and_drops_the_audio_rel(self, request):
        from pptx.parts.slide import SlidePart

        slide_part_ = instance_mock(request, SlidePart)
        rPr = element("a:rPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId2,name=applause.wav}")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part", return_value=slide_part_)

        hlink.remove_sound()

        assert slide_part_.drop_rel.call_args_list[0].args == ("rId2",)
        assert rPr.xml == xml("a:rPr/a:hlinkClick{r:id=rId1}")

    def it_is_a_noop_to_remove_sound_when_no_hyperlink_is_present(self, request):
        from pptx.parts.slide import SlidePart

        slide_part_ = instance_mock(request, SlidePart)
        rPr = element("a:rPr")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part", return_value=slide_part_)

        hlink.remove_sound()

        slide_part_.drop_rel.assert_not_called()

    def it_is_a_noop_to_remove_sound_when_hlinkClick_has_no_snd(self, request):
        from pptx.parts.slide import SlidePart

        slide_part_ = instance_mock(request, SlidePart)
        rPr = element("a:rPr/a:hlinkClick{r:id=rId1}")
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part", return_value=slide_part_)

        hlink.remove_sound()

        slide_part_.drop_rel.assert_not_called()
        assert rPr.find(qn("a:hlinkClick")) is not None

    def it_drops_an_audio_rel_when_removing_a_hyperlink_carrying_sound(self, request):
        """Clearing the URL drops both the URL rel *and* any snd rel."""
        rPr = element("a:rPr/a:hlinkClick{r:id=rId3}" "/a:snd{r:embed=rId4,name=applause.wav}")
        hlink = _Hyperlink(rPr, None)
        part_ = instance_mock(request, XmlPart)
        property_mock(request, _Hyperlink, "part", return_value=part_)

        hlink.address = None

        drop_rel_args = [c.args[0] for c in part_.drop_rel.call_args_list]
        assert "rId3" in drop_rel_args
        assert "rId4" in drop_rel_args
        assert rPr.find(qn("a:hlinkClick")) is None

    # fixtures ---------------------------------------------

    @pytest.fixture
    def change_hlink_fixture_(self, request, hlink_with_hlinkClick, rId, rId_2, part_, url_2):
        hlinkClick_bldr = an_hlinkClick().with_rId(rId_2)
        new_rPr_xml = an_rPr().with_nsdecls("a", "r").with_child(hlinkClick_bldr).xml()
        part_.relate_to.return_value = rId_2
        property_mock(request, _Hyperlink, "part", return_value=part_)
        return hlink_with_hlinkClick, rId, url_2, new_rPr_xml

    @pytest.fixture
    def hlink(self, request, part_):
        rPr = an_rPr().with_nsdecls("a", "r").element
        hlink = _Hyperlink(rPr, None)
        property_mock(request, _Hyperlink, "part", return_value=part_)
        return hlink

    @pytest.fixture
    def hlink_with_hlinkClick(self, request, rPr_with_hlinkClick_bldr):
        rPr = rPr_with_hlinkClick_bldr.element
        return _Hyperlink(rPr, None)

    @pytest.fixture
    def hlink_with_url_(self, request, part_, hlink_with_hlinkClick, rId, url):
        property_mock(request, _Hyperlink, "part", return_value=part_)
        return hlink_with_hlinkClick, rId, url

    @pytest.fixture
    def part_(self, request, url, rId):
        """
        Mock Part instance suitable for patching into _Hyperlink.part
        property. It returns url for target_ref() and rId for relate_to().
        """
        part_ = instance_mock(request, XmlPart)
        part_.target_ref.return_value = url
        part_.relate_to.return_value = rId
        return part_

    @pytest.fixture
    def rId(self):
        return "rId2"

    @pytest.fixture
    def rId_2(self):
        return "rId6"

    @pytest.fixture
    def remove_hlink_fixture_(self, request, hlink_with_hlinkClick, rPr_xml, rId):
        property_mock(request, _Hyperlink, "part")
        return hlink_with_hlinkClick, rPr_xml, rId

    @pytest.fixture
    def rPr_with_hlinkClick_bldr(self, rId):
        hlinkClick_bldr = an_hlinkClick().with_rId(rId)
        rPr_bldr = an_rPr().with_nsdecls("a", "r").with_child(hlinkClick_bldr)
        return rPr_bldr

    @pytest.fixture
    def rPr_with_hlinkClick_xml(self, rPr_with_hlinkClick_bldr):
        return rPr_with_hlinkClick_bldr.xml()

    @pytest.fixture
    def rPr_xml(self):
        return an_rPr().with_nsdecls("a", "r").xml()

    @pytest.fixture
    def url(self):
        return "https://github.com/scanny/python-pptx"

    @pytest.fixture
    def url_2(self):
        return "https://pypi.python.org/pypi/python-pptx"


class Describe_Paragraph(object):
    """Unit test suite for pptx.text.text._Paragraph object."""

    def it_can_add_a_line_break(self, line_break_fixture):
        paragraph, expected_xml = line_break_fixture
        paragraph.add_line_break()
        assert paragraph._p.xml == expected_xml

    def it_can_add_a_run(self, paragraph, p_with_r_xml):
        run = paragraph.add_run()
        assert paragraph._p.xml == p_with_r_xml
        assert isinstance(run, _Run)

    def it_can_add_a_run_with_text_in_one_call(self, paragraph):
        """`_Paragraph.add_run(text=...)` sets the run's text — issue #134."""
        run = paragraph.add_run("payload")

        assert isinstance(run, _Run)
        assert run.text == "payload"

    def it_applies_font_kwargs_to_the_new_run(self, paragraph):
        """Run-level kwargs populate the new run's rPr — issue #134."""
        run = paragraph.add_run(
            "body",
            bold=True,
            italic=False,
            size=Pt(14),
            color=RGBColor(0x10, 0x20, 0x30),
            font_name="Arial",
        )

        assert run.text == "body"
        assert run.font.bold is True
        assert run.font.italic is False
        assert run.font.size == Pt(14)
        assert run.font.name == "Arial"
        assert run.font.color.rgb == RGBColor(0x10, 0x20, 0x30)

    def it_accepts_theme_color_on_add_run(self, paragraph):
        """`color=MSO_THEME_COLOR.X` sets theme_color on the new run — issue #134."""
        from pptx.enum.dml import MSO_THEME_COLOR

        run = paragraph.add_run("x", color=MSO_THEME_COLOR.ACCENT_2)

        assert run.font.color.theme_color == MSO_THEME_COLOR.ACCENT_2

    def it_accepts_font_kwargs_without_text_on_add_run(self, paragraph):
        """`add_run(bold=True)` applies formatting even with an empty run — issue #134."""
        run = paragraph.add_run(bold=True, italic=True)

        assert run.text == ""
        assert run.font.bold is True
        assert run.font.italic is True

    def it_accepts_int_emu_for_size_on_add_run(self, paragraph):
        """A plain int is treated as EMU, matching `Font.size` semantics — issue #134."""
        run = paragraph.add_run("x", size=Pt(20))

        # -- bare int (already EMU) is also accepted --
        run2 = paragraph.add_run("y", size=int(Pt(18)))
        assert run.font.size == Pt(20)
        assert run2.font.size == Pt(18)

    def it_raises_on_add_run_with_bad_color_type(self, paragraph):
        """An unsupported `color` type raises TypeError — issue #134."""
        with pytest.raises(TypeError, match="RGBColor or MSO_THEME_COLOR"):
            paragraph.add_run("x", color=42)  # pyright: ignore[reportArgumentType]

    def it_can_add_an_auto_refresh_field(self, paragraph):
        field = paragraph.add_field("slidenum", "#")

        assert isinstance(field, _Field)
        assert field.field_type == "slidenum"
        assert field.text == "#"
        # ---field is attached into the paragraph XML---
        assert len(paragraph._p.fld_lst) == 1
        fld = paragraph._p.fld_lst[0]
        assert fld.type == "slidenum"
        # ---auto-assigned id looks like a GUID in braces---
        assert fld.id.startswith("{") and fld.id.endswith("}") and len(fld.id) == 38

    def it_can_add_a_math_equation(self):
        """Regression test for issue #528 (insert OMML into a paragraph).

        Writes an ``mc:AlternateContent/mc:Choice[Requires="a14"]/a14:m`` scaffold around the
        caller's OMML fragment, alongside an ``mc:Fallback/a:r`` plain-text rendering. Any
        existing runs in the paragraph are preserved.
        """
        from pptx.oxml.ns import qn

        p = cast("CT_TextParagraph", element('a:p/a:r/a:t"x = "'))
        paragraph = _Paragraph(p, None)
        omml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>1+2</m:t></m:r>"
            "</m:oMath>"
        )

        paragraph.add_math_equation(omml)

        # -- the pre-existing run is preserved --
        assert len(paragraph._p.r_lst) == 1
        assert paragraph._p.r_lst[0].text == "x = "
        # -- exactly one mc:AlternateContent child is appended --
        ac_children = list(paragraph._p.iterchildren(qn("mc:AlternateContent")))
        assert len(ac_children) == 1
        ac = ac_children[0]
        # -- mc:Choice/a14:m/m:oMath path is present --
        choice = ac.find(qn("mc:Choice"))
        assert choice is not None
        assert choice.get("Requires") == "a14"
        a14_m = choice.find(qn("a14:m"))
        assert a14_m is not None
        oMath_elms = a14_m.findall(qn("m:oMath"))
        assert len(oMath_elms) == 1
        # -- mc:Fallback/a:r/a:t carries the extracted plain text --
        fallback = ac.find(qn("mc:Fallback"))
        assert fallback is not None
        fallback_r = fallback.find(qn("a:r"))
        assert fallback_r is not None
        assert fallback_r.findtext(qn("a:t")) == "1+2"

    def it_can_add_a_math_equation_accepting_oMathPara(self):
        from pptx.oxml.ns import qn

        p = cast("CT_TextParagraph", element("a:p"))
        paragraph = _Paragraph(p, None)
        omml = (
            '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:oMath><m:r><m:t>y=</m:t></m:r><m:r><m:t>2</m:t></m:r></m:oMath>"
            "</m:oMathPara>"
        )

        paragraph.add_math_equation(omml)

        ac = paragraph._p.find(qn("mc:AlternateContent"))
        assert ac is not None
        # -- the full m:oMathPara subtree is preserved --
        para = ac.find("./" + qn("mc:Choice") + "/" + qn("a14:m") + "/" + qn("m:oMathPara"))
        assert para is not None
        assert para.find(qn("m:oMath")) is not None
        # -- fallback text concatenates all m:t descendants --
        fallback = ac.find(qn("mc:Fallback"))
        assert fallback is not None
        fallback_r = fallback.find(qn("a:r"))
        assert fallback_r is not None
        assert fallback_r.findtext(qn("a:t")) == "y=2"

    def it_inserts_math_equation_before_endParaRPr(self):
        from pptx.oxml.ns import qn

        p = cast("CT_TextParagraph", element('a:p/(a:r/a:t"foo",a:endParaRPr)'))
        paragraph = _Paragraph(p, None)
        omml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>z</m:t></m:r></m:oMath>"
        )

        paragraph.add_math_equation(omml)

        # -- child order: a:r, mc:AlternateContent, a:endParaRPr --
        tags = [child.tag for child in paragraph._p.iterchildren()]
        assert tags == [qn("a:r"), qn("mc:AlternateContent"), qn("a:endParaRPr")]

    def it_raises_ValueError_on_malformed_omml(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)
        with pytest.raises(ValueError, match="not well-formed XML"):
            paragraph.add_math_equation("not xml")

    def it_raises_ValueError_when_root_is_not_oMath(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)
        with pytest.raises(ValueError, match="root element must be"):
            paragraph.add_math_equation(
                '<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
            )

    def it_raises_ValueError_when_oMathPara_has_no_oMath(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)
        with pytest.raises(ValueError, match="must contain at least one"):
            paragraph.add_math_equation(
                '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/>'
            )

    def it_round_trips_an_added_math_equation_through_save_and_reload(self):
        """Regression test for issue #528 end-to-end.

        Author an equation via ``_Paragraph.add_math_equation``, save the presentation,
        reload it, and verify the equation surfaces on the hosting shape via the #126
        read API (``has_math_equation`` / ``math_equation_xml``).
        """
        import io

        from pptx import Presentation
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        paragraph = shape.text_frame.paragraphs[0]
        paragraph.text = "Formula: "
        omml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>E=mc^2</m:t></m:r></m:oMath>"
        )
        paragraph.add_math_equation(omml)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reloaded = Presentation(buf)

        # -- the authored shape reports an equation after reload --
        equation_shapes = [s for s in reloaded.slides[0].shapes if s.has_math_equation]
        assert len(equation_shapes) == 1
        oMath_xml = equation_shapes[0].math_equation_xml
        assert oMath_xml is not None
        assert oMath_xml.startswith("<m:oMath")
        assert "<m:t>E=mc^2</m:t>" in oMath_xml

    def it_knows_its_horizontal_alignment(self, alignment_get_fixture):
        paragraph, expected_value = alignment_get_fixture
        assert paragraph.alignment == expected_value

    def it_can_change_its_horizontal_alignment(self, alignment_set_fixture):
        paragraph, new_value, expected_xml = alignment_set_fixture
        paragraph.alignment = new_value
        assert paragraph._element.xml == expected_xml

    def it_can_clear_itself_of_content(self, clear_fixture):
        paragraph, expected_xml = clear_fixture
        paragraph.clear()
        assert paragraph._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("txBody_cxml", "p_idx", "expected_cxml"),
        [
            # -- first of multiple paragraphs removed, others preserved --
            (
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"foo",a:p/a:r/a:t"bar")',
                0,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"bar")',
            ),
            # -- middle paragraph removed --
            (
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"a",a:p/a:r/a:t"b",a:p/a:r/a:t"c")',
                1,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"a",a:p/a:r/a:t"c")',
            ),
            # -- last paragraph of multiple removed --
            (
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"foo",a:p/a:r/a:t"bar")',
                1,
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"foo")',
            ),
            # -- sole paragraph removed; empty <a:p> added to preserve invariant --
            (
                'p:txBody/(a:bodyPr,a:p/a:r/a:t"only")',
                0,
                "p:txBody/(a:bodyPr,a:p)",
            ),
            # -- a:txBody form (table-cell) is also handled --
            (
                'a:txBody/(a:bodyPr,a:p/a:r/a:t"only")',
                0,
                "a:txBody/(a:bodyPr,a:p)",
            ),
        ],
    )
    def it_can_delete_itself_from_its_text_frame(
        self, txBody_cxml: str, p_idx: int, expected_cxml: str
    ):
        txBody = element(txBody_cxml)
        p = txBody.p_lst[p_idx]
        paragraph = _Paragraph(p, None)

        paragraph.delete()

        assert txBody.xml == xml(expected_cxml)

    def it_provides_access_to_the_default_paragraph_font(self, paragraph, Font_):
        font = paragraph.font
        Font_.assert_called_once_with(paragraph._defRPr)
        assert font == Font_.return_value

    def it_provides_access_to_its_bullet_format(self, paragraph: _Paragraph):
        bullet = paragraph.bullet
        assert isinstance(bullet, _BulletFormat)
        # -- bullet is a lazyproperty so the same instance is returned every time --
        assert paragraph.bullet is bullet
        # -- bullet wraps the paragraph's a:pPr element --
        assert bullet._pPr is paragraph._pPr  # pyright: ignore[reportPrivateUsage]

    def it_knows_its_indentation_level(self, level_get_fixture):
        paragraph, expected_value = level_get_fixture
        assert paragraph.level == expected_value

    def it_can_change_its_indentation_level(self, level_set_fixture):
        paragraph, new_value, expected_xml = level_set_fixture
        paragraph.level = new_value
        assert paragraph._element.xml == expected_xml

    def it_knows_its_line_spacing(self, spacing_get_fixture):
        paragraph, expected_value = spacing_get_fixture
        assert paragraph.line_spacing == expected_value

    def it_can_change_its_line_spacing(self, spacing_set_fixture):
        paragraph, new_value, expected_xml = spacing_set_fixture
        paragraph.line_spacing = new_value
        assert paragraph._element.xml == expected_xml

    def it_provides_access_to_its_runs(self, runs_fixture):
        paragraph, expected_text = runs_fixture
        runs = paragraph.runs
        assert tuple(r.text for r in runs) == expected_text
        for r in runs:
            assert isinstance(r, _Run)
            assert r._parent == paragraph

    def it_knows_its_space_after(self, after_get_fixture):
        paragraph, expected_value = after_get_fixture
        assert paragraph.space_after == expected_value

    def it_can_change_its_space_after(self, after_set_fixture):
        paragraph, new_value, expected_xml = after_set_fixture
        paragraph.space_after = new_value
        assert paragraph._element.xml == expected_xml

    def it_knows_its_space_before(self, before_get_fixture):
        paragraph, expected_value = before_get_fixture
        assert paragraph.space_before == expected_value

    def it_can_change_its_space_before(self, before_set_fixture):
        paragraph, new_value, expected_xml = before_set_fixture
        paragraph.space_before = new_value
        assert paragraph._element.xml == expected_xml

    def it_knows_what_text_it_contains(self, text_get_fixture):
        p, expected_value = text_get_fixture
        paragraph = _Paragraph(p, None)

        text = paragraph.text

        assert text == expected_value
        assert isinstance(text, str)

    @pytest.mark.parametrize(
        ("p_cxml", "value", "expected_cxml"),
        [
            ('a:p/(a:r/a:t"foo",a:r/a:t"bar")', "foobar", 'a:p/a:r/a:t"foobar"'),
            ("a:p", "", "a:p"),
            ("a:p", "foobar", 'a:p/a:r/a:t"foobar"'),
            ("a:p", "foo\nbar", 'a:p/(a:r/a:t"foo",a:br,a:r/a:t"bar")'),
            ("a:p", "\vfoo\n", 'a:p/(a:br,a:r/a:t"foo",a:br)'),
            ("a:p", "\n\nfoo", 'a:p/(a:br,a:br,a:r/a:t"foo")'),
            ("a:p", "foo\n", 'a:p/(a:r/a:t"foo",a:br)'),
            ("a:p", "foo\x07\n", 'a:p/(a:r/a:t"foo_x0007_",a:br)'),
            ("a:p", "ŮŦƑ-8\x1bliteral", 'a:p/a:r/a:t"ŮŦƑ-8_x001B_literal"'),
            (
                "a:p",
                "utf-8 unicode: Hér er texti",
                'a:p/a:r/a:t"utf-8 unicode: Hér er texti"',
            ),
        ],
    )
    def it_can_change_its_text(self, p_cxml: str, value: str, expected_cxml: str):
        p = cast("CT_TextParagraph", element(p_cxml))
        paragraph = _Paragraph(p, None)

        paragraph.text = value

        assert paragraph._element.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("p_cxml", "find", "replace", "expected_count", "expected_cxml"),
        [
            # -- match inside a single run --
            (
                'a:p/a:r/a:t"Hello {NAME}!"',
                "{NAME}",
                "Alice",
                1,
                'a:p/a:r/a:t"Hello Alice!"',
            ),
            # -- match spans two consecutive runs (the issue-#836 case) --
            (
                'a:p/(a:r/a:t"{NA",a:r/a:t"ME}")',
                "{NAME}",
                "Alice",
                1,
                'a:p/a:r/a:t"Alice"',
            ),
            # -- formatting of the run that starts the match is preserved;
            # -- the run where the match ends keeps its surviving suffix
            # -- along with its own formatting
            (
                'a:p/(a:r/(a:rPr{b=1},a:t"{N"),a:r/(a:rPr{i=1},a:t"AM"),'
                'a:r/(a:rPr{u=sng},a:t"E}tail"))',
                "{NAME}",
                "X",
                1,
                'a:p/(a:r/(a:rPr{b=1},a:t"X"),a:r/(a:rPr{u=sng},a:t"tail"))',
            ),
            # -- two matches in a single paragraph: one within a run, one
            # -- spanning two runs (the first run absorbs the replacement
            # -- of the cross-run match; the last run keeps its suffix) --
            (
                'a:p/(a:r/a:t"{X} and {",a:r/a:t"X}!")',
                "{X}",
                "Y",
                2,
                'a:p/(a:r/a:t"Y and Y",a:r/a:t"!")',
            ),
            # -- match does NOT cross an a:br boundary --
            (
                'a:p/(a:r/a:t"{NA",a:br,a:r/a:t"ME}")',
                "{NAME}",
                "X",
                0,
                'a:p/(a:r/a:t"{NA",a:br,a:r/a:t"ME}")',
            ),
            # -- match does NOT cross an a:fld boundary --
            (
                'a:p/(a:r/a:t"{NA",a:fld{id=abc,type=slidenum}/a:t"#",a:r/a:t"ME}")',
                "{NAME}",
                "X",
                0,
                'a:p/(a:r/a:t"{NA",a:fld{id=abc,type=slidenum}/a:t"#",a:r/a:t"ME}")',
            ),
            # -- replacing with empty string removes the matched text --
            (
                'a:p/a:r/a:t"hello {NAME} world"',
                "{NAME} ",
                "",
                1,
                'a:p/a:r/a:t"hello world"',
            ),
            # -- replacements do not overlap; "aa" in "aaaa" replaces twice --
            (
                'a:p/a:r/a:t"aaaa"',
                "aa",
                "b",
                2,
                'a:p/a:r/a:t"bb"',
            ),
            # -- no match is a zero-count no-op --
            (
                'a:p/a:r/a:t"Hello"',
                "XYZ",
                "W",
                0,
                'a:p/a:r/a:t"Hello"',
            ),
            # -- empty paragraph (no runs) --
            (
                "a:p",
                "X",
                "Y",
                0,
                "a:p",
            ),
        ],
    )
    def it_can_replace_text_across_multiple_runs(
        self,
        p_cxml: str,
        find: str,
        replace: str,
        expected_count: int,
        expected_cxml: str,
    ):
        paragraph = _Paragraph(cast("CT_TextParagraph", element(p_cxml)), None)
        assert paragraph.replace_text(find, replace) == expected_count
        assert paragraph._element.xml == xml(expected_cxml)

    def it_raises_ValueError_when_replace_text_find_is_empty(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element('a:p/a:r/a:t"x"')), None)
        with pytest.raises(ValueError):
            paragraph.replace_text("", "y")

    def it_can_write_rich_text_in_one_call_issue_753(self):
        """Regression test for issue #753 — mixed-formatting authoring helper."""
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        result = paragraph.write_rich(
            "plain ",
            ("bold ", {"bold": True}),
            ("italic", {"italic": True}),
        )

        # -- returns the paragraph for chaining --
        assert result is paragraph
        runs = paragraph.runs
        assert len(runs) == 3
        assert [r.text for r in runs] == ["plain ", "bold ", "italic"]
        assert runs[0].font.bold is None
        assert runs[0].font.italic is None
        assert runs[1].font.bold is True
        assert runs[1].font.italic is None
        assert runs[2].font.bold is None
        assert runs[2].font.italic is True

    def it_applies_every_supported_formatting_key_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        paragraph.write_rich(
            (
                "all of it",
                {
                    "bold": True,
                    "italic": True,
                    "underline": True,
                    "size": Pt(24),
                    "color": RGBColor(0xFF, 0x00, 0x00),
                    "font_name": "Calibri",
                },
            )
        )

        (run,) = paragraph.runs
        assert run.text == "all of it"
        assert run.font.bold is True
        assert run.font.italic is True
        assert run.font.underline is True
        assert run.font.size == Pt(24)
        assert run.font.name == "Calibri"
        assert run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)

    def it_accepts_a_mapping_part_with_a_text_key_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        paragraph.write_rich({"text": "hi", "bold": True, "italic": False})

        (run,) = paragraph.runs
        assert run.text == "hi"
        assert run.font.bold is True
        assert run.font.italic is False

    def it_accepts_an_MSO_UNDERLINE_value_for_underline_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        paragraph.write_rich(("wavy", {"underline": MSO_UNDERLINE.WAVY_LINE}))

        (run,) = paragraph.runs
        assert run.font.underline == MSO_UNDERLINE.WAVY_LINE

    def it_appends_runs_after_existing_content_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element('a:p/a:r/a:t"existing"')), None)

        paragraph.write_rich(" added")

        assert [r.text for r in paragraph.runs] == ["existing", " added"]

    def it_accepts_zero_parts_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        paragraph.write_rich()

        assert paragraph.runs == ()

    def it_raises_ValueError_on_unknown_formatting_key_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        with pytest.raises(ValueError, match="unknown rich-text formatting key"):
            paragraph.write_rich(("x", {"weight": 700}))

    def it_raises_TypeError_on_non_str_non_tuple_non_mapping_part_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        with pytest.raises(TypeError, match="each part must be"):
            paragraph.write_rich(42)

    def it_raises_ValueError_on_wrong_length_tuple_part_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        with pytest.raises(ValueError, match="tuple part must be"):
            paragraph.write_rich(("only-one-element",))

    def it_raises_ValueError_on_mapping_part_missing_text_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        with pytest.raises(ValueError, match="must contain a 'text' key"):
            paragraph.write_rich({"bold": True})

    def it_raises_TypeError_on_non_str_text_in_tuple_in_write_rich(self):
        paragraph = _Paragraph(cast("CT_TextParagraph", element("a:p")), None)

        with pytest.raises(TypeError, match="must be str"):
            paragraph.write_rich((123, {}))

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[
            ("a:p", None),
            ("a:p/a:pPr", None),
            ("a:p/a:pPr/a:spcAft/a:spcPct{val=150000}", None),
            ("a:p/a:pPr/a:spcAft/a:spcPts{val=600}", 76200),
        ]
    )
    def after_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        return paragraph, expected_value

    @pytest.fixture(
        params=[
            ("a:p", Pt(8.333), "a:p/a:pPr/a:spcAft/a:spcPts{val=833}"),
            (
                "a:p/a:pPr/a:spcAft/a:spcPts{val=600}",
                Pt(42),
                "a:p/a:pPr/a:spcAft/a:spcPts{val=4200}",
            ),
            (
                "a:p/a:pPr/a:spcAft/a:spcPct{val=150000}",
                Pt(24),
                "a:p/a:pPr/a:spcAft/a:spcPts{val=2400}",
            ),
            ("a:p/a:pPr/a:spcAft/a:spcPts{val=600}", None, "a:p/a:pPr"),
        ]
    )
    def after_set_fixture(self, request):
        p_cxml, new_value, expected_p_cxml = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        expected_xml = xml(expected_p_cxml)
        return paragraph, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:p", None),
            ("a:p/a:pPr{algn=ctr}", PP_ALIGN.CENTER),
            ("a:p/a:pPr{algn=r}", PP_ALIGN.RIGHT),
        ]
    )
    def alignment_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        return paragraph, expected_value

    @pytest.fixture(
        params=[
            ("a:p", PP_ALIGN.LEFT, "a:p/a:pPr{algn=l}"),
            ("a:p/a:pPr{algn=l}", PP_ALIGN.JUSTIFY, "a:p/a:pPr{algn=just}"),
            ("a:p/a:pPr{algn=just}", None, "a:p/a:pPr"),
        ]
    )
    def alignment_set_fixture(self, request):
        p_cxml, new_value, expected_p_cxml = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        expected_xml = xml(expected_p_cxml)
        return paragraph, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:p", None),
            ("a:p/a:pPr", None),
            ("a:p/a:pPr/a:spcBef/a:spcPct{val=150000}", None),
            ("a:p/a:pPr/a:spcBef/a:spcPts{val=600}", 76200),
        ]
    )
    def before_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        return paragraph, expected_value

    @pytest.fixture(
        params=[
            ("a:p", Pt(8.333), "a:p/a:pPr/a:spcBef/a:spcPts{val=833}"),
            (
                "a:p/a:pPr/a:spcBef/a:spcPts{val=600}",
                Pt(42),
                "a:p/a:pPr/a:spcBef/a:spcPts{val=4200}",
            ),
            (
                "a:p/a:pPr/a:spcBef/a:spcPct{val=150000}",
                Pt(24),
                "a:p/a:pPr/a:spcBef/a:spcPts{val=2400}",
            ),
            ("a:p/a:pPr/a:spcBef/a:spcPts{val=600}", None, "a:p/a:pPr"),
        ]
    )
    def before_set_fixture(self, request):
        p_cxml, new_value, expected_p_cxml = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        expected_xml = xml(expected_p_cxml)
        return paragraph, new_value, expected_xml

    @pytest.fixture(
        params=[
            ('a:p/a:r/a:t"foo"', "a:p"),
            ('a:p/(a:br,a:r/a:t"foo")', "a:p"),
            ('a:p/(a:fld,a:br,a:r/a:t"foo")', "a:p"),
        ]
    )
    def clear_fixture(self, request):
        p_cxml, expected_p_cxml = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        expected_xml = xml(expected_p_cxml)
        return paragraph, expected_xml

    @pytest.fixture(params=[("a:p", 0), ("a:p/a:pPr{lvl=2}", 2)])
    def level_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        return paragraph, expected_value

    @pytest.fixture(
        params=[
            ("a:p", 1, "a:p/a:pPr{lvl=1}"),
            ("a:p/a:pPr{lvl=1}", 2, "a:p/a:pPr{lvl=2}"),
            ("a:p/a:pPr{lvl=2}", 0, "a:p/a:pPr"),
        ]
    )
    def level_set_fixture(self, request):
        p_cxml, new_value, expected_p_cxml = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        expected_xml = xml(expected_p_cxml)
        return paragraph, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("a:p", "a:p/a:br"),
            ("a:p/a:r", "a:p/(a:r,a:br)"),
            ("a:p/a:br", "a:p/(a:br,a:br)"),
        ]
    )
    def line_break_fixture(self, request):
        cxml, expected_cxml = request.param
        paragraph = _Paragraph(element(cxml), None)
        expected_xml = xml(expected_cxml)
        return paragraph, expected_xml

    @pytest.fixture
    def runs_fixture(self):
        p_cxml = 'a:p/(a:r/a:t"Foo",a:r/a:t"Bar",a:r/a:t"Baz")'
        paragraph = _Paragraph(element(p_cxml), None)
        expected_text = ("Foo", "Bar", "Baz")
        return paragraph, expected_text

    @pytest.fixture(
        params=[
            ("a:p", None),
            ("a:p/a:pPr", None),
            ("a:p/a:pPr/a:lnSpc/a:spcPts{val=1800}", 228600),
            ("a:p/a:pPr/a:lnSpc/a:spcPct{val=142000}", 1.42),
            ("a:p/a:pPr/a:lnSpc/a:spcPct{val=124.64%}", 1.2464),
        ]
    )
    def spacing_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        return paragraph, expected_value

    @pytest.fixture(
        params=[
            ("a:p", 1.42, "a:p/a:pPr/a:lnSpc/a:spcPct{val=142000}"),
            ("a:p", Pt(42), "a:p/a:pPr/a:lnSpc/a:spcPts{val=4200}"),
            (
                "a:p/a:pPr/a:lnSpc/a:spcPct{val=110000}",
                0.875,
                "a:p/a:pPr/a:lnSpc/a:spcPct{val=87500}",
            ),
            (
                "a:p/a:pPr/a:lnSpc/a:spcPts{val=600}",
                Pt(42),
                "a:p/a:pPr/a:lnSpc/a:spcPts{val=4200}",
            ),
            (
                "a:p/a:pPr/a:lnSpc/a:spcPts{val=1900}",
                0.925,
                "a:p/a:pPr/a:lnSpc/a:spcPct{val=92500}",
            ),
            (
                "a:p/a:pPr/a:lnSpc/a:spcPct{val=150000}",
                Pt(24),
                "a:p/a:pPr/a:lnSpc/a:spcPts{val=2400}",
            ),
            ("a:p/a:pPr/a:lnSpc/a:spcPts{val=600}", None, "a:p/a:pPr"),
            ("a:p/a:pPr/a:lnSpc/a:spcPct{val=150000}", None, "a:p/a:pPr"),
        ]
    )
    def spacing_set_fixture(self, request):
        p_cxml, new_value, expected_p_cxml = request.param
        paragraph = _Paragraph(element(p_cxml), None)
        expected_xml = xml(expected_p_cxml)
        return paragraph, new_value, expected_xml

    @pytest.fixture(
        params=[
            # ---single-run---
            ('a:p/a:r/a:t"foobar"', "foobar"),
            # ---multiple-runs---
            ('a:p/(a:r/a:t"foo",a:r/a:t"bar")', "foobar"),
            # ---line-break between runs---
            ('a:p/(a:r/a:t"foo",a:br,a:r/a:t"bar")', "foo\vbar"),
            # ---field between runs---
            ('a:p/(a:r/a:t"foo ",a:fld/a:t"42",a:r/a:t" bar")', "foo 42 bar"),
            # ---line-break and field---
            ('a:p/(a:r/a:t" foo",a:br,a:fld/a:t"42")', " foo\v42"),
            # ---other common p child elements included---
            ('a:p/(a:pPr,a:r/a:t"foobar",a:endParaRPr)', "foobar"),
            # ---field by itself---
            ('a:p/a:fld/a:t"42"', "42"),
            # ---line-break by itself---
            ("a:p/a:br", "\v"),
        ]
    )
    def text_get_fixture(self, request):
        p_cxml, expected_value = request.param
        p = element(p_cxml)
        return p, expected_value

    # fixture components -----------------------------------

    @pytest.fixture
    def Font_(self, request):
        return class_mock(request, "pptx.text.text.Font")

    @pytest.fixture
    def p_bldr(self):
        return a_p().with_nsdecls()

    @pytest.fixture
    def p_with_r_xml(self):
        run_bldr = an_r().with_child(a_t())
        return a_p().with_nsdecls().with_child(run_bldr).xml()

    @pytest.fixture
    def paragraph(self, p_bldr):
        return _Paragraph(p_bldr.element, None)


class Describe_BulletFormat(object):
    """Unit-test suite for `pptx.text.text._BulletFormat` object."""

    # -- type getter --------------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_type"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buNone", "none"),
            ("a:pPr/a:buChar{char=-}", "char"),
            ("a:pPr/a:buAutoNum{type=arabicPeriod}", "autonum"),
        ],
    )
    def it_knows_its_bullet_type(self, pPr_cxml: str, expected_type: str | None):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)
        assert bullet.type == expected_type

    # -- char getter --------------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_char"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buNone", None),
            ("a:pPr/a:buChar{char=-}", "-"),
            ("a:pPr/a:buChar{char=x}", "x"),
        ],
    )
    def it_knows_its_bullet_char(self, pPr_cxml: str, expected_char: str | None):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)
        assert bullet.char == expected_char

    # -- number_scheme getter -----------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_scheme"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buNone", None),
            ("a:pPr/a:buAutoNum{type=arabicPeriod}", PP_AUTO_NUMBER.ARABIC_PERIOD),
            ("a:pPr/a:buAutoNum{type=romanUcPeriod}", PP_AUTO_NUMBER.ROMAN_UC_PERIOD),
        ],
    )
    def it_knows_its_number_scheme(
        self, pPr_cxml: str, expected_scheme: PP_AUTO_NUMBER_SCHEME | None
    ):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)
        assert bullet.number_scheme == expected_scheme

    # -- start_at getter ----------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_start_at"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buAutoNum{type=arabicPeriod}", 1),
            ("a:pPr/a:buAutoNum{type=arabicPeriod,startAt=5}", 5),
        ],
    )
    def it_knows_its_start_at(self, pPr_cxml: str, expected_start_at: int | None):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)
        assert bullet.start_at == expected_start_at

    # -- character() setter -------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "char", "expected_cxml"),
        [
            ("a:pPr", "-", "a:pPr/a:buChar{char=-}"),
            ("a:pPr/a:buNone", "-", "a:pPr/a:buChar{char=-}"),
            (
                "a:pPr/a:buAutoNum{type=arabicPeriod}",
                "-",
                "a:pPr/a:buChar{char=-}",
            ),
            ("a:pPr/a:buChar{char=-}", "x", "a:pPr/a:buChar{char=x}"),
        ],
    )
    def it_can_set_a_character_bullet(self, pPr_cxml: str, char: str, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        return_value = bullet.character(char)

        assert pPr.xml == xml(expected_cxml)
        assert return_value is bullet

    @pytest.mark.parametrize("bad_char", ["", "xy", 42, None])
    def it_raises_on_invalid_bullet_character(self, bad_char: object):
        pPr = element("a:pPr")
        bullet = _BulletFormat(pPr)

        with pytest.raises(ValueError, match="single-character"):
            bullet.character(bad_char)  # pyright: ignore[reportArgumentType]

    # -- auto_number() setter -----------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "scheme", "start_at", "expected_cxml"),
        [
            (
                "a:pPr",
                PP_AUTO_NUMBER.ARABIC_PERIOD,
                None,
                "a:pPr/a:buAutoNum{type=arabicPeriod}",
            ),
            (
                "a:pPr/a:buNone",
                PP_AUTO_NUMBER.ROMAN_LC_PERIOD,
                None,
                "a:pPr/a:buAutoNum{type=romanLcPeriod}",
            ),
            (
                "a:pPr/a:buChar{char=-}",
                PP_AUTO_NUMBER.ARABIC_PERIOD,
                3,
                "a:pPr/a:buAutoNum{type=arabicPeriod,startAt=3}",
            ),
        ],
    )
    def it_can_set_an_auto_number_bullet(
        self,
        pPr_cxml: str,
        scheme: PP_AUTO_NUMBER_SCHEME,
        start_at: int | None,
        expected_cxml: str,
    ):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        return_value = bullet.auto_number(scheme, start_at)

        assert pPr.xml == xml(expected_cxml)
        assert return_value is bullet

    def it_raises_on_invalid_auto_number_scheme(self):
        pPr = element("a:pPr")
        bullet = _BulletFormat(pPr)
        # -- an int outside the member set raises ValueError --
        with pytest.raises(ValueError):
            bullet.auto_number(9999)  # pyright: ignore[reportArgumentType]

    # -- none() setter ------------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_cxml"),
        [
            ("a:pPr", "a:pPr/a:buNone"),
            ("a:pPr/a:buNone", "a:pPr/a:buNone"),
            ("a:pPr/a:buChar{char=-}", "a:pPr/a:buNone"),
            (
                "a:pPr/a:buAutoNum{type=arabicPeriod,startAt=5}",
                "a:pPr/a:buNone",
            ),
        ],
    )
    def it_can_explicitly_suppress_a_bullet(self, pPr_cxml: str, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        return_value = bullet.none()

        assert pPr.xml == xml(expected_cxml)
        assert return_value is bullet

    # -- clear() setter -----------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_cxml"),
        [
            ("a:pPr", "a:pPr"),
            ("a:pPr/a:buNone", "a:pPr"),
            ("a:pPr/a:buChar{char=-}", "a:pPr"),
            ("a:pPr/a:buAutoNum{type=arabicPeriod,startAt=5}", "a:pPr"),
        ],
    )
    def it_can_remove_any_bullet_setting(self, pPr_cxml: str, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        return_value = bullet.clear()

        assert pPr.xml == xml(expected_cxml)
        assert return_value is bullet

    # -- font getter/setter ------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_font"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buNone", None),
            ("a:pPr/a:buFont{typeface=Wingdings}", "Wingdings"),
            ("a:pPr/(a:buFont{typeface=Arial},a:buChar{char=-})", "Arial"),
        ],
    )
    def it_knows_its_bullet_font(self, pPr_cxml: str, expected_font: str | None):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)
        assert bullet.font == expected_font

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", "Wingdings", "a:pPr/a:buFont{typeface=Wingdings}"),
            (
                "a:pPr/a:buFont{typeface=Arial}",
                "Wingdings",
                "a:pPr/a:buFont{typeface=Wingdings}",
            ),
            ("a:pPr/a:buFont{typeface=Wingdings}", None, "a:pPr"),
            # -- setting font preserves bullet choice element ordering --
            (
                "a:pPr/a:buChar{char=-}",
                "Wingdings",
                "a:pPr/(a:buFont{typeface=Wingdings},a:buChar{char=-})",
            ),
        ],
    )
    def it_can_change_its_bullet_font(self, pPr_cxml: str, value: str | None, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        bullet.font = value

        assert pPr.xml == xml(expected_cxml)

    # -- size_pct getter/setter --------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buSzPct{val=75%}", 0.75),
            ("a:pPr/a:buSzPts{val=1400}", None),
        ],
    )
    def it_knows_its_bullet_size_pct(self, pPr_cxml: str, expected_value: float | None):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)
        assert bullet.size_pct == expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", 0.75, "a:pPr/a:buSzPct{val=75%}"),
            # -- assigning pct replaces any buSzPts --
            (
                "a:pPr/a:buSzPts{val=1400}",
                0.5,
                "a:pPr/a:buSzPct{val=50%}",
            ),
            ("a:pPr/a:buSzPct{val=75%}", None, "a:pPr"),
        ],
    )
    def it_can_change_its_bullet_size_pct(
        self, pPr_cxml: str, value: float | None, expected_cxml: str
    ):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        bullet.size_pct = value

        assert pPr.xml == xml(expected_cxml)

    def it_raises_on_out_of_range_bullet_size_pct(self):
        pPr = element("a:pPr")
        bullet = _BulletFormat(pPr)
        with pytest.raises(ValueError):
            bullet.size_pct = 5.0  # 500%, above max 400%
        with pytest.raises(ValueError):
            bullet.size_pct = 0.1  # 10%, below min 25%

    # -- size_points getter/setter -----------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buSzPts{val=1400}", 177800),
            ("a:pPr/a:buSzPct{val=75%}", None),
        ],
    )
    def it_knows_its_bullet_size_points(self, pPr_cxml: str, expected_value: int | None):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)
        assert bullet.size_points == expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", Pt(14), "a:pPr/a:buSzPts{val=1400}"),
            # -- assigning points replaces any buSzPct --
            (
                "a:pPr/a:buSzPct{val=75%}",
                Pt(18),
                "a:pPr/a:buSzPts{val=1800}",
            ),
            ("a:pPr/a:buSzPts{val=1400}", None, "a:pPr"),
        ],
    )
    def it_can_change_its_bullet_size_points(self, pPr_cxml: str, value, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        bullet.size_points = value

        assert pPr.xml == xml(expected_cxml)

    # -- clear_size() ------------------------------------------------

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_cxml"),
        [
            ("a:pPr", "a:pPr"),
            ("a:pPr/a:buSzPct{val=75%}", "a:pPr"),
            ("a:pPr/a:buSzPts{val=1400}", "a:pPr"),
        ],
    )
    def it_can_clear_its_bullet_size(self, pPr_cxml: str, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = _BulletFormat(pPr)

        return_value = bullet.clear_size()

        assert pPr.xml == xml(expected_cxml)
        assert return_value is bullet

    # -- color ------------------------------------------------------

    def it_provides_a_ColorFormat_for_the_bullet_color(self):
        pPr = element("a:pPr")
        bullet = _BulletFormat(pPr)

        color = bullet.color

        assert isinstance(color, ColorFormat)
        # -- the accessor is a lazyproperty; same instance is returned --
        assert bullet.color is color

    def it_creates_buClr_on_access_and_allows_setting_rgb(self):
        pPr = element("a:pPr")
        bullet = _BulletFormat(pPr)

        bullet.color.rgb = RGBColor(0xFF, 0x00, 0x00)

        assert pPr.xml == xml("a:pPr/a:buClr/a:srgbClr{val=FF0000}")

    def it_can_clear_its_bullet_color(self):
        pPr = element("a:pPr/a:buClr/a:srgbClr{val=FF0000}")
        bullet = _BulletFormat(pPr)

        return_value = bullet.clear_color()

        assert pPr.xml == xml("a:pPr")
        assert return_value is bullet


class Describe_Run(object):
    """Unit-test suite for `pptx.text.text._Run` object."""

    def it_provides_access_to_its_font(self, font_fixture):
        run, rPr, Font_, font_ = font_fixture
        font = run.font
        Font_.assert_called_once_with(rPr, parent=run)
        assert font == font_

    def it_provides_access_to_a_hyperlink_proxy(self, hyperlink_fixture):
        run, rPr, _Hyperlink_, hlink_ = hyperlink_fixture
        hlink = run.hyperlink
        _Hyperlink_.assert_called_once_with(rPr, run)
        assert hlink is hlink_

    def it_can_get_the_text_of_the_run(self, text_get_fixture):
        run, expected_value = text_get_fixture
        text = run.text
        assert text == expected_value
        assert isinstance(text, str)

    @pytest.mark.parametrize(
        "r_cxml, new_value, expected_r_cxml",
        (
            ("a:r/a:t", "barfoo", 'a:r/a:t"barfoo"'),
            ("a:r/a:t", "bar\x1bfoo", 'a:r/a:t"bar_x001B_foo"'),
            ("a:r/a:t", "bar\tfoo", 'a:r/a:t"bar\tfoo"'),
        ),
    )
    def it_can_change_its_text(self, r_cxml, new_value, expected_r_cxml):
        run = _Run(element(r_cxml), None)
        run.text = new_value
        assert run._r.xml == xml(expected_r_cxml)

    @pytest.mark.parametrize(
        ("p_cxml", "r_idx", "expected_cxml"),
        [
            # -- first of two runs removed; sibling run preserved --
            (
                'a:p/(a:r/a:t"foo",a:r/a:t"bar")',
                0,
                'a:p/a:r/a:t"bar"',
            ),
            # -- second of two runs removed --
            (
                'a:p/(a:r/a:t"foo",a:r/a:t"bar")',
                1,
                'a:p/a:r/a:t"foo"',
            ),
            # -- sole run removed; paragraph remains (possibly empty of runs) --
            (
                'a:p/a:r/a:t"only"',
                0,
                "a:p",
            ),
            # -- pPr and line-break siblings are preserved --
            (
                'a:p/(a:pPr,a:r/a:t"foo",a:br,a:r/a:t"bar")',
                0,
                'a:p/(a:pPr,a:br,a:r/a:t"bar")',
            ),
        ],
    )
    def it_can_delete_itself_from_its_paragraph(self, p_cxml: str, r_idx: int, expected_cxml: str):
        p = element(p_cxml)
        r = p.r_lst[r_idx]
        run = _Run(r, None)

        run.delete()

        assert p.xml == xml(expected_cxml)

    # fixtures ---------------------------------------------

    @pytest.fixture
    def font_fixture(self, Font_, font_):
        r = element("a:r/a:rPr")
        rPr = r.rPr
        run = _Run(r, None)
        return run, rPr, Font_, font_

    @pytest.fixture
    def hyperlink_fixture(self, _Hyperlink_, hlink_):
        r = element("a:r/a:rPr")
        rPr = r.rPr
        run = _Run(r, None)
        return run, rPr, _Hyperlink_, hlink_

    @pytest.fixture
    def text_get_fixture(self):
        r = element('a:r/a:t"foobar"')
        run = _Run(r, None)
        return run, "foobar"

    # fixture components -----------------------------------

    @pytest.fixture
    def Font_(self, request, font_):
        return class_mock(request, "pptx.text.text.Font", return_value=font_)

    @pytest.fixture
    def font_(self, request):
        return instance_mock(request, Font)

    @pytest.fixture
    def _Hyperlink_(self, request, hlink_):
        return class_mock(request, "pptx.text.text._Hyperlink", return_value=hlink_)

    @pytest.fixture
    def hlink_(self, request):
        return instance_mock(request, _Hyperlink)


class Describe_Field(object):
    """Unit-test suite for `pptx.text.text._Field` object."""

    def it_knows_its_field_type(self):
        fld = element('a:fld{id=1,type=slidenum}/a:t"#"')
        field = _Field(fld, None)
        assert field.field_type == "slidenum"

    def it_can_change_its_field_type(self):
        fld = element('a:fld{id=1,type=slidenum}/a:t"#"')
        field = _Field(fld, None)

        field.field_type = "datetime"

        assert field.field_type == "datetime"
        assert fld.type == "datetime"

    def it_knows_the_display_text_of_the_field(self):
        fld = element('a:fld{id=1,type=slidenum}/a:t"42"')
        field = _Field(fld, None)
        assert field.text == "42"

    def it_can_set_the_display_text_of_the_field(self):
        fld = element("a:fld{id=1,type=slidenum}")
        field = _Field(fld, None)

        field.text = "‹#›"

        assert field.text == "‹#›"

    def it_provides_access_to_its_font(self):
        fld = element("a:fld{id=1,type=slidenum}")
        field = _Field(fld, None)
        font = field.font
        assert isinstance(font, Font)
        assert fld.rPr is not None
