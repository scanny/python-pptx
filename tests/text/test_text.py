# encoding: utf-8

"""Test suite for pptx.text.text module."""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.compat import is_unicode
from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, MSO_UNDERLINE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.shapes.autoshape import Shape
from pptx.text.text import Font, _Hyperlink, _Paragraph, _Run, TextFrame
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


class DescribeTextFrame(object):
    """Unit-test suite for `pptx.text.text.TextFrame` object."""

    def it_can_add_a_paragraph_to_itself(self, add_paragraph_fixture):
        text_frame, expected_xml = add_paragraph_fixture
        text_frame.add_paragraph()
        assert text_frame._txBody.xml == expected_xml

    def it_knows_its_autosize_setting(self, autosize_get_fixture):
        text_frame, expected_value = autosize_get_fixture
        assert text_frame.auto_size == expected_value

    def it_can_change_its_autosize_setting(self, autosize_set_fixture):
        text_frame, value, expected_xml = autosize_set_fixture
        text_frame.auto_size = value
        assert text_frame._txBody.xml == expected_xml

    def it_knows_its_margin_settings(self, margin_get_fixture):
        text_frame, prop_name, unit, expected_value = margin_get_fixture
        margin_value = getattr(text_frame, prop_name)
        assert getattr(margin_value, unit) == expected_value

    def it_can_change_its_margin_settings(self, margin_set_fixture):
        text_frame, prop_name, new_value, expected_xml = margin_set_fixture
        setattr(text_frame, prop_name, new_value)
        assert text_frame._txBody.xml == expected_xml

    def it_knows_its_vertical_alignment(self, anchor_get_fixture):
        text_frame, expected_value = anchor_get_fixture
        assert text_frame.vertical_anchor == expected_value

    def it_can_change_its_vertical_alignment(self, anchor_set_fixture):
        text_frame, new_value, expected_xml = anchor_set_fixture
        text_frame.vertical_anchor = new_value
        assert text_frame._element.xml == expected_xml

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

    def it_knows_the_part_it_belongs_to(self, text_frame_with_parent_):
        text_frame, parent_ = text_frame_with_parent_
        part = text_frame.part
        assert part is parent_.part

    def it_knows_what_text_it_contains(
        self, request, text_get_fixture, paragraphs_prop_
    ):
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

    def it_can_resize_its_text_to_best_fit(
        self, text_prop_, _best_fit_font_size_, _apply_fit_
    ):
        family, max_size, bold, italic, font_file, font_size = (
            "Family",
            42,
            "bold",
            "italic",
            "font_file",
            21,
        )
        text_prop_.return_value = "some text"
        _best_fit_font_size_.return_value = font_size
        text_frame = TextFrame(None, None)

        text_frame.fit_text(family, max_size, bold, italic, font_file)

        text_frame._best_fit_font_size.assert_called_once_with(
            family, max_size, bold, italic, font_file
        )
        text_frame._apply_fit.assert_called_once_with(family, font_size, bold, italic)

    def it_calculates_its_best_fit_font_size_to_help_fit_text(self, size_font_fixture):
        text_frame, family, max_size, bold, italic = size_font_fixture[:5]
        FontFiles_, TextFitter_, text, extents = size_font_fixture[5:9]
        font_file_, font_size_ = size_font_fixture[9:]

        font_size = text_frame._best_fit_font_size(family, max_size, bold, italic, None)

        FontFiles_.find.assert_called_once_with(family, bold, italic)
        TextFitter_.best_fit_font_size.assert_called_once_with(
            text, extents, max_size, font_file_
        )
        assert font_size is font_size_

    def it_calculates_its_effective_size_to_help_fit_text(self):
        sp_cxml = (
            "p:sp/(p:spPr/a:xfrm/(a:off{x=914400,y=914400},a:ext{cx=914400,c"
            "y=914400}),p:txBody/(a:bodyPr,a:p))"
        )
        text_frame = Shape(element(sp_cxml), None).text_frame
        assert text_frame._extents == (731520, 822960)

    def it_applies_fit_to_help_fit_text(self, apply_fit_fixture):
        text_frame, family, font_size, bold, italic = apply_fit_fixture
        text_frame._apply_fit(family, font_size, bold, italic)
        assert text_frame.auto_size is MSO_AUTO_SIZE.NONE
        assert text_frame.word_wrap is True
        text_frame._set_font.assert_called_once_with(family, font_size, bold, italic)

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
            ("p:txBody/a:bodyPr{anchor=t}", MSO_ANCHOR.TOP),
            ("p:txBody/a:bodyPr{anchor=b}", MSO_ANCHOR.BOTTOM),
        ]
    )
    def anchor_get_fixture(self, request):
        txBody_cxml, expected_value = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        return text_frame, expected_value

    @pytest.fixture(
        params=[
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
        ]
    )
    def anchor_set_fixture(self, request):
        txBody_cxml, new_value, expected_cxml = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        expected_xml = xml(expected_cxml)
        return text_frame, new_value, expected_xml

    @pytest.fixture
    def apply_fit_fixture(self, _set_font_):
        txBody = element("p:txBody/a:bodyPr")
        text_frame = TextFrame(txBody, None)
        family, font_size, bold, italic = "Family", 42, True, False
        return text_frame, family, font_size, bold, italic

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
        ]
    )
    def autosize_set_fixture(self, request):
        txBody_cxml, value, expected_cxml = request.param
        text_frame = TextFrame(element(txBody_cxml), None)
        expected_xml = xml(expected_cxml)
        return text_frame, value, expected_xml

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

    @pytest.fixture(
        params=[(["foobar"], "foobar"), (["foo", "bar", "baz"], "foo\nbar\nbaz")]
    )
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
    def _apply_fit_(self, request):
        return method_mock(request, TextFrame, "_apply_fit")

    @pytest.fixture
    def _best_fit_font_size_(self, request):
        return method_mock(request, TextFrame, "_best_fit_font_size")

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
    def _set_font_(self, request):
        return method_mock(request, TextFrame, "_set_font")

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


class DescribeFont(object):
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

    def it_provides_access_to_its_color(self, font):
        assert isinstance(font.color, ColorFormat)

    def it_provides_access_to_its_fill(self, font):
        assert isinstance(font.fill, FillFormat)

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[("a:rPr", None), ("a:rPr{b=0}", False), ("a:rPr{b=1}", True)]
    )
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

    @pytest.fixture(
        params=[("a:rPr", None), ("a:rPr{i=0}", False), ("a:rPr{i=1}", True)]
    )
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

    @pytest.fixture(
        params=[("a:rPr", None), ("a:rPr/a:latin{typeface=Foobar}", "Foobar")]
    )
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

    @pytest.fixture(params=[("a:rPr", None), ("a:rPr{sz=2400}", 304800)])
    def size_get_fixture(self, request):
        rPr_cxml, expected_value = request.param
        font = Font(element(rPr_cxml))
        return font, expected_value

    @pytest.fixture(
        params=[("a:rPr", Pt(24), "a:rPr{sz=2400}"), ("a:rPr{sz=2400}", None, "a:rPr")]
    )
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

    # fixture components ---------------------------------------------

    @pytest.fixture
    def font(self):
        return Font(element("a:rPr"))


class Describe_Hyperlink(object):
    def it_knows_the_target_url_of_the_hyperlink(self, hlink_with_url_):
        hlink, rId, url = hlink_with_url_
        assert hlink.address == url
        hlink.part.target_ref.assert_called_once_with(rId)

    def it_has_None_for_address_when_no_hyperlink_is_present(self, hlink):
        assert hlink.address is None

    def it_can_set_the_target_url(self, hlink, rPr_with_hlinkClick_xml, url):
        hlink.address = url
        # verify -----------------------
        hlink.part.relate_to.assert_called_once_with(
            url, RT.HYPERLINK, is_external=True
        )
        assert hlink._rPr.xml == rPr_with_hlinkClick_xml
        assert hlink.address == url

    def it_can_remove_the_hyperlink(self, remove_hlink_fixture_):
        hlink, rPr_xml, rId = remove_hlink_fixture_
        hlink.address = None
        assert hlink._rPr.xml == rPr_xml
        hlink.part.drop_rel.assert_called_once_with(rId)

    def it_should_remove_the_hyperlink_when_url_set_to_empty_string(
        self, remove_hlink_fixture_
    ):
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
        hlink.part.relate_to.assert_called_once_with(
            new_url, RT.HYPERLINK, is_external=True
        )

    # fixtures ---------------------------------------------

    @pytest.fixture
    def change_hlink_fixture_(
        self, request, hlink_with_hlinkClick, rId, rId_2, part_, url_2
    ):
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
        part_ = instance_mock(request, Part)
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

    def it_provides_access_to_the_default_paragraph_font(self, paragraph, Font_):
        font = paragraph.font
        Font_.assert_called_once_with(paragraph._defRPr)
        assert font == Font_.return_value

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
        assert is_unicode(text)

    def it_can_change_its_text(self, text_set_fixture):
        p, value, expected_xml = text_set_fixture
        paragraph = _Paragraph(p, None)

        paragraph.text = value

        assert paragraph._element.xml == expected_xml

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

    @pytest.fixture(
        params=[
            ('a:p/(a:r/a:t"foo",a:r/a:t"bar")', "foobar", 'a:p/a:r/a:t"foobar"'),
            ("a:p", "", "a:p"),
            ("a:p", "foobar", 'a:p/a:r/a:t"foobar"'),
            ("a:p", "foo\nbar", 'a:p/(a:r/a:t"foo",a:br,a:r/a:t"bar")'),
            ("a:p", "\vfoo\n", 'a:p/(a:br,a:r/a:t"foo",a:br)'),
            ("a:p", "\n\nfoo", 'a:p/(a:br,a:br,a:r/a:t"foo")'),
            ("a:p", "foo\n", 'a:p/(a:r/a:t"foo",a:br)'),
            ("a:p", b"foo\x07\n", 'a:p/(a:r/a:t"foo_x0007_",a:br)'),
            ("a:p", b"7-bit str", 'a:p/a:r/a:t"7-bit str"'),
            ("a:p", b"8-\xc9\x93\xc3\xaf\xc8\xb6 str", 'a:p/a:r/a:t"8-ɓïȶ str"'),
            ("a:p", "ŮŦƑ-8\x1bliteral", 'a:p/a:r/a:t"ŮŦƑ-8_x001B_literal"'),
            (
                "a:p",
                "utf-8 unicode: Hér er texti",
                'a:p/a:r/a:t"utf-8 unicode: Hér er texti"',
            ),
        ]
    )
    def text_set_fixture(self, request):
        p_cxml, value, expected_cxml = request.param
        p = element(p_cxml)
        expected_xml = xml(expected_cxml)
        return p, value, expected_xml

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


class Describe_Run(object):
    """Unit-test suite for `pptx.text.text._Run` object."""

    def it_provides_access_to_its_font(self, font_fixture):
        run, rPr, Font_, font_ = font_fixture
        font = run.font
        Font_.assert_called_once_with(rPr)
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
        assert is_unicode(text)

    def it_can_change_its_text(self, text_set_fixture):
        r, new_value, expected_xml = text_set_fixture
        run = _Run(r, None)

        run.text = new_value

        print("run._r.xml == %s" % repr(run._r.xml))
        print("expected_xml == %s" % repr(expected_xml))
        assert run._r.xml == expected_xml

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

    @pytest.fixture(
        params=[
            ("a:r/a:t", "barfoo", 'a:r/a:t"barfoo"'),
            ("a:r/a:t", "bar\x1bfoo", 'a:r/a:t"bar_x001B_foo"'),
            ("a:r/a:t", "bar\tfoo", 'a:r/a:t"bar\tfoo"'),
        ]
    )
    def text_set_fixture(self, request):
        r_cxml, new_value, expected_r_cxml = request.param
        r = element(r_cxml)
        expected_xml = xml(expected_r_cxml)
        return r, new_value, expected_xml

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
