"""Unit-test suite for `pptx.presentation` module."""

from __future__ import annotations

import io

import pytest

from pptx import Presentation as open_presentation
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import NotesMasterPart
from pptx.presentation import Presentation, Section, _read_blob
from pptx.slide import Slide, SlideLayouts, SlideMaster, SlideMasters, Slides

from .unitutil.cxml import element, xml
from .unitutil.mock import class_mock, instance_mock, property_mock


class DescribePresentation(object):
    def it_knows_the_height_of_its_slides(self, sld_height_get_fixture):
        prs, expected_value = sld_height_get_fixture
        assert prs.slide_height == expected_value

    def it_can_change_the_height_of_its_slides(self, sld_height_set_fixture):
        prs, slide_height, expected_xml = sld_height_set_fixture
        prs.slide_height = slide_height
        assert prs._element.xml == expected_xml

    def it_knows_the_width_of_its_slides(self, sld_width_get_fixture):
        prs, expected_value = sld_width_get_fixture
        assert prs.slide_width == expected_value

    def it_can_change_the_width_of_its_slides(self, sld_width_set_fixture):
        prs, slide_width, expected_xml = sld_width_set_fixture
        prs.slide_width = slide_width
        assert prs._element.xml == expected_xml

    def it_knows_its_part(self, part_fixture):
        prs, prs_part_ = part_fixture
        assert prs.part is prs_part_

    def it_provides_access_to_its_core_properties(self, core_props_fixture):
        prs, core_properties_ = core_props_fixture
        assert prs.core_properties is core_properties_

    def it_provides_access_to_its_notes_master(self, notes_master_fixture):
        prs, notes_master_ = notes_master_fixture
        assert prs.notes_master is notes_master_

    def it_provides_access_to_its_slides(self, slides_fixture):
        prs, rename_slide_parts_, rIds = slides_fixture[:3]
        Slides_, slides_, expected_xml = slides_fixture[3:]
        slides = prs.slides
        rename_slide_parts_.assert_called_once_with(rIds)
        Slides_.assert_called_once_with(prs._element.xpath("p:sldIdLst")[0], prs)
        assert prs._element.xml == expected_xml
        assert slides is slides_

    def it_provides_access_to_its_slide_layouts(self, layouts_fixture):
        prs, slide_layouts_ = layouts_fixture
        assert prs.slide_layouts is slide_layouts_

    def it_provides_access_to_its_slide_master(self, master_fixture):
        prs, getitem_, slide_master_ = master_fixture
        slide_master = prs.slide_master
        getitem_.assert_called_once_with(0)
        assert slide_master is slide_master_

    def it_provides_access_to_its_slide_masters(self, masters_fixture):
        prs, SlideMasters_, slide_masters_, expected_xml = masters_fixture
        slide_masters = prs.slide_masters
        SlideMasters_.assert_called_once_with(prs._element.xpath("p:sldMasterIdLst")[0], prs)
        assert slide_masters is slide_masters_
        assert prs._element.xml == expected_xml

    def it_can_save_the_presentation_to_a_file(self, save_fixture):
        prs, file_, prs_part_ = save_fixture
        prs.save(file_)
        prs_part_.save.assert_called_once_with(file_, None, password=None)

    def and_it_forwards_zip_date_time_to_the_presentation_part(self, save_fixture):
        prs, file_, prs_part_ = save_fixture
        zdt = (2024, 6, 15, 9, 30, 0)
        prs.save(file_, zdt)
        prs_part_.save.assert_called_once_with(file_, zdt, password=None)

    def it_can_save_a_password_protected_presentation_to_a_file(self, save_fixture):
        prs, file_, prs_part_ = save_fixture
        prs.save(file_, password="s3cret")
        prs_part_.save.assert_called_once_with(file_, None, password="s3cret")

    def and_it_forwards_both_zip_date_time_and_password(self, save_fixture):
        prs, file_, prs_part_ = save_fixture
        zdt = (2024, 6, 15, 9, 30, 0)
        prs.save(file_, zdt, password="s3cret")
        prs_part_.save.assert_called_once_with(file_, zdt, password="s3cret")

    def it_starts_with_no_embedded_fonts(self):
        prs = Presentation(element("p:presentation"), None)
        assert prs.embedded_fonts == ()

    def it_lists_already_embedded_fonts_in_document_order(self):
        cxml = (
            "p:presentation/p:embeddedFontLst/(p:embeddedFont/p:font{typeface=A},"
            "p:embeddedFont/p:font{typeface=B})"
        )
        prs = Presentation(element(cxml), None)
        assert prs.embedded_fonts == ("A", "B")

    def it_can_embed_a_font_from_a_file_like_object(self, request):
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.add_embedded_font.return_value = "rId7"
        prs = Presentation(element("p:presentation"), prs_part_)
        font_stream = io.BytesIO(b"fake-ttf-bytes")

        prs.embed_font(font_stream, "Pacifico")

        prs_part_.add_embedded_font.assert_called_once_with(b"fake-ttf-bytes")
        assert prs.embedded_fonts == ("Pacifico",)
        assert prs._element.embeddedFontLst is not None
        entry = prs._element.embeddedFontLst.embeddedFont_lst[0]
        assert entry.rId_for_style("regular") == "rId7"

    def it_can_embed_additional_styles_under_an_existing_typeface(self, request):
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.add_embedded_font.side_effect = iter(("rId7", "rId8"))
        prs = Presentation(element("p:presentation"), prs_part_)

        prs.embed_font(io.BytesIO(b"a"), "Pacifico", style="regular")
        prs.embed_font(io.BytesIO(b"b"), "Pacifico", style="bold")

        assert prs.embedded_fonts == ("Pacifico",)
        entry = prs._element.embeddedFontLst.embeddedFont_lst[0]
        assert entry.rId_for_style("regular") == "rId7"
        assert entry.rId_for_style("bold") == "rId8"

    def it_drops_the_prior_relationship_when_replacing_a_style_slot(self, request):
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.add_embedded_font.side_effect = iter(("rId7", "rId8"))
        prs = Presentation(element("p:presentation"), prs_part_)

        prs.embed_font(io.BytesIO(b"a"), "Pacifico")
        prs.embed_font(io.BytesIO(b"b"), "Pacifico")  # same style → replace

        prs_part_.drop_rel.assert_called_once_with("rId7")
        entry = prs._element.embeddedFontLst.embeddedFont_lst[0]
        assert entry.rId_for_style("regular") == "rId8"

    def it_raises_on_invalid_style(self, request):
        prs_part_ = instance_mock(request, PresentationPart)
        prs = Presentation(element("p:presentation"), prs_part_)

        with pytest.raises(ValueError, match="style must be one of"):
            prs.embed_font(io.BytesIO(b"x"), "Pacifico", style="extrabold")

    def it_can_embed_a_font_from_a_filesystem_path(self, request, tmp_path):
        ttf = tmp_path / "fake.ttf"
        ttf.write_bytes(b"fake-ttf-bytes")
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.add_embedded_font.return_value = "rId9"
        prs = Presentation(element("p:presentation"), prs_part_)

        prs.embed_font(str(ttf), "Roboto")

        prs_part_.add_embedded_font.assert_called_once_with(b"fake-ttf-bytes")
        assert prs.embedded_fonts == ("Roboto",)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def core_props_fixture(self, prs_part_, core_properties_):
        prs = Presentation(None, prs_part_)
        prs_part_.core_properties = core_properties_
        return prs, core_properties_

    @pytest.fixture
    def layouts_fixture(self, masters_prop_, slide_layouts_):
        prs = Presentation(None, None)
        masters_prop_.return_value.__getitem__.return_value.slide_layouts = slide_layouts_
        return prs, slide_layouts_

    @pytest.fixture
    def master_fixture(self, masters_prop_, slide_master_):
        prs = Presentation(None, None)
        getitem_ = masters_prop_.return_value.__getitem__
        getitem_.return_value = slide_master_
        return prs, getitem_, slide_master_

    @pytest.fixture(
        params=[
            ("p:presentation", "p:presentation/p:sldMasterIdLst"),
            ("p:presentation/p:sldMasterIdLst", "p:presentation/p:sldMasterIdLst"),
        ]
    )
    def masters_fixture(self, request, SlideMasters_, slide_masters_):
        prs_cxml, expected_cxml = request.param
        prs = Presentation(element(prs_cxml), None)
        expected_xml = xml(expected_cxml)
        return prs, SlideMasters_, slide_masters_, expected_xml

    @pytest.fixture
    def notes_master_fixture(self, prs_part_, notes_master_):
        prs = Presentation(None, prs_part_)
        prs_part_.notes_master = notes_master_
        return prs, notes_master_

    @pytest.fixture
    def part_fixture(self, prs_part_):
        prs = Presentation(None, prs_part_)
        return prs, prs_part_

    @pytest.fixture
    def save_fixture(self, prs_part_):
        prs = Presentation(None, prs_part_)
        file_ = "foobar.docx"
        return prs, file_, prs_part_

    @pytest.fixture(params=[("p:presentation", None), ("p:presentation/p:sldSz{cy=42}", 42)])
    def sld_height_get_fixture(self, request):
        prs_cxml, expected_value = request.param
        prs = Presentation(element(prs_cxml), None)
        return prs, expected_value

    @pytest.fixture(
        params=[
            ("p:presentation", "p:presentation/p:sldSz{cy=914400}"),
            ("p:presentation/p:sldSz{cy=424242}", "p:presentation/p:sldSz{cy=914400}"),
        ]
    )
    def sld_height_set_fixture(self, request):
        prs_cxml, expected_cxml = request.param
        prs = Presentation(element(prs_cxml), None)
        expected_xml = xml(expected_cxml)
        return prs, 914400, expected_xml

    @pytest.fixture(params=[("p:presentation", None), ("p:presentation/p:sldSz{cx=42}", 42)])
    def sld_width_get_fixture(self, request):
        prs_cxml, expected_value = request.param
        prs = Presentation(element(prs_cxml), None)
        return prs, expected_value

    @pytest.fixture(
        params=[
            ("p:presentation", "p:presentation/p:sldSz{cx=914400}"),
            ("p:presentation/p:sldSz{cx=424242}", "p:presentation/p:sldSz{cx=914400}"),
        ]
    )
    def sld_width_set_fixture(self, request):
        prs_cxml, expected_cxml = request.param
        prs = Presentation(element(prs_cxml), None)
        expected_xml = xml(expected_cxml)
        return prs, 914400, expected_xml

    @pytest.fixture(
        params=[
            ("p:presentation", [], "p:presentation/p:sldIdLst"),
            (
                "p:presentation/p:sldIdLst/p:sldId{r:id=a}",
                ["a"],
                "p:presentation/p:sldIdLst/p:sldId{r:id=a}",
            ),
            (
                "p:presentation/p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})",
                ["a", "b"],
                "p:presentation/p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})",
            ),
        ]
    )
    def slides_fixture(self, request, part_prop_, Slides_, slides_):
        prs_cxml, rIds, expected_cxml = request.param
        prs = Presentation(element(prs_cxml), None)
        rename_slide_parts_ = part_prop_.return_value.rename_slide_parts
        expected_xml = xml(expected_cxml)
        return prs, rename_slide_parts_, rIds, Slides_, slides_, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def core_properties_(self, request):
        return instance_mock(request, CorePropertiesPart)

    @pytest.fixture
    def masters_prop_(self, request):
        return property_mock(request, Presentation, "slide_masters")

    @pytest.fixture
    def notes_master_(self, request):
        return instance_mock(request, NotesMasterPart)

    @pytest.fixture
    def part_prop_(self, request):
        return property_mock(request, Presentation, "part")

    @pytest.fixture
    def prs_part_(self, request):
        return instance_mock(request, PresentationPart)

    @pytest.fixture
    def slide_layouts_(self, request):
        return instance_mock(request, SlideLayouts)

    @pytest.fixture
    def SlideMasters_(self, request, slide_masters_):
        return class_mock(request, "pptx.presentation.SlideMasters", return_value=slide_masters_)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_masters_(self, request):
        return instance_mock(request, SlideMasters)

    @pytest.fixture
    def Slides_(self, request, slides_):
        return class_mock(request, "pptx.presentation.Slides", return_value=slides_)

    @pytest.fixture
    def slides_(self, request):
        return instance_mock(request, Slides)


class DescribeSections(object):
    """Unit-test suite for `pptx.presentation.Sections`."""

    def it_returns_an_empty_collection_when_no_sections_are_present(self):
        prs = Presentation(element("p:presentation"), None)
        assert len(prs.sections) == 0
        assert list(prs.sections) == []

    def it_is_the_same_instance_on_subsequent_access(self):
        prs = Presentation(element("p:presentation"), None)
        assert prs.sections is prs.sections

    def it_raises_IndexError_on_out_of_range_access(self):
        prs = Presentation(element("p:presentation"), None)
        with pytest.raises(IndexError):
            _ = prs.sections[0]

    def it_appends_a_new_section_with_a_generated_id(self):
        prs = Presentation(element("p:presentation"), None)

        section = prs.sections.add_section("Intro")

        assert isinstance(section, Section)
        assert section.name == "Intro"
        assert len(section.id) == 38  # -- "{" + 36 + "}"
        assert section.id.startswith("{")
        assert section.id.endswith("}")
        assert len(prs.sections) == 1

    def it_accepts_a_user_supplied_guid_for_the_new_section(self):
        prs = Presentation(element("p:presentation"), None)

        section = prs.sections.add_section("Intro", id="{11111111-2222-3333-4444-555555555555}")

        assert section.id == "{11111111-2222-3333-4444-555555555555}"

    def it_rejects_a_malformed_section_guid(self):
        prs = Presentation(element("p:presentation"), None)
        with pytest.raises(ValueError, match="GUID"):
            prs.sections.add_section("X", id="not-a-guid")

    def it_rejects_a_duplicate_section_id_case_insensitively(self):
        prs = Presentation(element("p:presentation"), None)
        prs.sections.add_section("A", id="{11111111-2222-3333-4444-555555555555}")

        with pytest.raises(ValueError, match="already exists"):
            prs.sections.add_section("B", id="{11111111-2222-3333-4444-555555555555}")
        with pytest.raises(ValueError, match="already exists"):
            prs.sections.add_section("B", id="{11111111-2222-3333-4444-555555555555}".lower())

    def it_finds_a_section_by_id_case_insensitively(self):
        prs = Presentation(element("p:presentation"), None)
        added = prs.sections.add_section("Intro", id="{11111111-2222-3333-4444-555555555555}")

        assert prs.sections.get_by_id(added.id) == added
        assert prs.sections.get_by_id(added.id.lower()) == added

    def but_it_returns_None_on_unknown_id(self):
        prs = Presentation(element("p:presentation"), None)
        assert prs.sections.get_by_id("{00000000-0000-0000-0000-000000000000}") is None

    def it_finds_a_section_by_name(self):
        prs = Presentation(element("p:presentation"), None)
        prs.sections.add_section("Intro")
        prs.sections.add_section("Body")

        section = prs.sections.get_by_name("Body")

        assert section is not None
        assert section.name == "Body"

    def but_it_returns_None_on_unknown_name(self):
        prs = Presentation(element("p:presentation"), None)
        assert prs.sections.get_by_name("NoSuchSection") is None

    def it_removes_a_section_from_the_collection(self):
        prs = Presentation(element("p:presentation"), None)
        s1 = prs.sections.add_section("Intro")
        s2 = prs.sections.add_section("Body")

        prs.sections.remove(s1)

        assert len(prs.sections) == 1
        assert prs.sections[0] == s2

    def and_it_prunes_empty_extLst_scaffolding_on_last_removal(self):
        prs = Presentation(element("p:presentation"), None)
        section = prs.sections.add_section("Intro")

        prs.sections.remove(section)

        assert len(prs.sections) == 0
        # -- the enclosing extLst should have been removed entirely --
        assert prs._element.extLst is None

    def it_raises_ValueError_when_removing_a_foreign_section(self):
        prs_a = Presentation(element("p:presentation"), None)
        prs_b = Presentation(element("p:presentation"), None)
        foreign = prs_b.sections.add_section("X")

        with pytest.raises(ValueError):
            prs_a.sections.remove(foreign)

    def it_reports_the_index_of_a_section(self):
        prs = Presentation(element("p:presentation"), None)
        s1 = prs.sections.add_section("Intro")
        s2 = prs.sections.add_section("Body")
        s3 = prs.sections.add_section("Outro")

        assert prs.sections.index(s1) == 0
        assert prs.sections.index(s2) == 1
        assert prs.sections.index(s3) == 2

    def it_raises_ValueError_on_index_of_a_foreign_section(self):
        prs_a = Presentation(element("p:presentation"), None)
        prs_b = Presentation(element("p:presentation"), None)
        foreign = prs_b.sections.add_section("X")

        with pytest.raises(ValueError, match="not a member"):
            prs_a.sections.index(foreign)

    def it_finds_the_section_containing_a_given_slide(self, request):
        slide_a = instance_mock(request, Slide, name="slide_a")
        slide_b = instance_mock(request, Slide, name="slide_b")
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.side_effect = lambda rId: (
            slide_a if rId == "rId2" else slide_b
        )
        prs = Presentation(
            element(
                "p:presentation/p:sldIdLst/"
                "(p:sldId{id=256,r:id=rId2},p:sldId{id=257,r:id=rId3})"
            ),
            prs_part_,
        )
        s_intro = prs.sections.add_section("Intro", slides=[slide_a])
        s_body = prs.sections.add_section("Body", slides=[slide_b])

        assert prs.sections.find_containing(slide_a) == s_intro
        assert prs.sections.find_containing(slide_b) == s_body

    def but_find_containing_returns_None_when_slide_is_unassigned(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        # -- no section created at all --
        assert prs.sections.find_containing(slide_) is None

        # -- section exists but does not include the slide --
        prs.sections.add_section("Empty")
        assert prs.sections.find_containing(slide_) is None

    def and_find_containing_raises_on_foreign_slide(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        # -- lookup against rId2 resolves to a *different* slide --
        prs_part_.related_slide.return_value = instance_mock(request, Slide)
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )

        with pytest.raises(ValueError, match="does not belong"):
            prs.sections.find_containing(slide_)


class DescribeSection(object):
    """Unit-test suite for `pptx.presentation.Section`."""

    def it_can_rename_the_underlying_section(self):
        prs = Presentation(element("p:presentation"), None)
        section = prs.sections.add_section("Intro")

        section.name = "Summary"

        assert section.name == "Summary"
        assert prs._element.sectionLst is not None
        assert prs._element.sectionLst.section_lst[0].name == "Summary"

    def it_exposes_its_guid_id(self):
        prs = Presentation(element("p:presentation"), None)
        section = prs.sections.add_section("Intro", id="{11111111-2222-3333-4444-555555555555}")
        assert section.id == "{11111111-2222-3333-4444-555555555555}"

    def it_reports_no_slides_for_a_fresh_section(self):
        prs = Presentation(element("p:presentation"), None)
        section = prs.sections.add_section("Intro")
        assert section.slides == ()

    def it_can_add_a_slide_by_reference(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )

        section = prs.sections.add_section("Intro", slides=[slide_])

        assert section.slides == (slide_,)
        assert section._section_slide_ids == (256,)

    def it_is_a_noop_to_add_a_slide_already_in_the_section(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        section = prs.sections.add_section("Intro", slides=[slide_])

        section.add_slide(slide_)

        assert section._section_slide_ids == (256,)

    def it_raises_ValueError_when_adding_a_foreign_slide(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = instance_mock(request, Slide)
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        section = prs.sections.add_section("Intro")

        with pytest.raises(ValueError, match="does not belong"):
            section.add_slide(slide_)

    def it_can_remove_a_slide_from_the_section(self, request):
        slide_a = instance_mock(request, Slide, name="slide_a")
        slide_b = instance_mock(request, Slide, name="slide_b")
        prs_part_ = instance_mock(request, PresentationPart)
        # -- rId2 -> slide_a, rId3 -> slide_b --
        prs_part_.related_slide.side_effect = lambda rId: (slide_a if rId == "rId2" else slide_b)
        prs = Presentation(
            element(
                "p:presentation/p:sldIdLst/" "(p:sldId{id=256,r:id=rId2},p:sldId{id=257,r:id=rId3})"
            ),
            prs_part_,
        )
        section = prs.sections.add_section("Intro", slides=[slide_a, slide_b])

        section.remove_slide(slide_a)

        assert section._section_slide_ids == (257,)

    def it_raises_ValueError_removing_a_slide_not_in_the_section(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        section = prs.sections.add_section("Intro")  # -- empty section --

        with pytest.raises(ValueError, match="not a member"):
            section.remove_slide(slide_)

    def it_supports_equality_on_the_underlying_element(self):
        prs = Presentation(element("p:presentation"), None)
        section = prs.sections.add_section("Intro")
        # -- re-access via indexing should yield an equal Section --
        assert section == prs.sections[0]
        assert section != "not a section"

    def it_reports_its_index_within_the_section_list(self):
        prs = Presentation(element("p:presentation"), None)
        s1 = prs.sections.add_section("Intro")
        s2 = prs.sections.add_section("Body")

        assert s1.index == 0
        assert s2.index == 1

    def it_raises_ValueError_on_index_after_removal(self):
        prs = Presentation(element("p:presentation"), None)
        s1 = prs.sections.add_section("Intro")
        prs.sections.add_section("Body")

        prs.sections.remove(s1)

        with pytest.raises(ValueError, match="not a member"):
            _ = s1.index

    def it_can_move_a_section_before_another(self):
        prs = Presentation(element("p:presentation"), None)
        s1 = prs.sections.add_section("Intro")
        s2 = prs.sections.add_section("Body")
        s3 = prs.sections.add_section("Outro")

        s3.move_before(s1)

        names = [s.name for s in prs.sections]
        assert names == ["Outro", "Intro", "Body"]
        assert [s2, s3][0].name == "Body"  # -- still the same Section objects --

    def it_can_move_a_section_after_another(self):
        prs = Presentation(element("p:presentation"), None)
        s1 = prs.sections.add_section("Intro")
        s2 = prs.sections.add_section("Body")
        s3 = prs.sections.add_section("Outro")

        s1.move_after(s2)

        names = [s.name for s in prs.sections]
        assert names == ["Body", "Intro", "Outro"]
        # -- moving s1 after s3 pushes it to the end --
        s1.move_after(s3)
        assert [s.name for s in prs.sections] == ["Body", "Outro", "Intro"]

    def it_treats_moving_a_section_relative_to_itself_as_a_noop(self):
        prs = Presentation(element("p:presentation"), None)
        s1 = prs.sections.add_section("Intro")
        prs.sections.add_section("Body")

        s1.move_before(s1)
        s1.move_after(s1)

        assert [s.name for s in prs.sections] == ["Intro", "Body"]

    def it_raises_ValueError_moving_a_foreign_section(self):
        prs_a = Presentation(element("p:presentation"), None)
        prs_b = Presentation(element("p:presentation"), None)
        s_a = prs_a.sections.add_section("Intro")
        foreign = prs_b.sections.add_section("X")

        with pytest.raises(ValueError, match="not a member"):
            s_a.move_before(foreign)
        with pytest.raises(ValueError, match="not a member"):
            s_a.move_after(foreign)

    def it_raises_ValueError_when_adding_a_slide_already_in_another_section(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        prs.sections.add_section("Intro", slides=[slide_])
        target = prs.sections.add_section("Body")

        with pytest.raises(ValueError, match="already assigned to section 'Intro'"):
            target.add_slide(slide_)

    def it_can_move_a_slide_from_another_section(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        source = prs.sections.add_section("Intro", slides=[slide_])
        target = prs.sections.add_section("Body")

        target.move_slide(slide_)

        assert source._section_slide_ids == ()
        assert target._section_slide_ids == (256,)

    def and_move_slide_is_a_noop_when_already_in_target(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        target = prs.sections.add_section("Body", slides=[slide_])

        target.move_slide(slide_)

        assert target._section_slide_ids == (256,)

    def and_move_slide_adopts_an_unassigned_slide(self, request):
        slide_ = instance_mock(request, Slide)
        prs_part_ = instance_mock(request, PresentationPart)
        prs_part_.related_slide.return_value = slide_
        prs = Presentation(
            element("p:presentation/p:sldIdLst/p:sldId{id=256,r:id=rId2}"),
            prs_part_,
        )
        target = prs.sections.add_section("Body")

        target.move_slide(slide_)

        assert target._section_slide_ids == (256,)


class DescribeIssue694RegressionSections(object):
    """End-to-end regression suite covering issue #694.

    Issue #694 asked for API to:

    1. Get a list of all sections on a presentation, or look one up by name.
    2. Iterate the slides that belong to each section.
    3. Sort or otherwise operate on sections via objects with stable identity
       (e.g. a GUID id plus a display name).

    These tests exercise those three user-stories against the API landed by
    foundation F7 (see ``feat/foundation-f7-sections``) using a real
    round-tripped ``.pptx`` package — they are a *verification* of #694 being
    resolved by F7 rather than a re-implementation of the same behaviour.
    """

    @pytest.fixture(autouse=True)
    def _isolate_part_factory(self):
        """Repair ``PartFactory.part_type_for`` so sibling mocks don't poison us.

        Some tests in this repo (e.g. ``DescribePartFactory``) register mocked
        part classes into the module-level ``PartFactory.part_type_for`` dict
        and do not clean up. Because these round-trip tests exercise real
        ``.pptx`` loading, a leaked mock registration causes
        ``Mock object has no attribute 'slide'`` on ``related_slide(...)``.

        This fixture re-applies the canonical default mapping (the one
        registered in ``pptx/__init__.py``) before each test in this class and
        restores the as-observed state afterwards so the test is isolated in
        both directions.
        """
        from pptx import content_type_to_part_class_map
        from pptx.opc.package import PartFactory

        saved = dict(PartFactory.part_type_for)
        PartFactory.part_type_for.update(content_type_to_part_class_map)
        try:
            yield
        finally:
            PartFactory.part_type_for.clear()
            PartFactory.part_type_for.update(saved)

    def it_iterates_sections_and_their_slides_after_round_trip(self):
        prs = open_presentation()
        layout = prs.slide_layouts[6]  # -- blank layout --
        for _ in range(4):
            prs.slides.add_slide(layout)
        slides = list(prs.slides)

        intro = prs.sections.add_section("Intro", slides=slides[:2])
        body = prs.sections.add_section("Body", slides=slides[2:3])
        appendix = prs.sections.add_section("Appendix", slides=slides[3:])

        assert [s.name for s in prs.sections] == ["Intro", "Body", "Appendix"]
        assert intro.slides == (slides[0], slides[1])
        assert body.slides == (slides[2],)
        assert appendix.slides == (slides[3],)

        # -- round-trip and verify sections + slide membership survive --
        stream = io.BytesIO()
        prs.save(stream)
        prs2 = open_presentation(stream)

        assert [s.name for s in prs2.sections] == ["Intro", "Body", "Appendix"]
        # -- issue #694 core ask: iterate slides *per section* --
        reloaded_slides = list(prs2.slides)
        assert prs2.sections[0].slides == (reloaded_slides[0], reloaded_slides[1])
        assert prs2.sections[1].slides == (reloaded_slides[2],)
        assert prs2.sections[2].slides == (reloaded_slides[3],)

    def it_looks_up_a_section_by_name_for_slide_extraction(self):
        """User story 1: search presentation for a section by name, then iterate its slides."""
        prs = open_presentation()
        layout = prs.slide_layouts[6]
        for _ in range(3):
            prs.slides.add_slide(layout)
        all_slides = list(prs.slides)

        prs.sections.add_section("Appendix", slides=all_slides[2:])
        prs.sections.add_section("Intro", slides=all_slides[:2])

        found = prs.sections.get_by_name("Appendix")

        assert found is not None
        assert found.name == "Appendix"
        # -- the slides for that section are what the user wants to "extract" --
        assert [slide for slide in found.slides] == [all_slides[2]]

    def it_supports_sorting_section_objects_by_name(self):
        """User story 2: sort ``section`` objects by name and operate on them in order."""
        prs = open_presentation()
        prs.sections.add_section("Zeta")
        prs.sections.add_section("Alpha")
        prs.sections.add_section("Mu")

        sorted_sections = sorted(prs.sections, key=lambda s: s.name)

        assert [s.name for s in sorted_sections] == ["Alpha", "Mu", "Zeta"]
        # -- renaming through the sorted handles round-trips via the live tree --
        sorted_sections[0].name = "Alpha (renamed)"
        assert prs.sections.get_by_name("Alpha (renamed)") is not None

    def it_round_trips_an_author_supplied_section_guid(self):
        """Author-supplied GUIDs must survive save/reload so downstream tools can key on them."""
        author_id = "{ABCDEF01-2345-6789-ABCD-EF0123456789}"
        prs = open_presentation()
        layout = prs.slide_layouts[6]
        prs.slides.add_slide(layout)

        section = prs.sections.add_section(
            "Preserved", slides=[prs.slides[0]], id=author_id
        )
        assert section.id == author_id

        stream = io.BytesIO()
        prs.save(stream)
        prs2 = open_presentation(stream)

        reloaded = prs2.sections.get_by_id(author_id)
        assert reloaded is not None, "author-supplied section id did not survive round-trip"
        assert reloaded.id == author_id
        assert reloaded.name == "Preserved"

    def it_rejects_adding_a_section_with_a_duplicate_author_guid(self):
        """A second ``add_section(id=...)`` with the same GUID must raise so callers don't silently collide."""
        prs = open_presentation()
        author_id = "{11111111-2222-3333-4444-555555555555}"
        prs.sections.add_section("A", id=author_id)

        with pytest.raises(ValueError, match="already exists"):
            prs.sections.add_section("B", id=author_id)


class Describe_read_blob(object):
    """Unit-test suite for `pptx.presentation._read_blob` helper."""

    def it_reads_a_filesystem_path(self, tmp_path):
        p = tmp_path / "x.bin"
        p.write_bytes(b"hello")
        assert _read_blob(str(p)) == b"hello"

    def it_reads_a_file_like_object_and_rewinds_it(self):
        stream = io.BytesIO(b"hello")
        stream.read()  # advance past the end
        assert _read_blob(stream) == b"hello"

    def it_reads_a_non_seekable_file_like_object(self):
        class _NonSeekable:
            def __init__(self, data: bytes):
                self._data = data
                self._consumed = False

            def read(self) -> bytes:
                if self._consumed:
                    return b""
                self._consumed = True
                return self._data

        assert _read_blob(_NonSeekable(b"hello")) == b"hello"
