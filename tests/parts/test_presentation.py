"""Unit-test suite for `pptx.parts.presentation` module."""

from __future__ import annotations

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.package import Package
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.extprops import ExtendedPropertiesPart
from pptx.parts.font import FontPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import NotesMasterPart, SlideMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.slide import NotesMaster, Slide, SlideLayout, SlideMaster

from ..unitutil.cxml import element
from ..unitutil.mock import call, class_mock, instance_mock, method_mock, property_mock


class DescribePresentationPart(object):
    """Unit-test suite for `pptx.parts.presentation.PresentationPart` objects."""

    def it_provides_access_to_its_presentation(self, request):
        prs_ = instance_mock(request, Presentation)
        Presentation_ = class_mock(
            request, "pptx.parts.presentation.Presentation", return_value=prs_
        )
        prs_elm = element("p:presentation")
        prs_part = PresentationPart(None, None, None, prs_elm)

        prs = prs_part.presentation

        Presentation_.assert_called_once_with(prs_elm, prs_part)
        assert prs is prs_

    def it_provides_access_to_its_core_properties(self, request, package_):
        core_properties_ = instance_mock(request, CorePropertiesPart)
        package_.core_properties = core_properties_
        prs_part = PresentationPart(None, None, package_, None)

        assert prs_part.core_properties is core_properties_

    def it_provides_access_to_its_extended_properties(self, request, package_):
        extended_properties_ = instance_mock(request, ExtendedPropertiesPart)
        package_.extended_properties = extended_properties_
        prs_part = PresentationPart(None, None, package_, None)

        assert prs_part.extended_properties is extended_properties_

    def it_provides_access_to_an_existing_notes_master_part(
        self, notes_master_part_, part_related_by_
    ):
        """This is the first of a two-part test to cover the existing notes master case.

        The notes master not-present case follows.
        """
        prs_part = PresentationPart(None, None, None, None)
        part_related_by_.return_value = notes_master_part_

        notes_master_part = prs_part.notes_master_part

        prs_part.part_related_by.assert_called_once_with(prs_part, RT.NOTES_MASTER)
        assert notes_master_part is notes_master_part_

    def but_it_adds_a_notes_master_part_when_needed(
        self, request, package_, notes_master_part_, part_related_by_, relate_to_
    ):
        """This is the second of a two-part test to cover notes-master-not-present case.

        The notes master present case is just above.
        """
        NotesMasterPart_ = class_mock(request, "pptx.parts.presentation.NotesMasterPart")
        NotesMasterPart_.create_default.return_value = notes_master_part_
        part_related_by_.side_effect = KeyError
        prs_part = PresentationPart(None, None, package_, None)

        notes_master_part = prs_part.notes_master_part

        NotesMasterPart_.create_default.assert_called_once_with(package_)
        relate_to_.assert_called_once_with(prs_part, notes_master_part_, RT.NOTES_MASTER)
        assert notes_master_part is notes_master_part_

    def it_provides_access_to_its_notes_master(self, request, notes_master_part_):
        notes_master_ = instance_mock(request, NotesMaster)
        property_mock(
            request,
            PresentationPart,
            "notes_master_part",
            return_value=notes_master_part_,
        )
        notes_master_part_.notes_master = notes_master_
        prs_part = PresentationPart(None, None, None, None)

        assert prs_part.notes_master is notes_master_

    def it_provides_access_to_a_related_slide(self, request, slide_, related_part_):
        slide_part_ = instance_mock(request, SlidePart, slide=slide_)
        related_part_.return_value = slide_part_
        prs_part = PresentationPart(None, None, None, None)

        slide = prs_part.related_slide("rId42")

        related_part_.assert_called_once_with(prs_part, "rId42")
        assert slide is slide_

    def it_provides_access_to_a_related_master(self, request, slide_master_, related_part_):
        slide_master_part_ = instance_mock(request, SlideMasterPart, slide_master=slide_master_)
        related_part_.return_value = slide_master_part_
        prs_part = PresentationPart(None, None, None, None)

        slide_master = prs_part.related_slide_master("rId42")

        related_part_.assert_called_once_with(prs_part, "rId42")
        assert slide_master is slide_master_

    def it_can_rename_related_slide_parts(self, request, related_part_):
        rIds = tuple("rId%d" % n for n in range(5, 0, -1))
        slide_parts = tuple(instance_mock(request, SlidePart) for _ in range(5))
        related_part_.side_effect = iter(slide_parts)
        prs_part = PresentationPart(None, None, None, None)

        prs_part.rename_slide_parts(rIds)

        assert related_part_.call_args_list == [call(prs_part, rId) for rId in rIds]
        assert [s.partname for s in slide_parts] == [
            PackURI("/ppt/slides/slide%d.xml" % (i + 1)) for i in range(len(rIds))
        ]

    def it_can_save_the_package_to_a_file(self, package_):
        PresentationPart(None, None, package_, None).save("prs.pptx")
        package_.save.assert_called_once_with("prs.pptx", None, password=None)

    def and_it_forwards_zip_date_time_to_the_package(self, package_):
        zdt = (2024, 6, 15, 9, 30, 0)
        PresentationPart(None, None, package_, None).save("prs.pptx", zdt)
        package_.save.assert_called_once_with("prs.pptx", zdt, password=None)

    def it_can_save_a_password_protected_package_to_a_file(self, package_):
        PresentationPart(None, None, package_, None).save("prs.pptx", password="s3cret")
        package_.save.assert_called_once_with("prs.pptx", None, password="s3cret")

    def and_it_forwards_both_zip_date_time_and_password_to_the_package(self, package_):
        zdt = (2024, 6, 15, 9, 30, 0)
        PresentationPart(None, None, package_, None).save(
            "prs.pptx", zdt, password="s3cret"
        )
        package_.save.assert_called_once_with("prs.pptx", zdt, password="s3cret")

    def it_can_save_the_package_as_flat_opc_xml(self, package_):
        PresentationPart(None, None, package_, None).save_flat_xml("prs.xml")
        package_.save_flat_xml.assert_called_once_with("prs.xml")

    def it_can_save_the_package_as_ppsx(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)

        # -- stub `package.save` to capture the content-type that was in effect
        # -- at the moment of the save call (the whole point of `save_ppsx`).
        observed_cts: list[str] = []

        def _observe_save(*args, **kwargs):  # type: ignore[no-untyped-def]
            observed_cts.append(prs_part.content_type)

        package_.save.side_effect = _observe_save

        prs_part.save_ppsx("prs.ppsx")

        # -- during the save, the content type is the slideshow variant --
        assert observed_cts == [CT.PML_SLIDESHOW_MAIN]
        # -- after the save, it is restored to the original --
        assert prs_part.content_type == CT.PML_PRESENTATION_MAIN
        package_.save.assert_called_once_with("prs.ppsx", None, password=None)

    def and_save_ppsx_forwards_zip_date_time_and_password(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        zdt = (2024, 6, 15, 9, 30, 0)
        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)

        prs_part.save_ppsx("prs.ppsx", zdt, password="s3cret")

        package_.save.assert_called_once_with("prs.ppsx", zdt, password="s3cret")
        # -- original content type is restored even when encryption is used --
        assert prs_part.content_type == CT.PML_PRESENTATION_MAIN

    def and_save_ppsx_restores_content_type_when_package_save_raises(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)
        package_.save.side_effect = RuntimeError("boom")

        with pytest.raises(RuntimeError, match="boom"):
            prs_part.save_ppsx("prs.ppsx")

        # -- original content type is restored via `finally:` even on error --
        assert prs_part.content_type == CT.PML_PRESENTATION_MAIN

    # -- issue #976 save .pptm / .ppsm auto-promotes content type --------

    def it_swaps_content_type_when_saving_with_pptm_extension(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)

        observed_cts: list[str] = []

        def _observe_save(*args, **kwargs):  # type: ignore[no-untyped-def]
            observed_cts.append(prs_part.content_type)

        package_.save.side_effect = _observe_save

        prs_part.save("deck.pptm")

        # -- during the save, content type is the macro-enabled variant --
        assert observed_cts == [CT.PML_PRES_MACRO_MAIN]
        # -- after the save, original content type is restored --
        assert prs_part.content_type == CT.PML_PRESENTATION_MAIN

    def it_swaps_content_type_when_saving_with_ppsm_extension(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)
        observed_cts: list[str] = []

        def _observe_save(*args, **kwargs):  # type: ignore[no-untyped-def]
            observed_cts.append(prs_part.content_type)

        package_.save.side_effect = _observe_save

        prs_part.save("deck.ppsm")

        assert observed_cts == [CT.PML_SLIDESHOW_MACRO_MAIN]
        assert prs_part.content_type == CT.PML_PRESENTATION_MAIN

    def it_pptm_extension_matching_is_case_insensitive(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)
        observed_cts: list[str] = []

        def _observe_save(*args, **kwargs):  # type: ignore[no-untyped-def]
            observed_cts.append(prs_part.content_type)

        package_.save.side_effect = _observe_save

        prs_part.save("DECK.PPTM")

        assert observed_cts == [CT.PML_PRES_MACRO_MAIN]

    def but_it_does_not_swap_when_saving_pptx(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)
        observed_cts: list[str] = []

        def _observe_save(*args, **kwargs):  # type: ignore[no-untyped-def]
            observed_cts.append(prs_part.content_type)

        package_.save.side_effect = _observe_save

        prs_part.save("deck.pptx")

        assert observed_cts == [CT.PML_PRESENTATION_MAIN]

    def but_it_does_not_swap_when_saving_to_a_stream(self, package_):
        import io as _io

        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)
        observed_cts: list[str] = []

        def _observe_save(*args, **kwargs):  # type: ignore[no-untyped-def]
            observed_cts.append(prs_part.content_type)

        package_.save.side_effect = _observe_save

        prs_part.save(_io.BytesIO())

        # -- a stream carries no extension to sniff; original CT preserved
        assert observed_cts == [CT.PML_PRESENTATION_MAIN]

    def it_is_a_noop_swap_when_already_macro_enabled(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        # -- file was loaded as .pptm, already has the macro content type
        prs_part = PresentationPart(None, CT.PML_PRES_MACRO_MAIN, package_, None)
        observed_cts: list[str] = []

        def _observe_save(*args, **kwargs):  # type: ignore[no-untyped-def]
            observed_cts.append(prs_part.content_type)

        package_.save.side_effect = _observe_save

        prs_part.save("deck.pptm")

        # -- content type unchanged; the fast path is taken
        assert observed_cts == [CT.PML_PRES_MACRO_MAIN]
        assert prs_part.content_type == CT.PML_PRES_MACRO_MAIN

    def it_restores_content_type_when_save_raises_during_pptm_save(self, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT

        prs_part = PresentationPart(None, CT.PML_PRESENTATION_MAIN, package_, None)
        package_.save.side_effect = RuntimeError("boom")

        with pytest.raises(RuntimeError, match="boom"):
            prs_part.save("deck.pptm")

        assert prs_part.content_type == CT.PML_PRESENTATION_MAIN

    def it_can_add_a_font_part_and_relate_it(self, request, package_, relate_to_):
        font_blob = b"fake-ttf"
        font_part_ = instance_mock(request, FontPart)
        FontPart_ = class_mock(request, "pptx.parts.presentation.FontPart")
        FontPart_.new.return_value = font_part_
        relate_to_.return_value = "rId7"
        prs_part = PresentationPart(None, None, package_, None)

        rId = prs_part.add_embedded_font(font_blob)

        FontPart_.new.assert_called_once_with(font_blob, package_)
        prs_part.relate_to.assert_called_once_with(prs_part, font_part_, RT.FONT)
        assert rId == "rId7"

    def it_can_add_a_new_slide(self, request, package_, slide_part_, slide_, relate_to_):
        slide_layout_ = instance_mock(request, SlideLayout)
        partname = PackURI("/ppt/slides/slide9.xml")
        property_mock(request, PresentationPart, "_next_slide_partname", return_value=partname)
        SlidePart_ = class_mock(request, "pptx.parts.presentation.SlidePart")
        SlidePart_.new.return_value = slide_part_
        relate_to_.return_value = "rId42"
        slide_layout_part_ = slide_layout_.part
        slide_part_.slide = slide_
        prs_part = PresentationPart(None, None, package_, None)

        rId, slide = prs_part.add_slide(slide_layout_)

        SlidePart_.new.assert_called_once_with(partname, package_, slide_layout_part_)
        prs_part.relate_to.assert_called_once_with(prs_part, slide_part_, RT.SLIDE)
        assert rId == "rId42"
        assert slide is slide_

    def it_can_add_a_slide_cloned_from_another_presentation(
        self, request, package_, slide_part_, slide_, relate_to_
    ):
        source_slide_ = instance_mock(request, Slide)
        source_slide_part_ = instance_mock(request, SlidePart)
        source_slide_.part = source_slide_part_
        slide_layout_ = instance_mock(request, SlideLayout)
        slide_layout_part_ = slide_layout_.part

        partname = PackURI("/ppt/slides/slide9.xml")
        property_mock(request, PresentationPart, "_next_slide_partname", return_value=partname)
        SlidePart_ = class_mock(request, "pptx.parts.presentation.SlidePart")
        SlidePart_.clone_from.return_value = slide_part_
        slide_part_.slide = slide_
        relate_to_.return_value = "rId42"
        prs_part = PresentationPart(None, None, package_, None)

        rId, slide = prs_part.add_slide_from_external(source_slide_, slide_layout_)

        SlidePart_.clone_from.assert_called_once_with(
            source_slide_part_, partname, package_, slide_layout_part_
        )
        prs_part.relate_to.assert_called_once_with(prs_part, slide_part_, RT.SLIDE)
        assert rId == "rId42"
        assert slide is slide_

    def it_can_duplicate_a_slide_within_this_presentation(
        self, request, package_, slide_part_, slide_, relate_to_
    ):
        source_slide_ = instance_mock(request, Slide)
        source_slide_part_ = instance_mock(request, SlidePart)
        source_slide_.part = source_slide_part_
        partname = PackURI("/ppt/slides/slide9.xml")
        property_mock(request, PresentationPart, "_next_slide_partname", return_value=partname)
        SlidePart_ = class_mock(request, "pptx.parts.presentation.SlidePart")
        SlidePart_.clone_within.return_value = slide_part_
        slide_part_.slide = slide_
        relate_to_.return_value = "rId42"
        prs_part = PresentationPart(None, None, package_, None)

        rId, slide = prs_part.duplicate_slide(source_slide_)

        SlidePart_.clone_within.assert_called_once_with(source_slide_part_, partname)
        prs_part.relate_to.assert_called_once_with(prs_part, slide_part_, RT.SLIDE)
        assert rId == "rId42"
        assert slide is slide_

    def it_finds_the_slide_id_of_a_slide_part(self, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        related_part_.side_effect = iter((None, slide_part_, None))
        prs_part = PresentationPart(None, None, None, prs_elm)

        _slide_id = prs_part.slide_id(slide_part_)

        assert related_part_.call_args_list == [
            call(prs_part, "a"),
            call(prs_part, "b"),
        ]
        assert _slide_id == 257

    def it_raises_on_slide_id_not_found(self, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        related_part_.return_value = "not the slide you're looking for"
        prs_part = PresentationPart(None, None, None, prs_elm)

        with pytest.raises(ValueError):
            prs_part.slide_id(slide_part_)

    @pytest.mark.parametrize("is_present", (True, False))
    def it_finds_a_slide_by_slide_id(self, is_present, slide_, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        slide_id = 257 if is_present else 666
        expected_value = slide_ if is_present else None
        related_part_.return_value = slide_part_
        slide_part_.slide = slide_
        prs_part = PresentationPart(None, None, None, prs_elm)

        slide = prs_part.get_slide(slide_id)

        assert slide == expected_value

    def it_knows_the_next_slide_partname_to_help(self):
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId,p:sldId)")
        prs_part = PresentationPart(None, None, None, prs_elm)

        assert prs_part._next_slide_partname == PackURI("/ppt/slides/slide3.xml")

    # fixture components ---------------------------------------------

    @pytest.fixture
    def notes_master_part_(self, request):
        return instance_mock(request, NotesMasterPart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_related_by_(self, request):
        return method_mock(request, PresentationPart, "part_related_by")

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(request, PresentationPart, "relate_to")

    @pytest.fixture
    def related_part_(self, request):
        return method_mock(request, PresentationPart, "related_part")

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)
