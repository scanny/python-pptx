"""Unit-test suite for `pptx.parts.slide` module."""

from __future__ import annotations

import pytest

from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE as XCT
from pptx.enum.shapes import PROG_ID
from pptx.media import Audio, Video
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide, CT_Slide
from pptx.oxml.theme import CT_OfficeStyleSheet
from pptx.package import Package
from pptx.parts.chart import ChartPart
from pptx.parts.embeddedpackage import EmbeddedPackagePart
from pptx.parts.image import Image, ImagePart
from pptx.parts.media import MediaPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import (
    BaseSlidePart,
    NotesMasterPart,
    NotesSlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
)
from pptx.slide import NotesMaster, NotesSlide, Slide, SlideLayout, SlideMaster

from ..unitutil.cxml import element
from ..unitutil.file import absjoin, test_file_dir
from ..unitutil.mock import (
    call,
    class_mock,
    initializer_mock,
    instance_mock,
    method_mock,
)


class DescribeBaseSlidePart(object):
    """Unit-test suite for `pptx.parts.slide.BaseSlidePart` objects."""

    def it_knows_its_name(self):
        slide_part = BaseSlidePart(None, None, None, element("p:sld/p:cSld{name=Foobar}"))
        assert slide_part.name == "Foobar"

    def it_can_get_a_related_image_by_rId(self, request, image_part_):
        image_ = instance_mock(request, Image)
        image_part_.image = image_
        related_part_ = method_mock(
            request, BaseSlidePart, "related_part", return_value=image_part_
        )
        slide_part = BaseSlidePart(None, None, None, None)

        image = slide_part.get_image("rId42")

        related_part_.assert_called_once_with(slide_part, "rId42")
        assert image is image_

    def it_can_add_an_image_part(self, request, image_part_):
        package_ = instance_mock(request, Package)
        package_.get_or_add_image_part.return_value = image_part_
        relate_to_ = method_mock(request, BaseSlidePart, "relate_to", return_value="rId6")
        slide_part = BaseSlidePart(None, None, package_, None)

        image_part, rId = slide_part.get_or_add_image_part("foobar.png")

        package_.get_or_add_image_part.assert_called_once_with("foobar.png")
        relate_to_.assert_called_once_with(slide_part, image_part_, RT.IMAGE)
        assert image_part is image_part_
        assert rId == "rId6"

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)


class DescribeNotesMasterPart(object):
    """Unit-test suite for `pptx.parts.slide.NotesMasterPart` objects."""

    def it_can_create_a_notes_master_part(self, request, package_, notes_master_part_, theme_part_):
        method_mock(
            request,
            NotesMasterPart,
            "_new",
            autospec=False,
            return_value=notes_master_part_,
        )
        method_mock(
            request,
            NotesMasterPart,
            "_new_theme_part",
            autospec=False,
            return_value=theme_part_,
        )
        notes_master_part = NotesMasterPart.create_default(package_)

        NotesMasterPart._new.assert_called_once_with(package_)
        NotesMasterPart._new_theme_part.assert_called_once_with(package_)
        notes_master_part.relate_to.assert_called_once_with(theme_part_, RT.THEME)
        assert notes_master_part is notes_master_part_

    def it_provides_access_to_its_notes_master(self, request):
        notes_master_ = instance_mock(request, NotesMaster)
        NotesMaster_ = class_mock(
            request, "pptx.parts.slide.NotesMaster", return_value=notes_master_
        )
        notesMaster = element("p:notesMaster")
        notes_master_part = NotesMasterPart(None, None, None, notesMaster)

        notes_master = notes_master_part.notes_master

        NotesMaster_.assert_called_once_with(notesMaster, notes_master_part)
        assert notes_master is notes_master_

    def it_creates_a_new_notes_master_part_to_help(self, request, package_, notes_master_part_):
        NotesMasterPart_ = class_mock(
            request, "pptx.parts.slide.NotesMasterPart", return_value=notes_master_part_
        )
        notesMaster = element("p:notesMaster")
        method_mock(
            request,
            CT_NotesMaster,
            "new_default",
            autospec=False,
            return_value=notesMaster,
        )

        notes_master_part = NotesMasterPart._new(package_)

        CT_NotesMaster.new_default.assert_called_once_with()
        NotesMasterPart_.assert_called_once_with(
            PackURI("/ppt/notesMasters/notesMaster1.xml"),
            CT.PML_NOTES_MASTER,
            package_,
            notesMaster,
        )
        assert notes_master_part is notes_master_part_

    def it_creates_a_new_theme_part_to_help(self, request, package_, theme_part_):
        ThemePart_ = class_mock(request, "pptx.parts.slide.ThemePart", return_value=theme_part_)
        theme_elm = element("p:theme")
        method_mock(
            request,
            CT_OfficeStyleSheet,
            "new_default",
            autospec=False,
            return_value=theme_elm,
        )
        pn_tmpl = "/ppt/theme/theme%d.xml"
        partname = PackURI("/ppt/theme/theme2.xml")
        package_.next_partname.return_value = partname

        theme_part = NotesMasterPart._new_theme_part(package_)

        package_.next_partname.assert_called_once_with(pn_tmpl)
        CT_OfficeStyleSheet.new_default.assert_called_once_with()
        ThemePart_.assert_called_once_with(partname, CT.OFC_THEME, package_, theme_elm)
        assert theme_part is theme_part_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def notes_master_part_(self, request):
        return instance_mock(request, NotesMasterPart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def theme_part_(self, request):
        return instance_mock(request, Part)


class DescribeNotesSlidePart(object):
    """Unit-test suite for `pptx.parts.slide.NotesSlidePart` objects."""

    def it_can_create_a_notes_slide_part(
        self,
        request,
        package_,
        slide_part_,
        notes_master_part_,
        notes_slide_,
        notes_master_,
        notes_slide_part_,
    ):
        presentation_part_ = instance_mock(request, PresentationPart)
        package_.presentation_part = presentation_part_
        presentation_part_.notes_master_part = notes_master_part_
        _add_notes_slide_part_ = method_mock(
            request,
            NotesSlidePart,
            "_add_notes_slide_part",
            autospec=False,
            return_value=notes_slide_part_,
        )
        notes_slide_part_.notes_slide = notes_slide_
        notes_master_part_.notes_master = notes_master_

        notes_slide_part = NotesSlidePart.new(package_, slide_part_)

        _add_notes_slide_part_.assert_called_once_with(package_, slide_part_, notes_master_part_)
        notes_slide_.clone_master_placeholders.assert_called_once_with(notes_master_)
        assert notes_slide_part is notes_slide_part_

    def it_provides_access_to_the_notes_master(self, request, notes_master_, notes_master_part_):
        part_related_by_ = method_mock(
            request, NotesSlidePart, "part_related_by", return_value=notes_master_part_
        )
        notes_slide_part = NotesSlidePart(None, None, None, None)
        notes_master_part_.notes_master = notes_master_

        notes_master = notes_slide_part.notes_master

        part_related_by_.assert_called_once_with(notes_slide_part, RT.NOTES_MASTER)
        assert notes_master is notes_master_

    def it_provides_access_to_its_notes_slide(self, request, notes_slide_):
        NotesSlide_ = class_mock(request, "pptx.parts.slide.NotesSlide", return_value=notes_slide_)
        notes = element("p:notes")
        notes_slide_part = NotesSlidePart(None, None, None, notes)

        notes_slide = notes_slide_part.notes_slide

        NotesSlide_.assert_called_once_with(notes, notes_slide_part)
        assert notes_slide is notes_slide_

    def it_adds_a_notes_slide_part_to_help(
        self, request, package_, slide_part_, notes_master_part_, notes_slide_part_
    ):
        NotesSlidePart_ = class_mock(
            request, "pptx.parts.slide.NotesSlidePart", return_value=notes_slide_part_
        )
        notes = element("p:notes")
        new_ = method_mock(request, CT_NotesSlide, "new", autospec=False, return_value=notes)
        package_.next_partname.return_value = PackURI("/ppt/notesSlides/notesSlide42.xml")

        notes_slide_part = NotesSlidePart._add_notes_slide_part(
            package_, slide_part_, notes_master_part_
        )

        package_.next_partname.assert_called_once_with("/ppt/notesSlides/notesSlide%d.xml")
        new_.assert_called_once_with()
        NotesSlidePart_.assert_called_once_with(
            PackURI("/ppt/notesSlides/notesSlide42.xml"),
            CT.PML_NOTES_SLIDE,
            package_,
            notes,
        )
        assert notes_slide_part_.relate_to.call_args_list == [
            call(notes_master_part_, RT.NOTES_MASTER),
            call(slide_part_, RT.SLIDE),
        ]
        assert notes_slide_part is notes_slide_part_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def notes_master_(self, request):
        return instance_mock(request, NotesMaster)

    @pytest.fixture
    def notes_master_part_(self, request):
        return instance_mock(request, NotesMasterPart)

    @pytest.fixture
    def notes_slide_(self, request):
        return instance_mock(request, NotesSlide)

    @pytest.fixture
    def notes_slide_part_(self, request):
        return instance_mock(request, NotesSlidePart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)


class DescribeSlidePart(object):
    """Unit-test suite for `pptx.parts.slide.SlidePart` objects."""

    def it_knows_its_slide_id(self, slide_id_fixture):
        slide_part, presentation_part_, slide_id = slide_id_fixture
        _slide_id = slide_part.slide_id
        presentation_part_.slide_id.assert_called_once_with(slide_part)
        assert _slide_id is slide_id

    def it_knows_whether_it_has_a_notes_slide(self, has_notes_slide_fixture):
        slide_part, expected_value = has_notes_slide_fixture
        value = slide_part.has_notes_slide
        slide_part.part_related_by.assert_called_once_with(slide_part, RT.NOTES_SLIDE)
        assert value is expected_value

    def it_can_add_a_chart_part(self, request, package_, relate_to_):
        chart_data_ = instance_mock(request, ChartData)
        chart_part_ = instance_mock(request, ChartPart)
        ChartPart_ = class_mock(request, "pptx.parts.slide.ChartPart")
        ChartPart_.new.return_value = chart_part_
        relate_to_.return_value = "rId42"
        slide_part = SlidePart(None, None, package_, None)

        rId = slide_part.add_chart_part(XCT.RADAR, chart_data_)

        ChartPart_.new.assert_called_once_with(XCT.RADAR, chart_data_, package_)
        relate_to_.assert_called_once_with(slide_part, chart_part_, RT.CHART)
        assert rId == "rId42"

    @pytest.mark.parametrize(
        "prog_id, rel_type",
        (
            (PROG_ID.DOCX, RT.PACKAGE),
            (PROG_ID.PPTX, RT.PACKAGE),
            (PROG_ID.XLSX, RT.PACKAGE),
            ("Foo.Bar.18", RT.OLE_OBJECT),
        ),
    )
    def it_can_add_an_embedded_ole_object_part(
        self, request, package_, relate_to_, prog_id, rel_type
    ):
        _blob_from_file_ = method_mock(
            request, SlidePart, "_blob_from_file", return_value=b"012345"
        )
        embedded_package_part_ = instance_mock(request, EmbeddedPackagePart)
        EmbeddedPackagePart_ = class_mock(request, "pptx.parts.slide.EmbeddedPackagePart")
        EmbeddedPackagePart_.factory.return_value = embedded_package_part_
        relate_to_.return_value = "rId9"
        slide_part = SlidePart(None, None, package_, None)

        _rId = slide_part.add_embedded_ole_object_part(prog_id, "workbook.xlsx")

        _blob_from_file_.assert_called_once_with(slide_part, "workbook.xlsx")
        EmbeddedPackagePart_.factory.assert_called_once_with(
            prog_id, b"012345", package_, None
        )
        relate_to_.assert_called_once_with(slide_part, embedded_package_part_, rel_type)
        assert _rId == "rId9"

    def it_can_get_or_add_a_video_part(self, package_, video_, relate_to_, media_part_):
        media_rId, video_rId = "rId1", "rId2"
        package_.get_or_add_media_part.return_value = media_part_
        relate_to_.side_effect = [media_rId, video_rId]
        slide_part = SlidePart(None, None, package_, None)

        result = slide_part.get_or_add_video_media_part(video_)

        package_.get_or_add_media_part.assert_called_once_with(video_)
        assert relate_to_.call_args_list == [
            call(slide_part, media_part_, RT.MEDIA),
            call(slide_part, media_part_, RT.VIDEO),
        ]
        assert result == (media_rId, video_rId)

    def it_can_get_or_add_a_sound_part(
        self, request, package_, relate_to_, media_part_
    ):
        audio_ = instance_mock(request, Audio)
        package_.get_or_add_media_part.return_value = media_part_
        relate_to_.return_value = "rId7"
        slide_part = SlidePart(None, None, package_, None)

        rId = slide_part.get_or_add_sound_media_part(audio_)

        package_.get_or_add_media_part.assert_called_once_with(audio_)
        relate_to_.assert_called_once_with(slide_part, media_part_, RT.AUDIO)
        assert rId == "rId7"

    def it_can_create_a_new_slide_part(self, request, package_, relate_to_):
        partname = PackURI("/foobar.xml")
        _init_ = initializer_mock(request, SlidePart)
        slide_layout_part_ = instance_mock(request, SlideLayoutPart)
        CT_Slide_ = class_mock(request, "pptx.parts.slide.CT_Slide")
        CT_Slide_.new.return_value = sld = element("c:sld")

        slide_part = SlidePart.new(partname, package_, slide_layout_part_)

        _init_.assert_called_once_with(slide_part, partname, CT.PML_SLIDE, package_, sld)
        slide_part.relate_to.assert_called_once_with(
            slide_part, slide_layout_part_, RT.SLIDE_LAYOUT
        )
        assert isinstance(slide_part, SlidePart)

    def it_can_clone_a_slide_part_from_another_presentation(self, request):
        from unittest.mock import MagicMock

        # -- Source SlidePart with: layout rel + image rel + hyperlink rel. --
        src_sld = element("p:sld/p:cSld/p:spTree/p:pic/p:blipFill/a:blip{r:embed=rId9}")
        src_part = SlidePart(PackURI("/ppt/slides/slide1.xml"), CT.PML_SLIDE, None, src_sld)
        layout_rel = MagicMock(reltype=RT.SLIDE_LAYOUT, is_external=False)
        image_rel = MagicMock(reltype=RT.IMAGE, is_external=False)
        image_rel.target_part = instance_mock(request, ImagePart, blob=b"PNGBYTES")
        hl_rel = MagicMock(
            reltype=RT.HYPERLINK, is_external=True, target_ref="https://example.com/"
        )
        src_part.rels._rels.update(  # type: ignore[attr-defined]
            {"rId1": layout_rel, "rId9": image_rel, "rId7": hl_rel}
        )

        # -- Target package: stub get_or_add_image_part to return a fake image part. --
        target_package_ = instance_mock(request, Package)
        target_image_part_ = instance_mock(request, ImagePart)
        target_package_.get_or_add_image_part.return_value = target_image_part_
        slide_layout_part_ = instance_mock(request, SlideLayoutPart)
        # -- Stub SlidePart.relate_to so no real relationship graph is touched. --
        relate_to_ = method_mock(request, SlidePart, "relate_to", autospec=True)
        relate_to_.side_effect = ["rIdL", "rIdH", "rIdI"]

        partname = PackURI("/ppt/slides/slide2.xml")
        new_part = SlidePart.clone_from(src_part, partname, target_package_, slide_layout_part_)

        assert isinstance(new_part, SlidePart)
        assert new_part.partname == partname
        # -- Shape tree is deep-copied, not aliased. --
        assert new_part._element is not src_sld
        # -- `r:embed` on the cloned blip was remapped from rId9 to the new image rId. --
        embed_attr = new_part._element.xpath(".//a:blip/@r:embed")[0]
        assert embed_attr == "rIdH"
        # -- Image part was fetched from target package, not the source. --
        target_package_.get_or_add_image_part.assert_called_once()

    def it_can_clone_a_slide_part_within_the_same_package(self, request):
        """SlidePart.clone_within duplicates a slide in its own package via F1."""
        # -- Source SlidePart whose shape tree holds an `r:embed` reference. --
        src_sld = element("p:sld/p:cSld/p:spTree/p:pic/p:blipFill/a:blip{r:embed=rId9}")
        package_ = instance_mock(request, Package)
        src_part = SlidePart(PackURI("/ppt/slides/slide1.xml"), CT.PML_SLIDE, package_, src_sld)

        # -- Stub the source's layout lookup and the cloner's remap result. --
        slide_layout_part_ = instance_mock(request, SlideLayoutPart)
        part_related_by_ = method_mock(
            request, SlidePart, "part_related_by", return_value=slide_layout_part_, autospec=True
        )
        relate_to_ = method_mock(
            request, SlidePart, "relate_to", return_value="rIdL", autospec=True
        )
        remapped_sld = element("p:sld/p:cSld/p:spTree/p:pic/p:blipFill/a:blip{r:embed=rIdNEW}")
        PartRelationshipCloner_ = class_mock(
            request, "pptx.parts.slide.PartRelationshipCloner"
        )
        PartRelationshipCloner_.clone.return_value = remapped_sld

        partname = PackURI("/ppt/slides/slide2.xml")
        new_part = SlidePart.clone_within(src_part, partname)

        # -- a new SlidePart instance bound to the target partname and package --
        assert isinstance(new_part, SlidePart)
        assert new_part.partname == partname
        assert new_part.package is package_
        # -- the layout rel was established via part_related_by + relate_to --
        part_related_by_.assert_called_once_with(src_part, RT.SLIDE_LAYOUT)
        relate_to_.assert_called_once_with(new_part, slide_layout_part_, RT.SLIDE_LAYOUT)
        # -- the cloner was invoked with the canonical (src_part, new_part, src_sld) triple --
        PartRelationshipCloner_.clone.assert_called_once_with(src_part, new_part, src_sld)
        # -- and the cloner's return value replaced the new part's element tree. --
        assert new_part._element is remapped_sld

    def it_drops_the_notes_slide_when_cloning(self, request):
        from unittest.mock import MagicMock

        src_sld = element("p:sld/p:cSld/p:spTree")
        src_part = SlidePart(PackURI("/ppt/slides/slide1.xml"), CT.PML_SLIDE, None, src_sld)
        src_part.rels._rels.update(  # type: ignore[attr-defined]
            {
                "rId1": MagicMock(reltype=RT.SLIDE_LAYOUT, is_external=False),
                "rId2": MagicMock(reltype=RT.NOTES_SLIDE, is_external=False),
            }
        )
        target_package_ = instance_mock(request, Package)
        slide_layout_part_ = instance_mock(request, SlideLayoutPart)
        relate_to_ = method_mock(request, SlidePart, "relate_to", autospec=True)
        relate_to_.return_value = "rIdL"

        new_part = SlidePart.clone_from(
            src_part,
            PackURI("/ppt/slides/slide2.xml"),
            target_package_,
            slide_layout_part_,
        )

        # -- No notes-slide relationship is established on the new part. --
        relate_to_.assert_called_once_with(new_part, slide_layout_part_, RT.SLIDE_LAYOUT)

    def it_clones_a_chart_rel_via_ChartPart_clone_from(self, request):
        """clone_from promotes charts to full fidelity via ChartPart.clone_from (F1+F5)."""
        from unittest.mock import MagicMock

        # -- Source SlidePart with a layout rel and a chart rel. --
        src_sld = element("p:sld/p:cSld/p:spTree")
        src_part = SlidePart(
            PackURI("/ppt/slides/slide1.xml"), CT.PML_SLIDE, None, src_sld
        )
        layout_rel = MagicMock(reltype=RT.SLIDE_LAYOUT, is_external=False)
        source_chart_part_ = instance_mock(request, ChartPart)
        chart_rel = MagicMock(
            reltype=RT.CHART, is_external=False, target_part=source_chart_part_
        )
        src_part.rels._rels.update(  # type: ignore[attr-defined]
            {"rId1": layout_rel, "rId2": chart_rel}
        )

        target_package_ = instance_mock(request, Package)
        slide_layout_part_ = instance_mock(request, SlideLayoutPart)
        cloned_chart_part_ = instance_mock(request, ChartPart)
        ChartPart_ = class_mock(request, "pptx.parts.slide.ChartPart")
        ChartPart_.clone_from.return_value = cloned_chart_part_

        relate_to_ = method_mock(request, SlidePart, "relate_to", autospec=True)
        relate_to_.side_effect = ["rIdL", "rIdC"]

        new_part = SlidePart.clone_from(
            src_part,
            PackURI("/ppt/slides/slide2.xml"),
            target_package_,
            slide_layout_part_,
        )

        # -- ChartPart.clone_from was invoked (source chart, target package). --
        ChartPart_.clone_from.assert_called_once_with(source_chart_part_, target_package_)
        # -- The new slide part has a CHART rel to the cloned chart part. --
        assert call(new_part, cloned_chart_part_, RT.CHART) in relate_to_.mock_calls

    def it_clones_media_rels_sharing_a_single_target_MediaPart(self, request):
        """VIDEO + MEDIA on the same source MediaPart must share one clone in the target."""
        from unittest.mock import MagicMock

        src_sld = element("p:sld/p:cSld/p:spTree")
        src_part = SlidePart(
            PackURI("/ppt/slides/slide1.xml"), CT.PML_SLIDE, None, src_sld
        )
        # -- One MediaPart, two rels (MEDIA + VIDEO), same source part target. --
        src_pkg_sentinel = object()
        src_media_part = MagicMock(spec=MediaPart)
        src_media_part.blob = b"MP4DATA"
        src_media_part.content_type = "video/mp4"
        src_media_part.package = src_pkg_sentinel
        src_media_part.partname = PackURI("/ppt/media/media1.mp4")
        layout_rel = MagicMock(reltype=RT.SLIDE_LAYOUT, is_external=False)
        media_rel = MagicMock(reltype=RT.MEDIA, is_external=False, target_part=src_media_part)
        video_rel = MagicMock(reltype=RT.VIDEO, is_external=False, target_part=src_media_part)
        src_part.rels._rels.update(  # type: ignore[attr-defined]
            {"rId1": layout_rel, "rId9": media_rel, "rId10": video_rel}
        )

        target_package_ = instance_mock(request, Package)
        target_package_.next_media_partname.return_value = PackURI("/ppt/media/media2.mp4")
        slide_layout_part_ = instance_mock(request, SlideLayoutPart)

        relate_to_ = method_mock(request, SlidePart, "relate_to", autospec=True)
        relate_to_.side_effect = ["rIdL", "rIdM", "rIdV"]

        SlidePart.clone_from(
            src_part,
            PackURI("/ppt/slides/slide2.xml"),
            target_package_,
            slide_layout_part_,
        )

        # -- next_media_partname is called exactly once (one new MediaPart
        # -- materialised for both MEDIA and VIDEO rels). --
        assert target_package_.next_media_partname.call_count == 1
        # -- Two rels on the new part, both pointing at the same cloned MediaPart. --
        media_calls = [
            c for c in relate_to_.mock_calls
            if len(c.args) >= 3 and c.args[2] in (RT.MEDIA, RT.VIDEO)
        ]
        assert len(media_calls) == 2
        assert media_calls[0].args[1] is media_calls[1].args[1]

    def it_shallow_clones_unknown_internal_rels_into_the_target_package(self, request):
        """OLE / embedded-package rels get a fresh part with the same blob in the target."""
        from unittest.mock import MagicMock

        src_sld = element("p:sld/p:cSld/p:spTree")
        src_part = SlidePart(
            PackURI("/ppt/slides/slide1.xml"), CT.PML_SLIDE, None, src_sld
        )
        src_pkg = object()
        target_package_ = instance_mock(request, Package)
        target_package_.next_partname.return_value = PackURI(
            "/ppt/embeddings/oleObject2.bin"
        )
        # -- Source OLE part (in *source* package, distinct from target). --
        # -- Must be a real Part subclass so `type(src_ole_part)(...)` constructs
        # -- a valid new instance in the target package. --
        src_ole_part = EmbeddedPackagePart(
            PackURI("/ppt/embeddings/oleObject1.bin"),
            "application/x-ole",
            src_pkg,  # type: ignore[arg-type]
            b"OLEDATA",
        )

        layout_rel = MagicMock(reltype=RT.SLIDE_LAYOUT, is_external=False)
        ole_rel = MagicMock(
            reltype=RT.OLE_OBJECT, is_external=False, target_part=src_ole_part
        )
        src_part.rels._rels.update(  # type: ignore[attr-defined]
            {"rId1": layout_rel, "rId7": ole_rel}
        )

        slide_layout_part_ = instance_mock(request, SlideLayoutPart)
        relate_to_ = method_mock(request, SlidePart, "relate_to", autospec=True)
        relate_to_.side_effect = ["rIdL", "rIdO"]

        SlidePart.clone_from(
            src_part,
            PackURI("/ppt/slides/slide2.xml"),
            target_package_,
            slide_layout_part_,
        )

        # -- A fresh partname was allocated in the target package. --
        target_package_.next_partname.assert_called_once_with(
            "/ppt/embeddings/oleObject%d.bin"
        )
        # -- The OLE_OBJECT rel on the new part points at a newly-materialised
        # -- EmbeddedPackagePart in the target package (same blob, new partname). --
        ole_calls = [
            c for c in relate_to_.mock_calls
            if len(c.args) >= 3 and c.args[2] == RT.OLE_OBJECT
        ]
        assert len(ole_calls) == 1
        new_ole_part = ole_calls[0].args[1]
        assert new_ole_part is not src_ole_part
        assert isinstance(new_ole_part, EmbeddedPackagePart)
        assert new_ole_part.package is target_package_
        assert new_ole_part.blob == b"OLEDATA"

    def it_provides_access_to_its_slide(self, slide_fixture):
        slide_part, Slide_, sld, slide_ = slide_fixture
        slide = slide_part.slide
        Slide_.assert_called_once_with(sld, slide_part)
        assert slide is slide_

    def it_provides_access_to_the_slide_layout(self, layout_fixture):
        slide_part, slide_layout_ = layout_fixture
        slide_layout = slide_part.slide_layout
        slide_part.part_related_by.assert_called_once_with(slide_part, RT.SLIDE_LAYOUT)
        assert slide_layout is slide_layout_

    def it_knows_the_minimal_element_xml_for_a_slide(self):
        path = absjoin(test_file_dir, "minimal_slide.xml")
        sld = CT_Slide.new()
        with open(path, "r") as f:
            expected_xml = f.read()
        assert sld.xml == expected_xml

    def it_gets_its_notes_slide_to_help(self, notes_slide_fixture):
        slide_part, calls, notes_slide_ = notes_slide_fixture

        notes_slide = slide_part.notes_slide

        slide_part.part_related_by.assert_called_once_with(slide_part, RT.NOTES_SLIDE)
        assert slide_part._add_notes_slide_part.call_args_list == calls
        assert notes_slide is notes_slide_

    def it_adds_a_notes_slide_part_to_help(
        self, package_, NotesSlidePart_, notes_slide_part_, relate_to_
    ):
        NotesSlidePart_.new.return_value = notes_slide_part_
        slide_part = SlidePart(None, None, package_, None)

        notes_slide_part = slide_part._add_notes_slide_part()

        assert notes_slide_part is notes_slide_part_
        NotesSlidePart_.new.assert_called_once_with(package_, slide_part)
        relate_to_.assert_called_once_with(slide_part, notes_slide_part, RT.NOTES_SLIDE)

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[True, False])
    def has_notes_slide_fixture(self, request, part_related_by_):
        has_notes_slide = request.param
        slide_part = SlidePart(None, None, None, None)
        part_related_by_.side_effect = None if has_notes_slide else KeyError
        expected_value = has_notes_slide
        return slide_part, expected_value

    @pytest.fixture
    def layout_fixture(self, slide_layout_, part_related_by_):
        slide_part = SlidePart(None, None, None, None)
        part_related_by_.return_value.slide_layout = slide_layout_
        return slide_part, slide_layout_

    @pytest.fixture(params=[True, False])
    def notes_slide_fixture(
        self,
        request,
        notes_slide_,
        part_related_by_,
        _add_notes_slide_part_,
        notes_slide_part_,
    ):
        has_notes_slide = request.param
        slide_part = SlidePart(None, None, None, None)
        part_related_by_.return_value = notes_slide_part_
        add_calls = []
        if not has_notes_slide:
            part_related_by_.side_effect = KeyError
            add_calls.append(call(slide_part))
        notes_slide_part_.notes_slide = notes_slide_
        return slide_part, add_calls, notes_slide_

    @pytest.fixture
    def slide_fixture(self, Slide_, slide_):
        sld = element("p:sld")
        slide_part = SlidePart(None, None, None, sld)
        return slide_part, Slide_, sld, slide_

    @pytest.fixture
    def slide_id_fixture(self, package_, presentation_part_):
        slide_part = SlidePart(None, None, package_, None)
        slide_id = 256
        package_.presentation_part = presentation_part_
        presentation_part_.slide_id.return_value = slide_id
        return slide_part, presentation_part_, slide_id

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _add_notes_slide_part_(self, request, notes_slide_part_):
        return method_mock(
            request,
            SlidePart,
            "_add_notes_slide_part",
            return_value=notes_slide_part_,
            autospec=True,
        )

    @pytest.fixture
    def media_part_(self, request):
        return instance_mock(request, MediaPart)

    @pytest.fixture
    def notes_slide_(self, request):
        return instance_mock(request, NotesSlide)

    @pytest.fixture
    def NotesSlidePart_(self, request, notes_slide_part_):
        return class_mock(
            request, "pptx.parts.slide.NotesSlidePart", return_value=notes_slide_part_
        )

    @pytest.fixture
    def notes_slide_part_(self, request):
        return instance_mock(request, NotesSlidePart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_related_by_(self, request):
        return method_mock(request, SlidePart, "part_related_by", autospec=True)

    @pytest.fixture
    def presentation_part_(self, request):
        return instance_mock(request, PresentationPart)

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(request, SlidePart, "relate_to", autospec=True)

    @pytest.fixture
    def Slide_(self, request, slide_):
        return class_mock(request, "pptx.parts.slide.Slide", return_value=slide_)

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def video_(self, request):
        return instance_mock(request, Video)


class DescribeSlideLayoutPart(object):
    """Unit-test suite for `pptx.parts.slide.SlideLayoutPart` objects."""

    def it_provides_access_to_its_slide_master(self, request):
        slide_master_ = instance_mock(request, SlideMaster)
        slide_master_part_ = instance_mock(request, SlideMasterPart, slide_master=slide_master_)
        part_related_by_ = method_mock(
            request, SlideLayoutPart, "part_related_by", return_value=slide_master_part_
        )
        slide_layout_part = SlideLayoutPart(None, None, None, None)

        slide_master = slide_layout_part.slide_master

        part_related_by_.assert_called_once_with(slide_layout_part, RT.SLIDE_MASTER)
        assert slide_master is slide_master_

    def it_provides_access_to_its_slide_layout(self, request):
        slide_layout_ = instance_mock(request, SlideLayout)
        SlideLayout_ = class_mock(
            request, "pptx.parts.slide.SlideLayout", return_value=slide_layout_
        )
        sldLayout = element("p:sldLayout")
        slide_layout_part = SlideLayoutPart(None, None, None, sldLayout)

        slide_layout = slide_layout_part.slide_layout

        SlideLayout_.assert_called_once_with(sldLayout, slide_layout_part)
        assert slide_layout is slide_layout_

    def it_can_be_cloned_from_another_layout(self, request):
        """SlideLayoutPart.clone_from deep-copies the layout and re-relates it.

        Verifies the layout's image rel is cloned into the target package,
        the SLIDE_MASTER rel is dropped/rewritten to the target master, and
        the `r:embed` attribute on the cloned XML is remapped.
        """
        from unittest.mock import MagicMock

        # -- Source SlideLayoutPart with: master rel + image rel + hyperlink rel. --
        src_sldLayout = element(
            "p:sldLayout/p:cSld/p:spTree/p:pic/p:blipFill/a:blip{r:embed=rId9}"
        )
        src_part = SlideLayoutPart(
            PackURI("/ppt/slideLayouts/slideLayout1.xml"),
            CT.PML_SLIDE_LAYOUT,
            None,
            src_sldLayout,
        )
        master_rel = MagicMock(reltype=RT.SLIDE_MASTER, is_external=False)
        image_rel = MagicMock(reltype=RT.IMAGE, is_external=False)
        image_rel.target_part = instance_mock(request, ImagePart, blob=b"PNGBYTES")
        hl_rel = MagicMock(
            reltype=RT.HYPERLINK, is_external=True, target_ref="https://example.com/"
        )
        src_part.rels._rels.update(  # type: ignore[attr-defined]
            {"rId1": master_rel, "rId9": image_rel, "rId7": hl_rel}
        )

        # -- Target package: stub clone_image_part to return a fake image part. --
        target_package_ = instance_mock(request, Package)
        target_image_part_ = instance_mock(request, ImagePart)
        target_package_.get_or_add_image_part.return_value = target_image_part_
        slide_master_part_ = instance_mock(request, SlideMasterPart)
        # -- Stub SlideLayoutPart.relate_to so no real relationship graph is touched.
        # -- Calls: (1) master rel -> "rIdM", (2) image rel -> "rIdI",
        # -- (3) hyperlink rel -> "rIdH".
        relate_to_ = method_mock(request, SlideLayoutPart, "relate_to", autospec=True)
        relate_to_.side_effect = ["rIdM", "rIdI", "rIdH"]

        partname = PackURI("/ppt/slideLayouts/slideLayout12.xml")
        new_part = SlideLayoutPart.clone_from(
            src_part, partname, target_package_, slide_master_part_
        )

        assert isinstance(new_part, SlideLayoutPart)
        assert new_part.partname == partname
        # -- shape tree is deep-copied, not aliased --
        assert new_part._element is not src_sldLayout
        # -- `r:embed` on the cloned blip was remapped from rId9 to the new image rId --
        embed_attr = new_part._element.xpath(".//a:blip/@r:embed")[0]
        assert embed_attr == "rIdI"
        # -- master rel went through the destination master, not the source --
        assert relate_to_.call_args_list[0] == call(
            new_part, slide_master_part_, RT.SLIDE_MASTER
        )
        # -- target package was asked for an image part --
        target_package_.get_or_add_image_part.assert_called_once()


class DescribeSlideMasterPart(object):
    """Unit-test suite for `pptx.parts.slide.SlideMasterPart` objects."""

    def it_provides_access_to_its_slide_master(self, request):
        slide_master_ = instance_mock(request, SlideMaster)
        SlideMaster_ = class_mock(
            request, "pptx.parts.slide.SlideMaster", return_value=slide_master_
        )
        sldMaster = element("p:sldMaster")
        slide_master_part = SlideMasterPart(None, None, None, sldMaster)

        slide_master = slide_master_part.slide_master

        SlideMaster_.assert_called_once_with(sldMaster, slide_master_part)
        assert slide_master is slide_master_

    def it_provides_access_to_a_related_slide_layout(self, request):
        slide_layout_ = instance_mock(request, SlideLayout)
        slide_layout_part_ = instance_mock(request, SlideLayoutPart, slide_layout=slide_layout_)
        related_part_ = method_mock(
            request, SlideMasterPart, "related_part", return_value=slide_layout_part_
        )
        slide_master_part = SlideMasterPart(None, None, None, None)

        slide_layout = slide_master_part.related_slide_layout("rId42")

        related_part_.assert_called_once_with(slide_master_part, "rId42")
        assert slide_layout is slide_layout_

    def it_can_add_a_layout_cloned_from_another_master(self, request):
        """add_layout_from() clones the layout part and establishes SLIDE_LAYOUT rel."""
        package_ = instance_mock(request, Package)
        package_.next_partname.return_value = PackURI(
            "/ppt/slideLayouts/slideLayout42.xml"
        )
        source_layout_part_ = instance_mock(request, SlideLayoutPart)
        new_layout_ = instance_mock(request, SlideLayout)
        new_layout_part_ = instance_mock(
            request, SlideLayoutPart, slide_layout=new_layout_
        )
        source_layout_ = instance_mock(request, SlideLayout, part=source_layout_part_)
        clone_from_ = method_mock(
            request,
            SlideLayoutPart,
            "clone_from",
            return_value=new_layout_part_,
            autospec=False,
        )
        relate_to_ = method_mock(
            request, SlideMasterPart, "relate_to", return_value="rIdX", autospec=True
        )
        slide_master_part = SlideMasterPart(
            PackURI("/ppt/slideMasters/slideMaster1.xml"),
            CT.PML_SLIDE_MASTER,
            package_,
            None,
        )

        rId, new_layout = slide_master_part.add_layout_from(source_layout_)

        package_.next_partname.assert_called_once_with(
            "/ppt/slideLayouts/slideLayout%d.xml"
        )
        clone_from_.assert_called_once_with(
            source_layout_part_,
            PackURI("/ppt/slideLayouts/slideLayout42.xml"),
            package_,
            slide_master_part,
        )
        relate_to_.assert_called_once_with(
            slide_master_part, new_layout_part_, RT.SLIDE_LAYOUT
        )
        assert rId == "rIdX"
        assert new_layout is new_layout_
