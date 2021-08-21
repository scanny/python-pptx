# encoding: utf-8

"""Unit-test suite for `pptx.parts.presentation` module."""

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.package import Package
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import NotesMasterPart, SlideMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.slide import NotesMaster, Slide, SlideLayout, SlideMaster

from ..unitutil.cxml import element
from ..unitutil.mock import call, class_mock, instance_mock, method_mock, property_mock


class DescribePresentationPart(object):
    """Unit-test suite for `pptx.parts.presentation.PresentationPart` objects."""

    def it_provides_access_to_its_presentation(self, prs_fixture):
        prs_part, Presentation_, prs_elm, prs_ = prs_fixture
        prs = prs_part.presentation
        Presentation_.assert_called_once_with(prs_elm, prs_part)
        assert prs is prs_

    def it_provides_access_to_its_core_properties(self, core_props_fixture):
        prs_part, core_properties_ = core_props_fixture
        core_properties = prs_part.core_properties
        assert core_properties is core_properties_

    def it_provides_access_to_the_notes_master_part(self, nmp_get_fixture):
        """
        This is the first of a two-part test to cover the existing notes
        master case. The notes master not-present case follows.
        """
        prs_part, notes_master_part_ = nmp_get_fixture
        notes_master_part = prs_part.notes_master_part
        prs_part.part_related_by.assert_called_once_with(prs_part, RT.NOTES_MASTER)
        assert notes_master_part is notes_master_part_

    def it_adds_a_notes_master_part_when_needed(self, nmp_add_fixture):
        """
        This is the second of a two-part test to cover the
        notes-master-not-present case. The notes master present case is just
        above.
        """
        prs_part, NotesMasterPart_ = nmp_add_fixture[:2]
        package_, notes_master_part_ = nmp_add_fixture[2:]

        notes_master_part = prs_part.notes_master_part

        NotesMasterPart_.create_default.assert_called_once_with(package_)
        prs_part.relate_to.assert_called_once_with(
            prs_part, notes_master_part_, RT.NOTES_MASTER
        )
        assert notes_master_part is notes_master_part_

    def it_provides_access_to_its_notes_master(self, notes_master_fixture):
        prs_part, notes_master_ = notes_master_fixture
        notes_master = prs_part.notes_master
        assert notes_master is notes_master_

    def it_provides_access_to_a_related_slide(self, request, slide_, related_part_):
        slide_part_ = instance_mock(request, SlidePart, slide=slide_)
        related_part_.return_value = slide_part_
        prs_part = PresentationPart(None, None, None, None)

        slide = prs_part.related_slide("rId42")

        related_part_.assert_called_once_with(prs_part, "rId42")
        assert slide is slide_

    def it_provides_access_to_a_related_master(
        self, request, slide_master_, related_part_
    ):
        slide_master_part_ = instance_mock(
            request, SlideMasterPart, slide_master=slide_master_
        )
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

    def it_can_save_the_package_to_a_file(self, save_fixture):
        prs_part, file_, package_ = save_fixture
        prs_part.save(file_)
        package_.save.assert_called_once_with(file_)

    def it_can_add_a_new_slide(self, add_slide_fixture):
        prs_part, slide_layout_, SlidePart_, partname = add_slide_fixture[:4]
        package_, slide_layout_part_, slide_part_ = add_slide_fixture[4:7]
        rId_, slide_ = add_slide_fixture[7:]

        rId, slide = prs_part.add_slide(slide_layout_)

        SlidePart_.new.assert_called_once_with(partname, package_, slide_layout_part_)
        prs_part.relate_to.assert_called_once_with(prs_part, slide_part_, RT.SLIDE)
        assert rId is rId_
        assert slide is slide_

    def it_finds_the_slide_id_of_a_slide_part(self, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        related_part_.side_effect = iter((None, slide_part_, None))
        prs_part = PresentationPart(None, None, prs_elm, None)

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
        prs_part = PresentationPart(None, None, prs_elm, None)

        with pytest.raises(ValueError):
            prs_part.slide_id(slide_part_)

    @pytest.mark.parametrize("is_present", (True, False))
    def it_finds_a_slide_by_slide_id(
        self, is_present, slide_, slide_part_, related_part_
    ):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        slide_id = 257 if is_present else 666
        expected_value = slide_ if is_present else None
        related_part_.return_value = slide_part_
        slide_part_.slide = slide_
        prs_part = PresentationPart(None, None, prs_elm, None)

        slide = prs_part.get_slide(slide_id)

        assert slide == expected_value

    def it_knows_the_next_slide_partname_to_help(self, next_fixture):
        prs_part, partname = next_fixture
        assert prs_part._next_slide_partname == partname

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_slide_fixture(
        self,
        package_,
        slide_layout_,
        SlidePart_,
        slide_part_,
        slide_,
        _next_slide_partname_prop_,
        relate_to_,
    ):
        prs_part = PresentationPart(None, None, None, package_)
        partname = _next_slide_partname_prop_.return_value
        rId_ = "rId42"
        SlidePart_.new.return_value = slide_part_
        relate_to_.return_value = rId_
        slide_layout_part_ = slide_layout_.part
        slide_part_.slide = slide_
        return (
            prs_part,
            slide_layout_,
            SlidePart_,
            partname,
            package_,
            slide_layout_part_,
            slide_part_,
            rId_,
            slide_,
        )

    @pytest.fixture
    def core_props_fixture(self, package_, core_properties_):
        prs_part = PresentationPart(None, None, None, package_)
        package_.core_properties = core_properties_
        return prs_part, core_properties_

    @pytest.fixture
    def next_fixture(self):
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId,p:sldId)")
        prs_part = PresentationPart(None, None, prs_elm)
        partname = PackURI("/ppt/slides/slide3.xml")
        return prs_part, partname

    @pytest.fixture
    def notes_master_fixture(
        self, notes_master_part_prop_, notes_master_part_, notes_master_
    ):
        prs_part = PresentationPart(None, None, None, None)
        notes_master_part_prop_.return_value = notes_master_part_
        notes_master_part_.notes_master = notes_master_
        return prs_part, notes_master_

    @pytest.fixture
    def nmp_add_fixture(
        self,
        package_,
        NotesMasterPart_,
        notes_master_part_,
        part_related_by_,
        relate_to_,
    ):
        prs_part = PresentationPart(None, None, None, package_)
        part_related_by_.side_effect = KeyError
        NotesMasterPart_.create_default.return_value = notes_master_part_
        return prs_part, NotesMasterPart_, package_, notes_master_part_

    @pytest.fixture
    def nmp_get_fixture(self, notes_master_part_, part_related_by_):
        prs_part = PresentationPart(None, None, None, None)
        part_related_by_.return_value = notes_master_part_
        return prs_part, notes_master_part_

    @pytest.fixture
    def prs_fixture(self, Presentation_, prs_):
        prs_elm = element("p:presentation")
        prs_part = PresentationPart(None, None, prs_elm)
        return prs_part, Presentation_, prs_elm, prs_

    @pytest.fixture
    def save_fixture(self, package_):
        prs_part = PresentationPart(None, None, None, package_)
        file_ = "foobar.docx"
        return prs_part, file_, package_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def core_properties_(self, request):
        return instance_mock(request, CorePropertiesPart)

    @pytest.fixture
    def _next_slide_partname_prop_(self, request):
        return property_mock(request, PresentationPart, "_next_slide_partname")

    @pytest.fixture
    def NotesMasterPart_(self, request, prs_):
        return class_mock(request, "pptx.parts.presentation.NotesMasterPart")

    @pytest.fixture
    def notes_master_(self, request):
        return instance_mock(request, NotesMaster)

    @pytest.fixture
    def notes_master_part_(self, request):
        return instance_mock(request, NotesMasterPart)

    @pytest.fixture
    def notes_master_part_prop_(self, request, notes_master_part_):
        return property_mock(request, PresentationPart, "notes_master_part")

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_related_by_(self, request):
        return method_mock(request, PresentationPart, "part_related_by", autospec=True)

    @pytest.fixture
    def Presentation_(self, request, prs_):
        return class_mock(
            request, "pptx.parts.presentation.Presentation", return_value=prs_
        )

    @pytest.fixture
    def prs_(self, request):
        return instance_mock(request, Presentation)

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(request, PresentationPart, "relate_to", autospec=True)

    @pytest.fixture
    def related_part_(self, request):
        return method_mock(request, PresentationPart, "related_part")

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)

    @pytest.fixture
    def SlidePart_(self, request):
        return class_mock(request, "pptx.parts.presentation.SlidePart")
