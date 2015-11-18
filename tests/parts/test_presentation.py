# encoding: utf-8

"""
Test suite for pptx.parts.presentation module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import _Relationship
from pptx.opc.packuri import PackURI
from pptx.oxml.parts.presentation import (
    CT_Presentation, CT_SlideId, CT_SlideIdList, CT_SlideMasterIdList
)
from pptx.parts.presentation import PresentationPart, _SlideMasters, _Slides
from pptx.parts.slide import Slide
from pptx.parts.slidelayout import SlideLayout
from pptx.parts.slidemaster import SlideMaster

from ..oxml.unitdata.presentation import (
    a_presentation, a_sldMasterId, a_sldMasterIdLst, a_sldSz
)
from ..unitutil.mock import (
    ANY, call, class_mock, instance_mock, MagicMock, method_mock,
    property_mock
)


class DescribePresentationPart(object):

    def it_knows_the_width_of_its_slides(self, slide_width_get_fixture):
        prs_part, slide_width = slide_width_get_fixture
        assert prs_part.slide_width == slide_width

    def it_can_change_the_width_of_its_slides(self, slide_width_set_fixture):
        prs_part, slide_width, expected_xml = slide_width_set_fixture
        prs_part.slide_width = slide_width
        assert prs_part._element.xml == expected_xml

    def it_knows_the_height_of_its_slides(self, slide_height_get_fixture):
        prs_part, slide_height = slide_height_get_fixture
        assert prs_part.slide_height == slide_height

    def it_can_change_the_height_of_its_slides(self, slide_height_set_fixture):
        prs_part, slide_height, expected_xml = slide_height_set_fixture
        prs_part.slide_height = slide_height
        assert prs_part._element.xml == expected_xml

    def it_provides_access_to_its_slide_masters(self, masters_fixture):
        presentation_part = masters_fixture
        slide_masters = presentation_part.slide_masters
        assert isinstance(slide_masters, _SlideMasters)

    def it_provides_access_to_its_sldMasterIdLst(self, presentation_part):
        sldMasterIdLst = presentation_part.sldMasterIdLst
        assert isinstance(sldMasterIdLst, CT_SlideMasterIdList)

    def it_creates_slide_collection_on_first_reference(
            self, prs, _Slides_, sldIdLst_, slides_):
        slides = prs.slides
        # verify -----------------------
        prs._element.get_or_add_sldIdLst.assert_called_once_with()
        _Slides_.assert_called_once_with(sldIdLst_, prs)
        slides.rename_slides.assert_called_once_with()
        assert slides == slides_

    def it_reuses_slide_collection_instance_on_later_references(self, prs):
        slides_1 = prs.slides
        slides_2 = prs.slides
        assert slides_2 is slides_1

    # fixtures ---------------------------------------------

    @pytest.fixture
    def masters_fixture(self):
        presentation_part = PresentationPart(None, None, None, None)
        return presentation_part

    @pytest.fixture
    def slide_height_get_fixture(self):
        cy = 5432109
        presentation_elm = (
            a_presentation().with_nsdecls().with_child(
                a_sldSz().with_cy(cy))
        ).element
        prs_part = PresentationPart(None, None, presentation_elm, None)
        slide_height = cy
        return prs_part, slide_height

    @pytest.fixture
    def slide_height_set_fixture(self):
        cy = 5432109
        presentation_elm = (
            a_presentation().with_nsdecls().with_child(
                a_sldSz().with_cy(0))
        ).element
        prs_part = PresentationPart(None, None, presentation_elm, None)
        expected_xml = (
            a_presentation().with_nsdecls().with_child(
                a_sldSz().with_cy(cy))
        ).xml()
        slide_height = cy
        return prs_part, slide_height, expected_xml

    @pytest.fixture
    def slide_width_get_fixture(self):
        cx = 8765432
        presentation_elm = (
            a_presentation().with_nsdecls().with_child(
                a_sldSz().with_cx(cx))
        ).element
        prs_part = PresentationPart(None, None, presentation_elm, None)
        slide_width = cx
        return prs_part, slide_width

    @pytest.fixture
    def slide_width_set_fixture(self):
        cx = 8765432
        presentation_elm = (
            a_presentation().with_nsdecls().with_child(
                a_sldSz().with_cx(0))
        ).element
        prs_part = PresentationPart(None, None, presentation_elm, None)
        expected_xml = (
            a_presentation().with_nsdecls().with_child(
                a_sldSz().with_cx(cx))
        ).xml()
        slide_width = cx
        return prs_part, slide_width, expected_xml

    # fixture components -----------------------------------

    @pytest.fixture
    def ct_presentation_(self, request, sldIdLst_):
        ct_presentation_ = instance_mock(request, CT_Presentation)
        ct_presentation_.get_or_add_sldIdLst.return_value = sldIdLst_
        return ct_presentation_

    @pytest.fixture
    def presentation_part(self, presentation_elm):
        return PresentationPart(None, None, presentation_elm, None)

    @pytest.fixture(params=[True, False])
    def presentation_elm(self, request):
        has_sldMasterIdLst = request.param
        presentation_bldr = a_presentation().with_nsdecls()
        if has_sldMasterIdLst:
            presentation_bldr.with_child(a_sldMasterIdLst())
        return presentation_bldr.element

    @pytest.fixture
    def prs(self, ct_presentation_):
        partname = PackURI('/ppt/presentation.xml')
        prs = PresentationPart(partname, None, ct_presentation_, None)
        return prs

    @pytest.fixture
    def sldIdLst_(self, request):
        return instance_mock(request, CT_SlideIdList)

    @pytest.fixture
    def _Slides_(self, request, slides_):
        _Slides_ = class_mock(
            request, 'pptx.parts.presentation._Slides'
        )
        _Slides_.return_value = slides_
        return _Slides_

    @pytest.fixture
    def slides_(self, request):
        return instance_mock(request, _Slides)


class DescribeSlideMasters(object):

    def it_knows_how_many_masters_it_contains(self, len_fixture):
        slide_masters, expected_count = len_fixture
        slide_master_count = len(slide_masters)
        assert slide_master_count == expected_count

    def it_can_iterate_over_the_slide_masters(self, iter_fixture):
        slide_masters, slide_master_, slide_master_2_ = iter_fixture
        assert [s for s in slide_masters] == [slide_master_, slide_master_2_]

    def it_iterates_over_rIds_to_help__iter__(self, iter_rIds_fixture):
        slide_masters, expected_rIds = iter_rIds_fixture
        assert [rId for rId in slide_masters._iter_rIds()] == expected_rIds

    def it_supports_indexed_access(self, getitem_fixture):
        slide_masters, idx, slide_master_ = getitem_fixture
        slide_master = slide_masters[idx]
        assert slide_master is slide_master_

    def it_raises_on_slide_master_index_out_of_range(self, getitem_fixture):
        slide_masters = getitem_fixture[0]
        with pytest.raises(IndexError):
            slide_masters[2]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getitem_fixture(
            self, presentation_part, related_parts_, slide_master_):
        slide_masters = _SlideMasters(presentation_part)
        related_parts_.return_value = {'rId1': None, 'rId2': slide_master_}
        idx = 1
        return slide_masters, idx, slide_master_

    @pytest.fixture
    def iter_fixture(
            self, presentation_, _iter_rIds_, slide_master_, slide_master_2_):
        presentation_.related_parts = {
            'rId1': slide_master_, 'rId2': slide_master_2_
        }
        slide_masters = _SlideMasters(presentation_)
        return slide_masters, slide_master_, slide_master_2_

    @pytest.fixture
    def iter_rIds_fixture(self, presentation_part):
        slide_masters = _SlideMasters(presentation_part)
        expected_rIds = ['rId1', 'rId2']
        return slide_masters, expected_rIds

    @pytest.fixture
    def len_fixture(self, presentation_):
        slide_masters = _SlideMasters(presentation_)
        presentation_.sldMasterIdLst = [1, 2]
        expected_count = 2
        return slide_masters, expected_count

    # fixture components -----------------------------------

    @pytest.fixture
    def _iter_rIds_(self, request):
        return method_mock(
            request, _SlideMasters, '_iter_rIds',
            return_value=iter(['rId1', 'rId2'])
        )

    @pytest.fixture
    def presentation_part(self, presentation_elm):
        return PresentationPart(None, None, presentation_elm, None)

    @pytest.fixture
    def presentation_(self, request):
        return instance_mock(request, PresentationPart)

    @pytest.fixture
    def presentation_elm(self, request):
        presentation_bldr = (
            a_presentation().with_nsdecls('p', 'r').with_child(
                a_sldMasterIdLst().with_child(
                    a_sldMasterId().with_rId('rId1')).with_child(
                    a_sldMasterId().with_rId('rId2'))
            )
        )
        return presentation_bldr.element

    @pytest.fixture
    def related_parts_(self, request):
        return property_mock(request, PresentationPart, 'related_parts')

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_master_2_(self, request):
        return instance_mock(request, SlideMaster)


class DescribeSlides(object):

    def it_supports_indexed_access(self, slides_with_slide_parts_, rIds_):
        slides, slide_, slide_2_ = slides_with_slide_parts_
        rId_, rId_2_ = rIds_
        # verify -----------------------
        assert slides[0] is slide_
        assert slides[1] is slide_2_
        slides._sldIdLst.__getitem__.assert_has_calls(
            [call(0), call(1)]
        )
        slides._prs.related_parts.__getitem__.assert_has_calls(
            [call(rId_), call(rId_2_)]
        )

    def it_raises_on_slide_index_out_of_range(self, slides):
        with pytest.raises(IndexError):
            slides[2]

    def it_can_find_the_index_of_a_slide(self, index_fixture):
        slides, slide, expected_value = index_fixture
        index = slides.index(slide)
        assert index == expected_value

    def it_raises_on_slide_not_in_collection(self, index_raise_fixture):
        slides, slide = index_raise_fixture
        with pytest.raises(ValueError):
            slides.index(slide)

    def it_can_iterate_over_the_slides(self, slides, slide_, slide_2_):
        assert [s for s in slides] == [slide_, slide_2_]

    def it_supports_len(self, slides):
        assert len(slides) == 2

    def it_can_add_a_new_slide(self, slides, slidelayout_, Slide_, slide_):
        slide = slides.add_slide(slidelayout_)
        Slide_.new.assert_called_once_with(
            slidelayout_, PackURI('/ppt/slides/slide3.xml'),
            slides._prs.package
        )
        slides._prs.relate_to.assert_called_once_with(slide_, RT.SLIDE)
        slides._sldIdLst.add_sldId.assert_called_once_with(ANY)
        assert slide is slide_

    def it_knows_the_next_available_slide_partname(
            self, slides_with_slide_parts_):
        slides = slides_with_slide_parts_[0]
        expected_partname = PackURI('/ppt/slides/slide3.xml')
        partname = slides._next_partname
        assert isinstance(partname, PackURI)
        assert partname == expected_partname

    def it_can_assign_partnames_to_the_slides(
            self, slides, slide_, slide_2_):
        slides.rename_slides()
        assert slide_.partname == '/ppt/slides/slide1.xml'
        assert slide_2_.partname == '/ppt/slides/slide2.xml'

    # fixtures -------------------------------------------------------
    #
    #   slides
    #   |
    #   +- ._sldIdLst = [sldId_, sldId_2_]
    #   |                |       |
    #   |                |       +- .rId = rId_2_
    #   |                |
    #   |                +- .rId = rId_
    #   +- ._prs
    #       |
    #       +- .related_parts = {rId_: slide_, rId_2_: slide_2_}
    #
    # ----------------------------------------------------------------

    @pytest.fixture
    def index_fixture(self, sldIdLst_, prs_, slide_2_):
        slides = _Slides(sldIdLst_, prs_)
        slide = slide_2_
        expected_value = 1
        return slides, slide, expected_value

    @pytest.fixture
    def index_raise_fixture(self, sldIdLst_, prs_):
        slides = _Slides(sldIdLst_, prs_)
        slide = "foobar"
        return slides, slide

    @pytest.fixture
    def slides_with_slide_parts_(self, sldIdLst_, prs_, slide_parts_):
        slide_, slide_2_ = slide_parts_
        slides = _Slides(sldIdLst_, prs_)
        return slides, slide_, slide_2_

    @pytest.fixture
    def slides(self, sldIdLst_, prs_):
        return _Slides(sldIdLst_, prs_)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def prs_(self, request, rel_, related_parts_):
        prs_ = instance_mock(request, PresentationPart)
        prs_.load_rel.return_value = rel_
        prs_.related_parts = related_parts_
        return prs_

    @pytest.fixture
    def rel_(self, request, rId_):
        return instance_mock(request, _Relationship, rId=rId_)

    @pytest.fixture
    def related_parts_(self, request, rIds_, slide_parts_):
        """
        Return pass-thru mock dict that both operates as a dict an records
        calls to __getitem__ for call asserts.
        """
        rId_, rId_2_ = rIds_
        slide_, slide_2_ = slide_parts_
        slide_rId_map = {rId_: slide_, rId_2_: slide_2_}

        def getitem(key):
            return slide_rId_map[key]

        related_parts_ = MagicMock()
        related_parts_.__getitem__.side_effect = getitem
        return related_parts_

    @pytest.fixture
    def rename_slides_(self, request):
        return method_mock(request, _Slides, 'rename_slides')

    @pytest.fixture
    def rId_(self, request):
        return 'rId1'

    @pytest.fixture
    def rId_2_(self, request):
        return 'rId2'

    @pytest.fixture
    def rIds_(self, request, rId_, rId_2_):
        return rId_, rId_2_

    @pytest.fixture
    def Slide_(self, request, slide_):
        Slide_ = class_mock(request, 'pptx.parts.presentation.Slide')
        Slide_.new.return_value = slide_
        return Slide_

    @pytest.fixture
    def sldId_(self, request, rId_):
        return instance_mock(request, CT_SlideId, rId=rId_)

    @pytest.fixture
    def sldId_2_(self, request, rId_2_):
        return instance_mock(request, CT_SlideId, rId=rId_2_)

    @pytest.fixture
    def sldIdLst_(self, request, sldId_, sldId_2_):
        sldIdLst_ = instance_mock(request, CT_SlideIdList)
        sldIdLst_.__getitem__.side_effect = [sldId_, sldId_2_]
        sldIdLst_.__iter__.return_value = iter([sldId_, sldId_2_])
        sldIdLst_.__len__.return_value = 2
        return sldIdLst_

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_2_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_parts_(self, request, slide_, slide_2_):
        return slide_, slide_2_

    @pytest.fixture
    def slidelayout_(self, request):
        return instance_mock(request, SlideLayout)
