# encoding: utf-8

"""
Test suite for pptx.parts.presentation module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.opc.packuri import PackURI
from pptx.oxml.presentation import (
    CT_Presentation, CT_SlideIdList, CT_SlideMasterIdList
)
from pptx.parts.presentation import PresentationPart, _SlideMasters
from pptx.parts.slidemaster import SlideMaster
from pptx.parts.slide import SlideCollection

from ..oxml.unitdata.presentation import (
    a_presentation, a_sldMasterId, a_sldMasterIdLst, a_sldSz
)
from ..unitutil import class_mock, instance_mock, method_mock, property_mock


class DescribePresentationPart(object):

    def it_knows_the_width_of_its_slides(self, slide_width_fixture):
        prs_part, slide_width = slide_width_fixture
        assert prs_part.slide_width == slide_width

    def it_provides_access_to_its_slide_masters(self, masters_fixture):
        presentation_part = masters_fixture
        slide_masters = presentation_part.slide_masters
        assert isinstance(slide_masters, _SlideMasters)

    def it_provides_access_to_its_sldMasterIdLst(self, presentation_part):
        sldMasterIdLst = presentation_part.sldMasterIdLst
        assert isinstance(sldMasterIdLst, CT_SlideMasterIdList)

    def it_creates_slide_collection_on_first_reference(
            self, prs, SlideCollection_, sldIdLst_, slides_):
        slides = prs.slides
        # verify -----------------------
        prs._element.get_or_add_sldIdLst.assert_called_once_with()
        SlideCollection_.assert_called_once_with(sldIdLst_, prs)
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
    def slide_width_fixture(self):
        cx = 8765432
        presentation_elm = (
            a_presentation().with_nsdecls().with_child(
                a_sldSz().with_cx(cx))
        ).element
        prs_part = PresentationPart(None, None, presentation_elm, None)
        slide_width = cx
        return prs_part, slide_width

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
    def SlideCollection_(self, request, slides_):
        SlideCollection_ = class_mock(
            request, 'pptx.parts.presentation.SlideCollection'
        )
        SlideCollection_.return_value = slides_
        return SlideCollection_

    @pytest.fixture
    def slides_(self, request):
        return instance_mock(request, SlideCollection)


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
