# encoding: utf-8

"""Test suite for pptx.presentation module."""

from __future__ import absolute_import, print_function

import pytest

from pptx.opc.packuri import PackURI
from pptx.oxml.presentation import (
    CT_Presentation, CT_SlideIdList, CT_SlideMasterIdList
)
from pptx.parts.coreprops import CoreProperties
from pptx.parts.part import PartCollection
from pptx.parts.slidemaster import SlideMaster
from pptx.parts.slides import SlideCollection
from pptx.presentation import Package, Presentation, _SlideMasters


from .oxml.unitdata.presentation import (
    a_presentation, a_sldMasterId, a_sldMasterIdLst
)
from .unitutil import (
    absjoin, class_mock, instance_mock, method_mock, property_mock,
    test_file_dir
)


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')


class DescribePackage(object):

    def it_loads_default_template_when_opened_with_no_path(self):
        prs = Package.open().presentation
        assert prs is not None
        slidemasters = prs.slidemasters
        assert slidemasters is not None
        assert len(slidemasters) == 1
        slide_layouts = slidemasters[0].slide_layouts
        assert slide_layouts is not None
        assert len(slide_layouts) == 11

    def it_gathers_package_image_parts_on_open(self):
        pkg = Package.open(images_pptx_path)
        assert len(pkg._images) == 7

    def it_provides_ref_to_package_presentation_part(self):
        pkg = Package.open()
        assert isinstance(pkg.presentation, Presentation)

    def it_provides_ref_to_package_core_properties_part(self):
        pkg = Package.open()
        assert isinstance(pkg.core_properties, CoreProperties)

    def it_can_save_itself_to_a_pptx_file(self, temp_pptx_path):
        """Package.save produces a .pptx with plausible contents"""
        # setup ------------------------
        pkg = Package.open()
        # exercise ---------------------
        pkg.save(temp_pptx_path)
        # verify -----------------------
        pkg = Package.open(temp_pptx_path)
        prs = pkg.presentation
        assert prs is not None
        slidemasters = prs.slidemasters
        assert slidemasters is not None
        assert len(slidemasters) == 1
        slide_layouts = slidemasters[0].slide_layouts
        assert slide_layouts is not None
        assert len(slide_layouts) == 11

    # fixtures ---------------------------------------------

    @pytest.fixture
    def temp_pptx_path(self, tmpdir):
        return absjoin(str(tmpdir), 'test-pptx.pptx')


class DescribePresentation(object):

    def it_provides_access_to_its_slide_masters(self, masters_fixture):
        presentation = masters_fixture
        slide_masters = presentation.slide_masters
        assert isinstance(slide_masters, _SlideMasters)

    def it_provides_access_to_its_sldMasterIdLst(self, presentation):
        sldMasterIdLst = presentation.sldMasterIdLst
        assert isinstance(sldMasterIdLst, CT_SlideMasterIdList)

    def it_provides_access_to_the_slide_masters(self, prs):
        assert isinstance(prs.slidemasters, PartCollection)

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
        presentation = Presentation(None, None, None, None)
        return presentation

    # fixture components -----------------------------------

    @pytest.fixture
    def ct_presentation_(self, request, sldIdLst_):
        ct_presentation_ = instance_mock(request, CT_Presentation)
        ct_presentation_.get_or_add_sldIdLst.return_value = sldIdLst_
        return ct_presentation_

    @pytest.fixture
    def presentation(self, presentation_elm):
        return Presentation(None, None, presentation_elm, None)

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
        prs = Presentation(partname, None, ct_presentation_, None)
        return prs

    @pytest.fixture
    def sldIdLst_(self, request):
        return instance_mock(request, CT_SlideIdList)

    @pytest.fixture
    def SlideCollection_(self, request, slides_):
        SlideCollection_ = class_mock(
            request, 'pptx.presentation.SlideCollection'
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
    def getitem_fixture(self, presentation, related_parts_, slide_master_):
        slide_masters = _SlideMasters(presentation)
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
    def iter_rIds_fixture(self, presentation):
        slide_masters = _SlideMasters(presentation)
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
    def presentation(self, presentation_elm):
        return Presentation(None, None, presentation_elm, None)

    @pytest.fixture
    def presentation_(self, request):
        return instance_mock(request, Presentation)

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
        return property_mock(request, Presentation, 'related_parts')

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_master_2_(self, request):
        return instance_mock(request, SlideMaster)
