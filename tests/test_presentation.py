# encoding: utf-8

"""Test suite for pptx.presentation module."""

from __future__ import absolute_import, print_function

import pytest

from pptx.opc.packuri import PackURI
from pptx.opc.package import RelationshipCollection
from pptx.oxml.presentation import CT_Presentation, CT_SlideIdList
from pptx.parts.coreprops import CoreProperties
from pptx.parts.part import PartCollection
from pptx.parts.slides import SlideCollection
from pptx.presentation import Package, Presentation

from .unitutil import absjoin, class_mock, instance_mock, test_file_dir


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')


class DescribePackage(object):

    def it_loads_default_template_when_opened_with_no_path(self):
        prs = Package.open().presentation
        assert prs is not None
        slidemasters = prs.slidemasters
        assert slidemasters is not None
        assert len(slidemasters) == 1
        slidelayouts = slidemasters[0].slidelayouts
        assert slidelayouts is not None
        assert len(slidelayouts) == 11

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
        slidelayouts = slidemasters[0].slidelayouts
        assert slidelayouts is not None
        assert len(slidelayouts) == 11

    # fixtures ---------------------------------------------

    @pytest.fixture
    def temp_pptx_path(self, tmpdir):
        return absjoin(str(tmpdir), 'test-pptx.pptx')


class DescribePresentation(object):

    def it_provides_access_to_the_slide_masters(self, prs):
        assert isinstance(prs.slidemasters, PartCollection)

    def it_creates_slide_collection_on_first_reference(
            self, prs, SlideCollection_, sldIdLst_, rels_, slides_):
        slides = prs.slides
        # verify -----------------------
        prs._element.get_or_add_sldIdLst.assert_called_once_with()
        SlideCollection_.assert_called_once_with(sldIdLst_, rels_, prs)
        assert slides == slides_

    def it_reuses_slide_collection_instance_on_later_references(self, prs):
        slides_1 = prs.slides
        slides_2 = prs.slides
        assert slides_2 is slides_1

    # fixtures ---------------------------------------------

    @pytest.fixture
    def ct_presentation_(self, request, sldIdLst_):
        ct_presentation_ = instance_mock(request, CT_Presentation)
        ct_presentation_.get_or_add_sldIdLst.return_value = sldIdLst_
        return ct_presentation_

    @pytest.fixture
    def prs(self, ct_presentation_, rels_):
        partname = PackURI('/ppt/presentation.xml')
        prs = Presentation(partname, None, ct_presentation_, None)
        setattr(prs, '__rels', rels_)
        return prs

    @pytest.fixture
    def rels_(self, request):
        return instance_mock(request, RelationshipCollection)

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
