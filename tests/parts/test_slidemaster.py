# encoding: utf-8

"""
Test suite for pptx.parts.slidemaster module
"""

from __future__ import absolute_import

import pytest

from pptx.parts.slidemaster import _SlideLayouts, SlideMaster
from pptx.oxml.slidemaster import CT_SlideLayoutIdList

from ..oxml.unitdata.slides import a_sldLayoutIdLst, a_sldMaster
from ..unitutil import absjoin, instance_mock, test_file_dir


test_pptx_path = absjoin(test_file_dir, 'test.pptx')


class DescribeSlideMaster(object):

    def it_provides_access_to_its_slide_layouts(self, layouts_fixture):
        slide_master = layouts_fixture
        slide_layouts = slide_master.slide_layouts
        assert isinstance(slide_layouts, _SlideLayouts)

    def it_provides_access_to_its_sldLayoutIdLst(self, slide_master):
        sldLayoutIdLst = slide_master.sldLayoutIdLst
        assert isinstance(sldLayoutIdLst, CT_SlideLayoutIdList)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def layouts_fixture(self):
        slide_master = SlideMaster(None, None, None, None)
        return slide_master

    @pytest.fixture(params=[True, False])
    def sldMaster(self, request):
        has_sldLayoutIdLst = request.param
        sldMaster_bldr = a_sldMaster().with_nsdecls()
        if has_sldLayoutIdLst:
            sldMaster_bldr.with_child(a_sldLayoutIdLst())
        return sldMaster_bldr.element

    @pytest.fixture
    def slide_master(self, sldMaster):
        return SlideMaster(None, None, sldMaster, None)


class DescribeSlideLayouts(object):

    def it_knows_how_many_layouts_it_contains(self, len_fixture):
        slide_layouts, expected_count = len_fixture
        slide_layout_count = len(slide_layouts)
        assert slide_layout_count == expected_count

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def len_fixture(self, slide_master_):
        slide_layouts = _SlideLayouts(slide_master_)
        slide_master_.sldLayoutIdLst = [1, 2]
        expected_count = 2
        return slide_layouts, expected_count

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)
