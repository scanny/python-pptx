# encoding: utf-8

"""
Test suite for pptx.parts.slidemaster module
"""

from __future__ import absolute_import

import pytest

from pptx.parts.slidemaster import _SlideLayouts, SlideMaster

from ..unitutil import absjoin, instance_mock, test_file_dir


test_pptx_path = absjoin(test_file_dir, 'test.pptx')


class DescribeSlideMaster(object):

    def it_provides_access_to_its_slide_layouts(self, layouts_fixture):
        slide_master = layouts_fixture
        slide_layouts = slide_master.slide_layouts
        assert isinstance(slide_layouts, _SlideLayouts)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def layouts_fixture(self):
        slide_master = SlideMaster(None, None, None, None)
        return slide_master


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
