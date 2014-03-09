# encoding: utf-8

"""
Test suite for pptx.parts.slidemaster module
"""

from __future__ import absolute_import

import pytest

from pptx.parts.slidemaster import _SlideLayouts, SlideMaster

from ..unitutil import absjoin, test_file_dir


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
