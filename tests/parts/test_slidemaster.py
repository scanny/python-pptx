# encoding: utf-8

"""
Test suite for pptx.parts.slidemaster module
"""

from __future__ import absolute_import

import pytest

from pptx.opc.packuri import PackURI
from pptx.parts.slidemaster import SlideMaster
from pptx.presentation import Package

from ..unitutil import absjoin, test_file_dir


test_pptx_path = absjoin(test_file_dir, 'test.pptx')


class DescribeSlideMaster(object):

    def test_slidelayouts_property_empty_on_construction(self, slidemaster):
        assert len(slidemaster.slidelayouts) == 0

    def test_slidelayouts_correct_length_after_open(self):
        slidemaster = Package.open(test_pptx_path).presentation.slidemasters[0]
        slidelayouts = slidemaster.slidelayouts
        assert len(slidelayouts) == 11

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def slidemaster(self):
        partname = PackURI('/ppt/slideMasters/slideMaster1.xml')
        return SlideMaster(partname, None, None, None)
