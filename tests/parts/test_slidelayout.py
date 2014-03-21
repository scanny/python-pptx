# encoding: utf-8

"""
Test suite for pptx.parts.slidelayout module
"""

from __future__ import absolute_import

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.slidelayout import SlideLayout
from pptx.parts.slidemaster import SlideMaster

from ..unitutil import instance_mock, method_mock


class DescribeSlideLayout(object):

    def it_knows_the_slide_master_it_inherits_from(self, master_fixture):
        slide_layout, slide_master_ = master_fixture
        slide_master = slide_layout.slide_master
        slide_layout.part_related_by.assert_called_once_with(RT.SLIDE_MASTER)
        assert slide_master is slide_master_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def master_fixture(self, slide_master_, part_related_by_):
        slide_layout = SlideLayout(None, None, None, None)
        return slide_layout, slide_master_

    # fixture components -----------------------------------

    @pytest.fixture
    def part_related_by_(self, request, slide_master_):
        return method_mock(
            request, SlideLayout, 'part_related_by',
            return_value=slide_master_
        )

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)
