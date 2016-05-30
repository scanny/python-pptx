# encoding: utf-8

"""
Test suite for pptx.parts.slidelayout module
"""

from __future__ import absolute_import

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.slidelayout import SlideLayoutPart
from pptx.parts.slidemaster import SlideMasterPart
from pptx.slide import SlideLayout

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, instance_mock, method_mock


class DescribeSlideLayoutPart(object):

    def it_provides_access_to_its_slide_master(self, master_fixture):
        slide_layout_part, part_related_by_, slide_master_ = master_fixture
        slide_master = slide_layout_part.slide_master
        part_related_by_.assert_called_once_with(RT.SLIDE_MASTER)
        assert slide_master is slide_master_

    def it_provides_access_to_its_slide_layout(self, layout_fixture):
        slide_layout_part, SlideLayout_ = layout_fixture[:2]
        sldLayout, slide_layout_ = layout_fixture[2:]
        slide_layout = slide_layout_part.slide_layout
        SlideLayout_.assert_called_once_with(sldLayout, slide_layout_part)
        assert slide_layout is slide_layout_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def layout_fixture(self, SlideLayout_, slide_layout_):
        sldLayout = element('w:sldLayout')
        slide_layout_part = SlideLayoutPart(None, None, sldLayout)
        return slide_layout_part, SlideLayout_, sldLayout, slide_layout_

    @pytest.fixture
    def master_fixture(self, slide_master_, part_related_by_):
        slide_layout_part = SlideLayoutPart(None, None, None, None)
        return slide_layout_part, part_related_by_, slide_master_

    # fixture components -----------------------------------

    @pytest.fixture
    def part_related_by_(self, request, slide_master_):
        return method_mock(
            request, SlideLayoutPart, 'part_related_by',
            return_value=slide_master_
        )

    @pytest.fixture
    def SlideLayout_(self, request, slide_layout_):
        return class_mock(
            request, 'pptx.parts.slidelayout.SlideLayout',
            return_value=slide_layout_
        )

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMasterPart)
