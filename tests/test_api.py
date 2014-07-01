# encoding: utf-8

"""
Test suite for pptx.api module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.api import Presentation
from pptx.parts.presentation import PresentationPart

from .unitutil.mock import call, property_mock


class DescribePresentation(object):

    def it_knows_the_width_of_its_slides(self, slide_width_get_fixture):
        prs, slide_width = slide_width_get_fixture
        assert prs.slide_width == slide_width

    def it_can_change_the_width_of_its_slides(
            self, slide_width_set_fixture):
        prs, slide_width, part_slide_width_ = slide_width_set_fixture
        prs.slide_width = slide_width
        assert part_slide_width_.mock_calls == [call(slide_width)]

    def it_knows_the_height_of_its_slides(self, slide_height_get_fixture):
        prs, slide_height = slide_height_get_fixture
        assert prs.slide_height == slide_height

    def it_can_change_the_height_of_its_slides(
            self, slide_height_set_fixture):
        prs, slide_height, part_slide_height_ = slide_height_set_fixture
        prs.slide_height = slide_height
        assert part_slide_height_.mock_calls == [call(slide_height)]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def slide_height_get_fixture(self, part_slide_height_, slide_height):
        prs = Presentation()
        part_slide_height_.return_value = slide_height
        return prs, slide_height

    @pytest.fixture
    def slide_height_set_fixture(self, part_slide_height_, slide_height):
        prs = Presentation()
        return prs, slide_height, part_slide_height_

    @pytest.fixture
    def slide_width_get_fixture(self, part_slide_width_, slide_width):
        prs = Presentation()
        part_slide_width_.return_value = slide_width
        return prs, slide_width

    @pytest.fixture
    def slide_width_set_fixture(self, part_slide_width_, slide_width):
        prs = Presentation()
        return prs, slide_width, part_slide_width_

    # fixtures components --------------------------------------------

    @pytest.fixture
    def part_slide_height_(self, request):
        return property_mock(request, PresentationPart, 'slide_height')

    @pytest.fixture
    def part_slide_width_(self, request):
        return property_mock(request, PresentationPart, 'slide_width')

    @pytest.fixture
    def slide_height(self):
        return 7654321

    @pytest.fixture
    def slide_width(self):
        return 9876543
