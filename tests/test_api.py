# encoding: utf-8

"""
Test suite for pptx.api module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.api import Presentation
from pptx.parts.presentation import PresentationPart

from .unitutil import property_mock


class DescribePresentation(object):

    def it_knows_its_slide_width(self, slide_width_fixture):
        prs, slide_width = slide_width_fixture
        assert prs.slide_width == slide_width

    def it_knows_its_slide_height(self, slide_height_fixture):
        prs, slide_height = slide_height_fixture
        assert prs.slide_height == slide_height

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def slide_height_fixture(self, part_slide_height_, slide_height):
        prs = Presentation()
        part_slide_height_.return_value = slide_height
        return prs, slide_height

    @pytest.fixture
    def slide_width_fixture(self, part_slide_width_, slide_width):
        prs = Presentation()
        part_slide_width_.return_value = slide_width
        return prs, slide_width

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
