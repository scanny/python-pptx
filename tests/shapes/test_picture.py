# encoding: utf-8

"""
Test suite for pptx.shapes.picture module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.parts.image import Image
from pptx.parts.slide import Slide
from pptx.shapes.picture import Picture
from pptx.util import Pt

from ..unitutil.cxml import element
from ..unitutil.mock import instance_mock, property_mock


class Describe_Picture(object):

    def it_provides_access_to_its_image(self, image_fixture):
        picture, slide_, rId, image_ = image_fixture
        image = picture.image
        slide_.get_image.assert_called_once_with(rId)
        assert image is image_

    def it_has_a_line(self, picture):
        assert isinstance(picture.line, LineFormat)
        # exercise line to test has_line interface, .ln and .get_or_add_ln()
        picture.line.width = Pt(1)
        assert picture.line.width == Pt(1)

    def it_knows_its_shape_type(self, picture):
        assert picture.shape_type == MSO_SHAPE_TYPE.PICTURE

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def image_fixture(self, part_prop_, slide_, image_):
        pic_cxml, rId = 'p:pic/p:blipFill/a:blip{r:embed=rId42}', 'rId42'
        picture = Picture(element(pic_cxml), None)
        slide_.get_image.return_value = image_
        return picture, slide_, rId, image_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def part_prop_(self, request, slide_):
        return property_mock(request, Picture, 'part', return_value=slide_)

    @pytest.fixture
    def picture(self):
        pic_cxml = 'p:pic/p:spPr'
        return Picture(element(pic_cxml), None)

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)
