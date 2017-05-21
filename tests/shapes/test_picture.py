# encoding: utf-8

"""
Test suite for pptx.shapes.picture module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_MEDIA_TYPE
from pptx.parts.image import Image
from pptx.parts.slide import SlidePart
from pptx.shapes.picture import _BasePicture, _MediaFormat, Movie, Picture
from pptx.util import Pt

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, instance_mock, property_mock


class Describe_BasePicture(object):

    def it_knows_its_cropping(self, crop_get_fixture):
        picture, prop_name, expected_value = crop_get_fixture
        crop = getattr(picture, prop_name)
        assert abs(crop - expected_value) < 0.000001

    def it_provides_access_to_its_outline(self, line_fixture):
        picture = line_fixture
        line = picture.line
        assert isinstance(line, LineFormat)
        # exercise line to test has_line interface, .ln and .get_or_add_ln()
        picture.line.width = Pt(1)
        assert picture.line.width == Pt(1)

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('p:pic/p:blipFill',                     'left',    0.0),
        ('p:pic/p:blipFill/a:srcRect',           'top',     0.0),
        ('p:pic/p:blipFill/a:srcRect{l=99999}',  'bottom',  0.0),
        ('p:pic/p:blipFill/a:srcRect{l=42424}',  'left',    0.42424),
        ('p:pic/p:blipFill/a:srcRect{t=-10000}', 'top',    -0.1),
        ('p:pic/p:blipFill/a:srcRect{r=250000}', 'right',   2.5),
        ('p:pic/p:blipFill/a:srcRect{b=33333}',  'bottom',  0.33333),
    ])
    def crop_get_fixture(self, request):
        pic_cxml, side, expected_value = request.param
        picture = Picture(element(pic_cxml), None)
        prop_name = 'crop_%s' % side
        return picture, prop_name, expected_value

    @pytest.fixture
    def line_fixture(self):
        return _BasePicture(element('p:pic/p:spPr'), None)


class DescribeMovie(object):

    def it_knows_its_shape_type(self, shape_type_fixture):
        movie = shape_type_fixture
        assert movie.shape_type == MSO_SHAPE_TYPE.MEDIA

    def it_knows_its_media_type(self, media_type_fixture):
        movie = media_type_fixture
        assert movie.media_type == PP_MEDIA_TYPE.MOVIE

    def it_provides_access_to_its_media_format(self, format_fixture):
        movie, MediaFormat_, pic, parent, media_format_ = format_fixture
        media_format = movie.media_format
        MediaFormat_.assert_called_once_with(pic, parent)
        assert media_format is media_format_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def format_fixture(self, _MediaFormat_, media_format_):
        pic = element('p:pic')
        parent = movie = Movie(pic, None)
        return movie, _MediaFormat_, pic, parent, media_format_

    @pytest.fixture
    def media_type_fixture(self):
        return Movie(None, None)

    @pytest.fixture
    def shape_type_fixture(self):
        return Movie(None, None)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _MediaFormat_(self, request, media_format_):
        return class_mock(
            request, 'pptx.shapes.picture._MediaFormat',
            return_value=media_format_
        )

    @pytest.fixture
    def media_format_(self, request):
        return instance_mock(request, _MediaFormat)


class DescribePicture(object):

    def it_knows_its_shape_type(self, shape_type_fixture):
        picture = shape_type_fixture
        assert picture.shape_type == MSO_SHAPE_TYPE.PICTURE

    def it_provides_access_to_its_image(self, image_fixture):
        picture, slide_, rId, image_ = image_fixture
        image = picture.image
        slide_.get_image.assert_called_once_with(rId)
        assert image is image_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def image_fixture(self, part_prop_, slide_, image_):
        pic_cxml, rId = 'p:pic/p:blipFill/a:blip{r:embed=rId42}', 'rId42'
        picture = Picture(element(pic_cxml), None)
        slide_.get_image.return_value = image_
        return picture, slide_, rId, image_

    @pytest.fixture
    def shape_type_fixture(self):
        return Picture(None, None)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def part_prop_(self, request, slide_):
        return property_mock(request, Picture, 'part', return_value=slide_)

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, SlidePart)
