# encoding: utf-8

"""
Test suite for pptx.shapes.picture module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.constants import MSO
from pptx.shapes.picture import Picture

from ..oxml.unitdata.shape import a_pic, an_ext, an_off, an_spPr, an_xfrm


class Describe_Picture(object):

    def it_has_a_position(self, picture_with_position):
        picture, left, top = picture_with_position
        assert picture.left == left
        assert picture.top == top

    def it_has_dimensions(self, picture_with_dimensions):
        picture, width, height = picture_with_dimensions
        assert picture.width == width
        assert picture.height == height

    def it_knows_its_shape_type(self, picture):
        assert picture.shape_type == MSO.PICTURE

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def picture(self):
        return Picture(None, None)

    @pytest.fixture
    def picture_with_dimensions(self):
        width, height = 321, 654
        pic = (
            a_pic().with_nsdecls().with_child(
                an_spPr().with_child(
                    an_xfrm().with_child(
                        an_ext().with_cx(width).with_cy(height))))
            .element
        )
        picture = Picture(pic, None)
        return picture, width, height

    @pytest.fixture
    def picture_with_position(self):
        left, top = 123, 456
        pic = (
            a_pic().with_nsdecls().with_child(
                an_spPr().with_child(
                    an_xfrm().with_child(
                        an_off().with_x(left).with_y(top))))
            .element
        )
        picture = Picture(pic, None)
        return picture, left, top
