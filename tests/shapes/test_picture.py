# encoding: utf-8

"""Test suite for pptx.picture module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_

from pptx.constants import MSO

from ..oxml.unitdata.autoshape import test_shapes
from ..unitutil import TestCase


class Test_Picture(TestCase):
    """Test Picture"""
    def test_shape_type_value_correct_for_picture(self):
        """Shape.shape_type value is correct for picture"""
        # setup ------------------------
        picture = test_shapes.picture
        # verify -----------------------
        assert_that(picture.shape_type, is_(equal_to(MSO.PICTURE)))
