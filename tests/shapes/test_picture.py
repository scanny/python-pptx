# encoding: utf-8

"""
Test suite for pptx.shapes.picture module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.picture import Picture
from pptx.util import Pt

from ..oxml.unitdata.shape import a_pic, an_spPr


class Describe_Picture(object):

    def it_has_a_line(self, picture):
        assert isinstance(picture.line, LineFormat)
        # exercise line to test has_line interface, .ln and .get_or_add_ln()
        picture.line.width = Pt(1)
        assert picture.line.width == Pt(1)

    def it_knows_its_shape_type(self, picture):
        assert picture.shape_type == MSO_SHAPE_TYPE.PICTURE

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def picture(self):
        pic = a_pic().with_nsdecls().with_child(an_spPr()).element
        return Picture(pic, None)
