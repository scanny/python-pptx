# encoding: utf-8

"""Unit-test suite for pptx.shapes.freeform module"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.shapes.freeform import FreeformBuilder
from pptx.shapes.shapetree import SlideShapes

from ..unitutil.mock import initializer_mock, instance_mock


class DescribeFreeformBuilder(object):

    def it_provides_a_constructor(self, new_fixture):
        shapes_, start_x, start_y, x_scale, y_scale = new_fixture[:5]
        _init_, start_x_int, start_y_int = new_fixture[5:]

        builder = FreeformBuilder.new(
            shapes_, start_x, start_y, x_scale, y_scale
        )

        _init_.assert_called_once_with(
            builder, shapes_, start_x_int, start_y_int, x_scale, y_scale
        )
        assert isinstance(builder, FreeformBuilder)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self, shapes_, _init_):
        start_x, start_y, x_scale, y_scale = 99.56, 200.49, 4.2, 2.4
        start_x_int, start_y_int = 100, 200
        return (
            shapes_, start_x, start_y, x_scale, y_scale, _init_, start_x_int,
            start_y_int
        )

    # fixture components -----------------------------------

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, FreeformBuilder, autospec=True)

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, SlideShapes)
