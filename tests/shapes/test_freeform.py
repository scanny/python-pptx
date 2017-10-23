# encoding: utf-8

"""Unit-test suite for pptx.shapes.freeform module"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.shapes.freeform import FreeformBuilder
from pptx.shapes.shapetree import SlideShapes

from ..unitutil.mock import (
    call, initializer_mock, instance_mock, method_mock
)


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

    def it_can_add_straight_line_segments(self, add_seg_fixture):
        builder, vertices, close, add_calls, close_calls = add_seg_fixture

        return_value = builder.add_line_segments(vertices, close)

        assert builder._add_line_segment.call_args_list == add_calls
        assert builder._add_close.call_args_list == close_calls
        assert return_value is builder

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (True,  [call()]),
        (False, []),
    ])
    def add_seg_fixture(self, request, _add_line_segment_, _add_close_):
        close, close_calls = request.param
        vertices = ((1, 2), (3, 4), (5, 6))
        builder = FreeformBuilder(None, None, None, None, None)
        add_calls = [call(1, 2), call(3, 4), call(5, 6)]
        return builder, vertices, close, add_calls, close_calls

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
    def _add_close_(self, request):
        return method_mock(request, FreeformBuilder, '_add_close')

    @pytest.fixture
    def _add_line_segment_(self, request):
        return method_mock(request, FreeformBuilder, '_add_line_segment')

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, FreeformBuilder, autospec=True)

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, SlideShapes)
