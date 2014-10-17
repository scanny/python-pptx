# encoding: utf-8

"""
Test suite for pptx.text.layout module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.text.layout import _LineSource, TextFitter

from ..unitutil.mock import (
    class_mock, initializer_mock, instance_mock, method_mock
)


class DescribeTextFitter(object):

    def it_can_determine_the_best_fit_font_size(self, best_fit_fixture):
        text, extents, max_size, font_file = best_fit_fixture[:4]
        _LineSource_, _init_, line_source_ = best_fit_fixture[4:7]
        _best_fit_font_size_, font_size_ = best_fit_fixture[7:]

        font_size = TextFitter.best_fit_font_size(
            text, extents, max_size, font_file
        )

        _LineSource_.assert_called_once_with(text)
        _init_.assert_called_once_with(line_source_, extents, font_file)
        _best_fit_font_size_.assert_called_once_with(max_size)
        assert font_size is font_size_

    # fixtures ---------------------------------------------

    @pytest.fixture
    def best_fit_fixture(self, _LineSource_, _init_, _best_fit_font_size_):
        text, extents, max_size = 'Foobar', (19, 20), 42
        font_file = 'foobar.ttf'
        line_source_ = _LineSource_.return_value
        font_size_ = _best_fit_font_size_.return_value
        return (
            text, extents, max_size, font_file, _LineSource_,
            _init_, line_source_, _best_fit_font_size_, font_size_
        )

    # fixture components -----------------------------------

    @pytest.fixture
    def _best_fit_font_size_(self, request):
        return method_mock(request, TextFitter, '_best_fit_font_size')

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, TextFitter)

    @pytest.fixture
    def _LineSource_(self, request, line_source_):
        return class_mock(
            request, 'pptx.text.layout._LineSource',
            return_value=line_source_
        )

    @pytest.fixture
    def line_source_(self, request):
        return instance_mock(request, _LineSource)
