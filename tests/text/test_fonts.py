# encoding: utf-8

"""
Test suite for pptx.text.fonts module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.text.fonts import FontFiles

from ..unitutil.mock import call, method_mock


class DescribeFontFiles(object):

    def it_can_find_a_system_font_file(self, find_fixture):
        family_name, is_bold, is_italic, expected_path = find_fixture
        path = FontFiles.find(family_name, is_bold, is_italic)
        assert path == expected_path

    def it_catalogs_the_system_fonts_to_help_find(self, installed_fixture):
        expected_call_args, expected_values = installed_fixture
        installed_fonts = FontFiles._installed_fonts()
        assert FontFiles._iter_font_files_in.call_args_list == (
            expected_call_args
        )
        assert installed_fonts == expected_values

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[
        ('Foobar',  False, False, 'foobar.ttf'),
        ('Foobar',  True,  False, 'foobarb.ttf'),
        ('Barfoo',  False, True,  'barfooi.ttf'),
    ])
    def find_fixture(self, request, _installed_fonts_):
        family_name, is_bold, is_italic, expected_path = request.param
        return family_name, is_bold, is_italic, expected_path

    @pytest.fixture
    def installed_fixture(self, _iter_font_files_in_, _font_directories_):
        _font_directories_.return_value = ['d', 'd_2']
        _iter_font_files_in_.side_effect = [
            [(('A', True,  False), 'a.ttf')],
            [(('B', False, True),  'b.ttf')],
        ]
        expected_call_args = [call('d'), call('d_2')]
        expected_values = {
            ('A', True,  False): 'a.ttf',
            ('B', False, True):  'b.ttf',
        }
        return expected_call_args, expected_values

    # fixture components -----------------------------------

    @pytest.fixture
    def _font_directories_(self, request):
        return method_mock(request, FontFiles, '_font_directories')

    @pytest.fixture
    def _installed_fonts_(self, request):
        _installed_fonts_ = method_mock(
            request, FontFiles, '_installed_fonts'
        )
        _installed_fonts_.return_value = {
            ('Foobar',  False, False): 'foobar.ttf',
            ('Foobar',  True,  False): 'foobarb.ttf',
            ('Barfoo',  False, True):  'barfooi.ttf',
        }
        return _installed_fonts_

    @pytest.fixture
    def _iter_font_files_in_(self, request):
        return method_mock(request, FontFiles, '_iter_font_files_in')
