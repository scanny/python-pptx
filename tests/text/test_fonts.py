# encoding: utf-8

"""
Test suite for pptx.text.fonts module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.text.fonts import FontFiles

from ..unitutil.mock import call, method_mock, var_mock


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

    def it_generates_font_dirs_to_help_find(self, font_dirs_fixture):
        expected_values = font_dirs_fixture
        font_dirs = FontFiles._font_directories()
        assert font_dirs == expected_values

    def it_knows_os_x_font_dirs_to_help_find(self, osx_dirs_fixture):
        expected_dirs = osx_dirs_fixture
        font_dirs = FontFiles._os_x_font_directories()
        print(font_dirs)
        print(expected_dirs)
        assert font_dirs == expected_dirs

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[
        ('Foobar',  False, False, 'foobar.ttf'),
        ('Foobar',  True,  False, 'foobarb.ttf'),
        ('Barfoo',  False, True,  'barfooi.ttf'),
    ])
    def find_fixture(self, request, _installed_fonts_):
        family_name, is_bold, is_italic, expected_path = request.param
        return family_name, is_bold, is_italic, expected_path

    @pytest.fixture(params=[
        ('darwin', ['a', 'b']),
        ('win32',  ['c', 'd']),
    ])
    def font_dirs_fixture(
            self, request, _os_x_font_directories_,
            _windows_font_directories_):
        platform, expected_dirs = request.param
        dirs_meth_mock = {
            'darwin': _os_x_font_directories_,
            'win32':  _windows_font_directories_,
        }[platform]
        sys_ = var_mock(request, 'pptx.text.fonts.sys')
        sys_.platform = platform
        dirs_meth_mock.return_value = expected_dirs
        return expected_dirs

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

    @pytest.fixture
    def osx_dirs_fixture(self, request):
        import os
        os_ = var_mock(request, 'pptx.text.fonts.os')
        os_.path = os.path
        os_.environ = {'HOME': '/Users/fbar'}
        return [
            '/Library/Fonts',
            '/Network/Library/Fonts',
            '/System/Library/Fonts',
            '/Users/fbar/Library/Fonts',
            '/Users/fbar/.fonts',
        ]

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

    @pytest.fixture
    def _os_x_font_directories_(self, request):
        return method_mock(request, FontFiles, '_os_x_font_directories')

    @pytest.fixture
    def _windows_font_directories_(self, request):
        return method_mock(request, FontFiles, '_windows_font_directories')
