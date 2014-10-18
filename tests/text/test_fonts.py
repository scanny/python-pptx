# encoding: utf-8

"""
Test suite for pptx.text.fonts module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.text.fonts import FontFiles

from ..unitutil.mock import method_mock


class DescribeFontFiles(object):

    def it_can_find_a_system_font_file(self, find_fixture):
        family_name, is_bold, is_italic, expected_path = find_fixture
        path = FontFiles.find(family_name, is_bold, is_italic)
        assert path == expected_path

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[
        ('Foobar',  False, False, 'foobar.ttf'),
        ('Foobar',  True,  False, 'foobarb.ttf'),
        ('Barfoo',  False, True,  'barfooi.ttf'),
    ])
    def find_fixture(self, request, _installed_fonts_):
        family_name, is_bold, is_italic, expected_path = request.param
        return family_name, is_bold, is_italic, expected_path

    # fixture components -----------------------------------

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
