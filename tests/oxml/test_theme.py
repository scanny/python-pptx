# encoding: utf-8

"""
Test suite for pptx.oxml.theme module
"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.oxml.theme import CT_OfficeStyleSheet

from ..unitutil.file import snippet_text


class DescribeCT_OfficeStyleSheet(object):
    def it_can_create_a_default_theme_element(self, new_fixture):
        expected_xml = new_fixture
        theme = CT_OfficeStyleSheet.new_default()
        assert theme.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self):
        expected_xml = snippet_text("default-theme")
        return expected_xml
