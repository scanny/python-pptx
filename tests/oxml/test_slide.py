# encoding: utf-8

"""
Test suite for pptx.oxml.slide module
"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide

from ..unitutil.file import snippet_text


class DescribeCT_NotesMaster(object):
    def it_can_create_a_default_notesMaster_element(self, new_fixture):
        expected_xml = new_fixture
        notesMaster = CT_NotesMaster.new_default()
        assert notesMaster.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self):
        expected_xml = snippet_text("default-notesMaster")
        return expected_xml


class DescribeCT_NotesSlide(object):
    def it_can_create_a_new_notes_element(self, new_fixture):
        expected_xml = new_fixture
        notes = CT_NotesSlide.new()
        assert notes.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self):
        expected_xml = snippet_text("default-notes")
        return expected_xml
