# encoding: utf-8

"""Unit-test suite for `pptx.oxml.slide` module."""

from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide

from ..unitutil.file import snippet_text


class DescribeCT_NotesMaster(object):
    """Unit-test suite for `pptx.oxml.slide.CT_NotesMaster` objects."""

    def it_can_create_a_default_notesMaster_element(self):
        notesMaster = CT_NotesMaster.new_default()
        assert notesMaster.xml == snippet_text("default-notesMaster")


class DescribeCT_NotesSlide(object):
    """Unit-test suite for `pptx.oxml.slide.CT_NotesSlide` objects."""

    def it_can_create_a_new_notes_element(self):
        notes = CT_NotesSlide.new()
        assert notes.xml == snippet_text("default-notes")
