# encoding: utf-8

"""Test suite for pptx.oxml.presentation module."""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from lxml import objectify

from pptx.oxml.presentation import (
    CT_Presentation, CT_SlideId, CT_SlideIdList
)

from .unitdata.presentation import a_presentation, a_sldId, a_sldIdLst
from ..unitutil import serialize_xml


def actual_xml(elm):
    objectify.deannotate(elm, cleanup_namespaces=True)
    return serialize_xml(elm, pretty_print=True)


class DescribeCT_Presentation(object):

    def it_is_used_by_the_parser_for_a_presentation_element(self):
        elm = a_presentation().element
        assert isinstance(elm, CT_Presentation)

    def it_can_get_the_sldIdLst_child_element(self):
        elm = a_presentation().with_sldIdLst().element
        sldIdLst = elm.get_or_add_sldIdLst()
        assert isinstance(sldIdLst, CT_SlideIdList)

    def it_can_add_a_sldIdLst_child_element(self):
        elm = a_presentation().with_sldSz().element
        elm.get_or_add_sldIdLst()
        expected_xml = a_presentation().with_sldIdLst().with_sldSz().xml_bytes
        assert actual_xml(elm) == expected_xml

    def it_adds_sldIdLst_before_notesSz_if_no_sldSz_elm(self):
        elm = a_presentation().with_notesSz().element
        elm.get_or_add_sldIdLst()
        expected_xml = (
            a_presentation().with_sldIdLst().with_notesSz().xml_bytes
        )
        assert actual_xml(elm) == expected_xml


class DescribeCT_SlideId(object):

    def it_is_used_by_the_parser_for_a_sldId_element(self, sldId):
        assert isinstance(sldId, CT_SlideId)

    def it_knows_the_rId(self, sldId):
        assert sldId.rId == 'rId1'


class DescribeCT_SlideIdList(object):

    def it_is_used_by_the_parser_for_a_sldIdLst_element(self, sldIdLst):
        assert isinstance(sldIdLst, CT_SlideIdList)

    def it_provides_indexed_access_to_the_sldIds(
            self, sldIdLst, sldId_xml, sldId_2_xml):
        assert actual_xml(sldIdLst[0]) == sldId_xml
        assert actual_xml(sldIdLst[1]) == sldId_2_xml

    def it_raises_IndexError_on_index_out_of_range(self, empty_sldIdLst):
        with pytest.raises(IndexError):
            empty_sldIdLst[0]

    def it_can_iterate_over_the_sldIds(
            self, sldIdLst, sldId_xml, sldId_2_xml):
        sldIds = [s for s in sldIdLst]
        assert actual_xml(sldIds[0]) == sldId_xml
        assert actual_xml(sldIds[1]) == sldId_2_xml

    def it_knows_the_sldId_count(self):
        pass

    def it_returns_sldId_count_for_len(self, sldIdLst):
        # objectify would return 1 if __len__ were not overridden
        assert len(sldIdLst) == 2

    def it_can_add_a_sldId_element_as_a_child(
            self, sldIdLst, sldIdLst_bldr):
        sldIdLst.add_sldId('rId3')
        expected_xml = (
            sldIdLst_bldr.with_sldId(
                a_sldId().with_id(258).with_rId('rId3')
            ).xml
        )
        assert actual_xml(sldIdLst) == expected_xml


# ===========================================================================
# fixtures
# ===========================================================================

@pytest.fixture
def empty_sldIdLst():
    return a_sldIdLst().with_nsdecls().element


@pytest.fixture
def sldId(sldId_bldr):
    return sldId_bldr.with_nsdecls().element


@pytest.fixture
def sldId_bldr():
    return a_sldId().with_id(256).with_rId('rId1')


@pytest.fixture
def sldId_2_bldr():
    return a_sldId().with_id(257).with_rId('rId2')


@pytest.fixture
def sldId_xml(sldId_bldr):
    return sldId_bldr.with_indent(0).with_nsdecls().xml_bytes


@pytest.fixture
def sldId_2_xml(sldId_2_bldr):
    return sldId_2_bldr.with_indent(0).with_nsdecls().xml_bytes


@pytest.fixture
def sldIdLst_bldr(sldId_bldr, sldId_2_bldr):
    b = a_sldIdLst()
    b = b.with_nsdecls()
    b = b.with_sldId(sldId_bldr)
    b = b.with_sldId(sldId_2_bldr)
    return b


@pytest.fixture
def sldIdLst(sldIdLst_bldr):
    return sldIdLst_bldr.element
