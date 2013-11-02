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


class DescribeCT_SlideIdList(object):

    def it_is_used_by_the_parser_for_a_sldIdLst_element(self, sldIdLst_elm):
        assert isinstance(sldIdLst_elm, CT_SlideIdList)

    def it_can_add_a_sldId_element_as_a_child(
            self, sldIdLst_elm, sldIdLst_with_sldId_xml):
        sldIdLst_elm.add_sldId(256, 'rId1')
        assert actual_xml(sldIdLst_elm) == sldIdLst_with_sldId_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def sldIdLst_bldr(self):
        return a_sldIdLst().with_nsdecls()

    @pytest.fixture
    def sldIdLst_elm(self, sldIdLst_bldr):
        return sldIdLst_bldr.element

    @pytest.fixture
    def sldIdLst_with_sldId_xml(self, sldIdLst_bldr):
        sldId_bldr = a_sldId().with_id(256).with_rId('rId1')
        sldIdLst_bldr = sldIdLst_bldr.with_sldId(sldId_bldr)
        return sldIdLst_bldr.xml


class DescribeCT_SlideId(object):

    def it_is_used_by_the_parser_for_a_sldId_element(self):
        elm = a_sldId().with_nsdecls().element
        assert isinstance(elm, CT_SlideId)
