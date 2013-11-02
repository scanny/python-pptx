# encoding: utf-8

"""Test suite for pptx.oxml.presentation module."""

from __future__ import absolute_import, print_function, unicode_literals

from pptx.oxml.presentation import (
    CT_Presentation, CT_SlideId, CT_SlideIdList
)

from .unitdata.presentation import a_presentation, a_sldId, a_sldIdLst
from ..unitutil import serialize_xml


def actual_xml(elm):
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

    def it_is_used_by_the_parser_for_a_sldIdLst_element(self):
        elm = a_sldIdLst().with_nsdecls().element
        assert isinstance(elm, CT_SlideIdList)


class DescribeCT_SlideId(object):

    def it_is_used_by_the_parser_for_a_sldId_element(self):
        elm = a_sldId().with_nsdecls().element
        assert isinstance(elm, CT_SlideId)
