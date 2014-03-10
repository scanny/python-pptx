# encoding: utf-8

"""
Test suite for pptx.oxml.presentation module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.oxml.presentation import (
    CT_Presentation, CT_SlideId, CT_SlideIdList
)

from .unitdata.presentation import (
    a_notesSz, a_presentation, a_sldId, a_sldIdLst, a_sldSz
)
from ..unitutil import actual_xml


class DescribeCT_Presentation(object):

    def it_is_used_by_the_parser_for_a_presentation_element(self, prs_elm):
        assert isinstance(prs_elm, CT_Presentation)

    def it_can_get_the_sldIdLst_child_element(self, prs_with_sldIdLst):
        sldIdLst = prs_with_sldIdLst.get_or_add_sldIdLst()
        assert isinstance(sldIdLst, CT_SlideIdList)

    def it_can_add_a_sldIdLst_child_element(
            self, prs_with_sldSz, prs_with_sldIdLst_sldSz_xml):
        prs = prs_with_sldSz
        prs.get_or_add_sldIdLst()
        assert actual_xml(prs) == prs_with_sldIdLst_sldSz_xml

    def it_adds_sldIdLst_before_notesSz_if_no_sldSz_elm(
            self, prs_with_notesSz, prs_with_sldIdLst_notesSz_xml):
        prs = prs_with_notesSz
        prs.get_or_add_sldIdLst()
        assert actual_xml(prs) == prs_with_sldIdLst_notesSz_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def prs_elm(self):
        return a_presentation().with_nsdecls().element

    @pytest.fixture
    def prs_with_notesSz(self):
        return (
            a_presentation().with_nsdecls()
                            .with_child(a_notesSz())
                            .element
        )

    @pytest.fixture
    def prs_with_sldIdLst(self):
        return (
            a_presentation().with_nsdecls()
                            .with_child(a_sldIdLst())
                            .element
        )

    @pytest.fixture
    def prs_with_sldIdLst_notesSz_xml(self):
        return (
            a_presentation().with_nsdecls()
                            .with_child(a_sldIdLst())
                            .with_child(a_notesSz())
                            .xml()
        )

    @pytest.fixture
    def prs_with_sldIdLst_sldSz_xml(self):
        return (
            a_presentation().with_nsdecls()
                            .with_child(a_sldIdLst())
                            .with_child(a_sldSz())
                            .xml()
        )

    @pytest.fixture
    def prs_with_sldSz(self):
        return (
            a_presentation().with_nsdecls()
                            .with_child(a_sldSz())
                            .element
        )


class DescribeCT_SlideId(object):

    def it_is_used_by_the_parser_for_a_sldId_element(self, sldId):
        assert isinstance(sldId, CT_SlideId)

    def it_knows_the_rId(self, sldId):
        assert sldId.rId == 'rId1'

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def sldId(self):
        return a_sldId().with_nsdecls('p', 'r').with_rId('rId1').element


class DescribeCT_SlideIdList(object):

    def it_is_used_by_the_parser_for_a_sldIdLst_element(self, sldIdLst):
        assert isinstance(sldIdLst, CT_SlideIdList)

    def it_provides_indexed_access_to_the_sldIds(self, sldIdLst_with_sldIds):
        sldIdLst, sldId_xml, sldId_2_xml = sldIdLst_with_sldIds
        assert actual_xml(sldIdLst[0]) == sldId_xml
        assert actual_xml(sldIdLst[1]) == sldId_2_xml

    def it_raises_IndexError_on_index_out_of_range(self, sldIdLst):
        with pytest.raises(IndexError):
            sldIdLst[0]

    def it_can_iterate_over_the_sldIds(self, sldIdLst_with_sldIds):
        sldIdLst, sldId_xml, sldId_2_xml = sldIdLst_with_sldIds
        sldIds = [s for s in sldIdLst]
        assert actual_xml(sldIds[0]) == sldId_xml
        assert actual_xml(sldIds[1]) == sldId_2_xml

    def it_knows_the_sldId_count(self):
        pass

    def it_returns_sldId_count_for_len(self, sldIdLst_with_sldIds):
        # objectify would return 1 if __len__ were not overridden
        sldIdLst = sldIdLst_with_sldIds[0]
        assert len(sldIdLst) == 2

    def it_can_add_a_sldId_element_as_a_child(
            self, sldIdLst, sldIdLst_with_sldId_xml):
        sldIdLst.add_sldId('rId1')
        assert actual_xml(sldIdLst) == sldIdLst_with_sldId_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def sldIdLst(self):
        return a_sldIdLst().with_nsdecls().element

    @pytest.fixture
    def sldIdLst_with_sldId_xml(self):
        sldId_bldr = a_sldId().with_id(256).with_rId('rId1')
        return (
            a_sldIdLst().with_nsdecls('p', 'r')
                        .with_child(sldId_bldr)
                        .xml()
        )

    @pytest.fixture
    def sldIdLst_with_sldIds(self):
        sldId_bldr = a_sldId().with_id(256).with_rId('rId1')
        sldId_2_bldr = a_sldId().with_id(257).with_rId('rId2')
        sldIdLst = (
            a_sldIdLst().with_nsdecls('p', 'r')
                        .with_child(sldId_bldr)
                        .with_child(sldId_2_bldr)
                        .element
        )
        sldId_xml = sldId_bldr.with_nsdecls().xml()
        sldId_2_xml = sldId_2_bldr.with_nsdecls().xml()
        return sldIdLst, sldId_xml, sldId_2_xml
