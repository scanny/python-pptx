# encoding: utf-8

"""
Test suite for pptx.oxml.presentation module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.oxml.presentation import CT_SlideIdList

from .unitdata.presentation import a_sldId, a_sldIdLst


class DescribeCT_SlideIdList(object):

    def it_is_used_by_the_parser_for_a_sldIdLst_element(self, sldIdLst):
        assert isinstance(sldIdLst, CT_SlideIdList)

    def it_provides_indexed_access_to_the_sldIds(self, sldIdLst_with_sldIds):
        sldIdLst, sldId_xml, sldId_2_xml = sldIdLst_with_sldIds
        assert sldIdLst[0].xml == sldId_xml
        assert sldIdLst[1].xml == sldId_2_xml

    def it_raises_IndexError_on_index_out_of_range(self, sldIdLst):
        with pytest.raises(IndexError):
            sldIdLst[0]

    def it_can_iterate_over_the_sldIds(self, sldIdLst_with_sldIds):
        sldIdLst, sldId_xml, sldId_2_xml = sldIdLst_with_sldIds
        sldIds = [s for s in sldIdLst]
        assert sldIds[0].xml == sldId_xml
        assert sldIds[1].xml == sldId_2_xml

    def it_returns_sldId_count_for_len(self, sldIdLst_with_sldIds):
        sldIdLst = sldIdLst_with_sldIds[0]
        assert len(sldIdLst) == 2

    def it_can_add_a_sldId_element_as_a_child(
            self, sldIdLst, sldIdLst_with_sldId_xml):
        sldIdLst.add_sldId('rId1')
        assert sldIdLst.xml == sldIdLst_with_sldId_xml

    def it_knows_the_next_available_slide_id(self, next_id_fixture):
        sldIdLst, expected_id = next_id_fixture
        print(sldIdLst.xml)
        assert sldIdLst._next_id == expected_id

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ((), '256'),
        ((256,), '257'), ((257,), '256'), ((300,), '256'), ((255,), '256'),
        ((257, 259), '256'), ((256, 258), '257'), ((256, 257), '258'),
        ((257, 258, 259), '256'), ((256, 258, 259), '257'),
        ((256, 257, 259), '258'), ((258, 256, 257), '259'),
    ])
    def next_id_fixture(self, request):
        existing_ids, expected_id = request.param
        sldIdLst_bldr = a_sldIdLst().with_nsdecls()
        for n in existing_ids:
            sldIdLst_bldr.with_child(a_sldId().with_id(n))
        sldIdLst = sldIdLst_bldr.element
        return sldIdLst, expected_id

    # fixture components ---------------------------------------------

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
