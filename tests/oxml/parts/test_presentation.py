# encoding: utf-8

"""
Test suite for pptx.oxml.presentation module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from ..unitdata.presentation import a_sldId, a_sldIdLst


class DescribeCT_SlideIdList(object):

    def it_can_add_a_sldId_element_as_a_child(self, add_fixture):
        sldIdLst, expected_xml = add_fixture
        sldIdLst.add_sldId('rId1')
        assert sldIdLst.xml == expected_xml

    def it_knows_the_next_available_slide_id(self, next_id_fixture):
        sldIdLst, expected_id = next_id_fixture
        assert sldIdLst._next_id == expected_id

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self):
        sldIdLst = a_sldIdLst().with_nsdecls().element
        expected_xml = (
            a_sldIdLst().with_nsdecls('p', 'r').with_child(
                a_sldId().with_id(256).with_rId('rId1'))
        ).xml()
        return sldIdLst, expected_xml

    @pytest.fixture(params=[
        ((), 256),
        ((256,), 257), ((257,), 256), ((300,), 256), ((255,), 256),
        ((257, 259), 256), ((256, 258), 257), ((256, 257), 258),
        ((257, 258, 259), 256), ((256, 258, 259), 257),
        ((256, 257, 259), 258), ((258, 256, 257), 259),
    ])
    def next_id_fixture(self, request):
        existing_ids, expected_id = request.param
        sldIdLst_bldr = a_sldIdLst().with_nsdecls()
        for n in existing_ids:
            sldIdLst_bldr.with_child(a_sldId().with_id(n))
        sldIdLst = sldIdLst_bldr.element
        return sldIdLst, expected_id
