# encoding: utf-8

"""
Test suite for pptx.oxml.presentation module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from ..unitutil.cxml import element, xml


class DescribeCT_SlideIdList(object):
    def it_can_add_a_sldId_element_as_a_child(self, add_fixture):
        sldIdLst, expected_xml = add_fixture
        sldIdLst.add_sldId("rId1")
        assert sldIdLst.xml == expected_xml

    def it_knows_the_next_available_slide_id(self, next_id_fixture):
        sldIdLst, expected_id = next_id_fixture
        assert sldIdLst._next_id == expected_id

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self):
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=rId4,id=256}")
        expected_xml = xml(
            "p:sldIdLst/(p:sldId{r:id=rId4,id=256},p:sldId{r:id=rId1,id=257})"
        )
        return sldIdLst, expected_xml

    @pytest.fixture(
        params=[
            ("p:sldIdLst", 256),
            ("p:sldIdLst/p:sldId{id=42}", 256),
            ("p:sldIdLst/p:sldId{id=256}", 257),
            ("p:sldIdLst/(p:sldId{id=256},p:sldId{id=712})", 713),
            ("p:sldIdLst/(p:sldId{id=280},p:sldId{id=257})", 281),
        ]
    )
    def next_id_fixture(self, request):
        sldIdLst_cxml, expected_value = request.param
        sldIdLst = element(sldIdLst_cxml)
        return sldIdLst, expected_value
