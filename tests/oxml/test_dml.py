# encoding: utf-8

"""
Test suite for pptx.oxml.dml module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.oxml.dml import CT_SolidColorFillProperties

from ..oxml.unitdata.dml import a_schemeClr, a_solidFill, an_srgbClr


class DescribeCT_SolidColorFillProperties(object):

    def it_is_used_by_the_parser_for_a_solidFill_element(self, solidFill):
        assert isinstance(solidFill, CT_SolidColorFillProperties)

    def it_can_get_the_schemeClr_child_element_if_there_is_one(
            self, solidFill, solidFill_with_schemeClr, schemeClr):
        assert solidFill.schemeClr is None
        assert solidFill_with_schemeClr.schemeClr is schemeClr

    def it_can_get_the_srgbClr_child_element_if_there_is_one(
            self, solidFill, solidFill_with_srgbClr, srgbClr):
        assert solidFill.srgbClr is None
        assert solidFill_with_srgbClr.srgbClr is srgbClr

    # fixtures ---------------------------------------------

    @pytest.fixture
    def solidFill(self):
        return a_solidFill().with_nsdecls().element

    @pytest.fixture
    def solidFill_with_schemeClr(self, schemeClr):
        solidFill = a_solidFill().with_nsdecls().element
        solidFill.append(schemeClr)
        return solidFill

    @pytest.fixture
    def solidFill_with_srgbClr(self, srgbClr):
        solidFill = a_solidFill().with_nsdecls().element
        solidFill.append(srgbClr)
        return solidFill

    @pytest.fixture
    def schemeClr(self):
        return a_schemeClr().with_nsdecls().element

    @pytest.fixture
    def srgbClr(self):
        return an_srgbClr().with_nsdecls().element
