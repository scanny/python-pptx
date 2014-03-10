# encoding: utf-8

"""
Test suite for pptx.oxml.slidemaster module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.oxml.slidemaster import CT_SlideMaster

from .unitdata.slides import a_sldMaster


class DescribeCT_SlideMaster(object):

    def it_is_used_by_the_parser_for_a_sldMaster_element(self, sldMaster_elm):
        assert isinstance(sldMaster_elm, CT_SlideMaster)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def sldMaster_elm(self):
        return a_sldMaster().with_nsdecls().element
