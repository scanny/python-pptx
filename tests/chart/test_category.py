# encoding: utf-8

"""
Unit test suite for the pptx.chart.category module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.chart.category import Categories

from ..unitutil.cxml import element


class DescribeCategories(object):

    def it_knows_its_length(self, len_fixture):
        categories, expected_len = len_fixture
        assert len(categories) == expected_len

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:barChart', 0),
        ('c:barChart/c:ser/c:cat/c:ptCount{val=4}', 4),
    ])
    def len_fixture(self, request):
        xChart_cxml, expected_len = request.param
        categories = Categories(element(xChart_cxml))
        return categories, expected_len
