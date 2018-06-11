# encoding: utf-8

"""Test suite for pptx.dml.effect module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.dml.effect import ShadowFormat

from ..unitutil.cxml import element


class DescribeShadowFormat(object):

    def it_knows_whether_it_inherits(self, inherit_get_fixture):
        shadow, expected_value = inherit_get_fixture
        inherit = shadow.inherit
        assert inherit is expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('p:spPr', True),
        ('p:spPr/a:effectLst', False),
        ('p:grpSpPr', True),
        ('p:grpSpPr/a:effectLst', False),
    ])
    def inherit_get_fixture(self, request):
        cxml, expected_value = request.param
        shadow = ShadowFormat(element(cxml))
        return shadow, expected_value
