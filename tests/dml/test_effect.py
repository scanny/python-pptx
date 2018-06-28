# encoding: utf-8

"""Test suite for pptx.dml.effect module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.dml.effect import ShadowFormat

from ..unitutil.cxml import element, xml


class DescribeShadowFormat(object):

    def it_knows_whether_it_inherits(self, inherit_get_fixture):
        shadow, expected_value = inherit_get_fixture
        inherit = shadow.inherit
        assert inherit is expected_value

    def it_can_change_whether_it_inherits(self, inherit_set_fixture):
        shadow, value, expected_xml = inherit_set_fixture
        shadow.inherit = value
        assert shadow._element.xml == expected_xml

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

    @pytest.fixture(params=[
        ('p:spPr{a:b=c}', False, 'p:spPr{a:b=c}/a:effectLst'),
        ('p:grpSpPr{a:b=c}', False, 'p:grpSpPr{a:b=c}/a:effectLst'),
        ('p:spPr{a:b=c}/a:effectLst', True, 'p:spPr{a:b=c}'),
        ('p:grpSpPr{a:b=c}/a:effectLst', True, 'p:grpSpPr{a:b=c}'),
        ('p:spPr', True, 'p:spPr'),
        ('p:grpSpPr', True, 'p:grpSpPr'),
        ('p:spPr/a:effectLst', False, 'p:spPr/a:effectLst'),
        ('p:grpSpPr/a:effectLst', False, 'p:grpSpPr/a:effectLst'),
    ])
    def inherit_set_fixture(self, request):
        cxml, value, expected_cxml = request.param
        shadow = ShadowFormat(element(cxml))
        expected_value = xml(expected_cxml)
        return shadow, value, expected_value
