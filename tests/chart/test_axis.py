# encoding: utf-8

"""
Test suite for pptx.chart module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.axis import _BaseAxis

from ..unitutil.cxml import element, xml


class Describe_BaseAxis(object):

    def it_knows_whether_it_is_visible(self, visible_get_fixture):
        axis, expected_bool_value = visible_get_fixture
        assert axis.visible is expected_bool_value

    def it_can_change_whether_it_is_visible(self, visible_set_fixture):
        axis, new_value, expected_xml = visible_set_fixture
        axis.visible = new_value
        assert axis._element.xml == expected_xml

    def it_raises_on_assign_non_bool_to_visible(self):
        axis = _BaseAxis(None)
        with pytest.raises(ValueError):
            axis.visible = 'foobar'

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:catAx',                     False),
        ('c:catAx/c:delete',            False),
        ('c:catAx/c:delete{val=0}',     True),
        ('c:catAx/c:delete{val=1}',     False),
        ('c:catAx/c:delete{val=false}', True),
        ('c:valAx',                     False),
        ('c:valAx/c:delete',            False),
        ('c:valAx/c:delete{val=0}',     True),
        ('c:valAx/c:delete{val=1}',     False),
        ('c:valAx/c:delete{val=false}', True),
    ])
    def visible_get_fixture(self, request):
        xAx_cxml, expected_bool_value = request.param
        axis = _BaseAxis(element(xAx_cxml))
        return axis, expected_bool_value

    @pytest.fixture(params=[
        ('c:catAx',                 True,  'c:catAx/c:delete{val=0}'),
        ('c:catAx',                 False, 'c:catAx/c:delete'),
        ('c:catAx/c:delete',        True,  'c:catAx/c:delete{val=0}'),
        ('c:catAx/c:delete',        False, 'c:catAx/c:delete'),
        ('c:catAx/c:delete{val=1}', True,  'c:catAx/c:delete{val=0}'),
        ('c:catAx/c:delete{val=1}', False, 'c:catAx/c:delete'),
        ('c:catAx/c:delete{val=0}', True,  'c:catAx/c:delete{val=0}'),
        ('c:catAx/c:delete{val=0}', False, 'c:catAx/c:delete'),
    ])
    def visible_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return axis, new_value, expected_xml
