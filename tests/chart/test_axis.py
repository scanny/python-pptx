# encoding: utf-8

"""
Test suite for pptx.chart module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.axis import _BaseAxis
from pptx.enum.chart import XL_TICK_MARK

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

    def it_knows_the_scale_maximum(self, maximum_scale_get_fixture):
        axis, expected_value = maximum_scale_get_fixture
        assert axis.maximum_scale == expected_value

    def it_can_change_the_scale_maximum(self, maximum_scale_set_fixture):
        axis, new_value, expected_xml = maximum_scale_set_fixture
        axis.maximum_scale = new_value
        assert axis._element.xml == expected_xml

    def it_knows_the_scale_minimum(self, minimum_scale_get_fixture):
        axis, expected_value = minimum_scale_get_fixture
        assert axis.minimum_scale == expected_value

    def it_can_change_the_scale_minimum(self, minimum_scale_set_fixture):
        axis, new_value, expected_xml = minimum_scale_set_fixture
        axis.minimum_scale = new_value
        assert axis._element.xml == expected_xml

    def it_knows_its_major_tick_setting(self, major_tick_get_fixture):
        axis, expected_value = major_tick_get_fixture
        assert axis.major_tick_mark == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:catAx',                          XL_TICK_MARK.CROSS),
        ('c:catAx/c:majorTickMark',          XL_TICK_MARK.CROSS),
        ('c:catAx/c:majorTickMark{val=out}', XL_TICK_MARK.OUTSIDE),
    ])
    def major_tick_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        axis = _BaseAxis(element(xAx_cxml))
        return axis, expected_value

    @pytest.fixture(params=[
        ('c:catAx/c:scaling/c:max{val=12.34}', 12.34),
        ('c:valAx/c:scaling/c:max{val=23.45}', 23.45),
        ('c:catAx/c:scaling',                  None),
        ('c:valAx/c:scaling',                  None),
    ])
    def maximum_scale_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        axis = _BaseAxis(element(xAx_cxml))
        return axis, expected_value

    @pytest.fixture(params=[
        ('c:catAx/c:scaling', 34.56, 'c:catAx/c:scaling/c:max{val=34.56}'),
        ('c:valAx/c:scaling', 45.67, 'c:valAx/c:scaling/c:max{val=45.67}'),
        ('c:catAx/c:scaling', None,  'c:catAx/c:scaling'),
        ('c:valAx/c:scaling/c:max{val=42.42}', 12.34,
         'c:valAx/c:scaling/c:max{val=12.34}'),
        ('c:catAx/c:scaling/c:max{val=42.42}', None,
         'c:catAx/c:scaling'),
    ])
    def maximum_scale_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return axis, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:catAx/c:scaling/c:min{val=12.34}', 12.34),
        ('c:valAx/c:scaling/c:min{val=23.45}', 23.45),
        ('c:catAx/c:scaling',                  None),
        ('c:valAx/c:scaling',                  None),
    ])
    def minimum_scale_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        axis = _BaseAxis(element(xAx_cxml))
        return axis, expected_value

    @pytest.fixture(params=[
        ('c:catAx/c:scaling', 34.56, 'c:catAx/c:scaling/c:min{val=34.56}'),
        ('c:valAx/c:scaling', 45.67, 'c:valAx/c:scaling/c:min{val=45.67}'),
        ('c:catAx/c:scaling', None,  'c:catAx/c:scaling'),
        ('c:valAx/c:scaling/c:min{val=42.42}', 12.34,
         'c:valAx/c:scaling/c:min{val=12.34}'),
        ('c:catAx/c:scaling/c:min{val=42.42}', None,
         'c:catAx/c:scaling'),
    ])
    def minimum_scale_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return axis, new_value, expected_xml

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
        ('c:valAx/c:delete',        True,  'c:valAx/c:delete{val=0}'),
        ('c:catAx/c:delete',        False, 'c:catAx/c:delete'),
        ('c:catAx/c:delete{val=1}', True,  'c:catAx/c:delete{val=0}'),
        ('c:valAx/c:delete{val=1}', False, 'c:valAx/c:delete'),
        ('c:valAx/c:delete{val=0}', True,  'c:valAx/c:delete{val=0}'),
        ('c:catAx/c:delete{val=0}', False, 'c:catAx/c:delete'),
    ])
    def visible_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return axis, new_value, expected_xml
