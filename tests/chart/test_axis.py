# encoding: utf-8

"""
Test suite for pptx.chart module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.axis import _BaseAxis, TickLabels, ValueAxis
from pptx.enum.chart import (
    XL_TICK_LABEL_POSITION as XL_TICK_LBL_POS, XL_TICK_MARK
)
from pptx.text.text import Font

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class Describe_BaseAxis(object):

    def it_knows_whether_it_has_major_gridlines(
            self, major_gridlines_get_fixture):
        base_axis, expected_value = major_gridlines_get_fixture
        assert base_axis.has_major_gridlines is expected_value

    def it_can_change_whether_it_has_major_gridlines(
            self, major_gridlines_set_fixture):
        base_axis, new_value, expected_xml = major_gridlines_set_fixture
        base_axis.has_major_gridlines = new_value
        assert base_axis._element.xml == expected_xml

    def it_knows_whether_it_has_minor_gridlines(
            self, minor_gridlines_get_fixture):
        base_axis, expected_value = minor_gridlines_get_fixture
        assert base_axis.has_minor_gridlines is expected_value

    def it_can_change_whether_it_has_minor_gridlines(
            self, minor_gridlines_set_fixture):
        base_axis, new_value, expected_xml = minor_gridlines_set_fixture
        base_axis.has_minor_gridlines = new_value
        assert base_axis._element.xml == expected_xml

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

    def it_can_change_its_major_tick_mark(self, major_tick_set_fixture):
        axis, new_value, expected_xml = major_tick_set_fixture
        axis.major_tick_mark = new_value
        assert axis._element.xml == expected_xml

    def it_knows_its_minor_tick_setting(self, minor_tick_get_fixture):
        axis, expected_value = minor_tick_get_fixture
        assert axis.minor_tick_mark == expected_value

    def it_can_change_its_minor_tick_mark(self, minor_tick_set_fixture):
        axis, new_value, expected_xml = minor_tick_set_fixture
        axis.minor_tick_mark = new_value
        assert axis._element.xml == expected_xml

    def it_knows_its_tick_label_position(self, tick_lbl_pos_get_fixture):
        axis, expected_value = tick_lbl_pos_get_fixture
        assert axis.tick_label_position == expected_value

    def it_can_change_its_tick_label_position(self, tick_lbl_pos_set_fixture):
        axis, new_value, expected_xml = tick_lbl_pos_set_fixture
        axis.tick_label_position = new_value
        assert axis._element.xml == expected_xml

    def it_provides_access_to_the_tick_labels(self, tick_labels_fixture):
        axis, tick_labels_, TickLabels_, xAx = tick_labels_fixture
        tick_labels = axis.tick_labels
        TickLabels_.assert_called_once_with(xAx)
        assert tick_labels is tick_labels_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:valAx',                  False),
        ('c:valAx/c:majorGridlines', True),
    ])
    def major_gridlines_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        base_axis = _BaseAxis(element(xAx_cxml))
        return base_axis, expected_value

    @pytest.fixture(params=[
        ('c:catAx',                  True,  'c:catAx/c:majorGridlines'),
        ('c:catAx/c:majorGridlines', True,  'c:catAx/c:majorGridlines'),
        ('c:catAx',                  False, 'c:catAx'),
        ('c:catAx/c:majorGridlines', False, 'c:catAx'),
    ])
    def major_gridlines_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        base_axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return base_axis, new_value, expected_xml

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
        ('c:catAx',                         XL_TICK_MARK.INSIDE,
         'c:catAx/c:majorTickMark{val=in}'),
        ('c:catAx',                         XL_TICK_MARK.CROSS,
         'c:catAx'),
        ('c:catAx/c:majorTickMark{val=in}', XL_TICK_MARK.OUTSIDE,
         'c:catAx/c:majorTickMark{val=out}'),
        ('c:catAx/c:majorTickMark{val=in}', XL_TICK_MARK.CROSS,
         'c:catAx'),
    ])
    def major_tick_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return axis, new_value, expected_xml

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
        ('c:catAx',                  False),
        ('c:catAx/c:minorGridlines', True),
    ])
    def minor_gridlines_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        base_axis = _BaseAxis(element(xAx_cxml))
        return base_axis, expected_value

    @pytest.fixture(params=[
        ('c:catAx',                  True,  'c:catAx/c:minorGridlines'),
        ('c:catAx/c:minorGridlines', True,  'c:catAx/c:minorGridlines'),
        ('c:catAx',                  False, 'c:catAx'),
        ('c:catAx/c:minorGridlines', False, 'c:catAx'),
    ])
    def minor_gridlines_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        base_axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return base_axis, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:valAx',                            XL_TICK_MARK.CROSS),
        ('c:valAx/c:minorTickMark',            XL_TICK_MARK.CROSS),
        ('c:valAx/c:minorTickMark{val=cross}', XL_TICK_MARK.CROSS),
        ('c:valAx/c:minorTickMark{val=out}',   XL_TICK_MARK.OUTSIDE),
    ])
    def minor_tick_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        axis = _BaseAxis(element(xAx_cxml))
        return axis, expected_value

    @pytest.fixture(params=[
        ('c:valAx',                         XL_TICK_MARK.INSIDE,
         'c:valAx/c:minorTickMark{val=in}'),
        ('c:valAx',                         XL_TICK_MARK.CROSS,
         'c:valAx'),
        ('c:valAx/c:minorTickMark{val=in}', XL_TICK_MARK.OUTSIDE,
         'c:valAx/c:minorTickMark{val=out}'),
        ('c:valAx/c:minorTickMark{val=in}', XL_TICK_MARK.CROSS,
         'c:valAx'),
    ])
    def minor_tick_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        axis = _BaseAxis(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return axis, new_value, expected_xml

    @pytest.fixture
    def tick_labels_fixture(self, TickLabels_, tick_labels_):
        xAx = element('c:valAx')
        axis = _BaseAxis(xAx)
        return axis, tick_labels_, TickLabels_, xAx

    @pytest.fixture(params=[
        ('c:valAx',                          XL_TICK_LBL_POS.NEXT_TO_AXIS),
        ('c:catAx/c:tickLblPos',             XL_TICK_LBL_POS.NEXT_TO_AXIS),
        ('c:valAx/c:tickLblPos{val=nextTo}', XL_TICK_LBL_POS.NEXT_TO_AXIS),
        ('c:catAx/c:tickLblPos{val=low}',    XL_TICK_LBL_POS.LOW),
    ])
    def tick_lbl_pos_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        axis = _BaseAxis(element(xAx_cxml))
        return axis, expected_value

    @pytest.fixture(params=[
        ('c:valAx',                           XL_TICK_LBL_POS.LOW,
         'c:valAx/c:tickLblPos{val=low}'),
        ('c:valAx',                           XL_TICK_LBL_POS.NEXT_TO_AXIS,
         'c:valAx/c:tickLblPos{val=nextTo}'),
        ('c:valAx/c:tickLblPos{val=low}',     XL_TICK_LBL_POS.HIGH,
         'c:valAx/c:tickLblPos{val=high}'),
        ('c:valAx/c:tickLblPos{val=low}',     XL_TICK_LBL_POS.NEXT_TO_AXIS,
         'c:valAx/c:tickLblPos{val=nextTo}'),
        ('c:valAx/c:tickLblPos{val=low}',     None,
         'c:valAx/c:tickLblPos'),
    ])
    def tick_lbl_pos_set_fixture(self, request):
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

    # fixture components ---------------------------------------------

    @pytest.fixture
    def TickLabels_(self, request, tick_labels_):
        return class_mock(
            request, 'pptx.chart.axis.TickLabels',
            return_value=tick_labels_
        )

    @pytest.fixture
    def tick_labels_(self, request):
        return instance_mock(request, TickLabels)


class DescribeTickLabels(object):

    def it_provides_access_to_its_font(self, font_fixture):
        tick_labels, Font_, defRPr, font_ = font_fixture
        font = tick_labels.font
        Font_.assert_called_once_with(defRPr)
        assert font is font_

    def it_adds_a_txPr_to_help_font(self, txPr_fixture):
        tick_labels, expected_xml = txPr_fixture
        tick_labels.font
        assert tick_labels._element.xml == expected_xml

    def it_knows_its_number_format(self, number_format_get_fixture):
        tick_labels, expected_value = number_format_get_fixture
        assert tick_labels.number_format == expected_value

    def it_can_change_its_number_format(self, number_format_set_fixture):
        tick_labels, new_value, expected_xml = number_format_set_fixture
        tick_labels.number_format = new_value
        assert tick_labels._element.xml == expected_xml

    def it_knows_whether_its_number_format_is_linked(
            self, number_format_is_linked_get_fixture):
        tick_labels, expected_value = number_format_is_linked_get_fixture
        assert tick_labels.number_format_is_linked is expected_value

    def it_can_change_whether_its_number_format_is_linked(
            self, number_format_is_linked_set_fixture):
        tick_labels, new_value, expected_xml = (
            number_format_is_linked_set_fixture
        )
        tick_labels.number_format_is_linked = new_value
        assert tick_labels._element.xml == expected_xml

    def it_knows_its_offset(self, offset_get_fixture):
        tick_labels, expected_value = offset_get_fixture
        assert tick_labels.offset == expected_value

    def it_can_change_its_offset(self, offset_set_fixture):
        tick_labels, new_value, expected_xml = offset_set_fixture
        tick_labels.offset = new_value
        assert tick_labels._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def font_fixture(self, Font_, font_):
        catAx = element('c:catAx/c:txPr/a:p/a:pPr/a:defRPr')
        defRPr = catAx.xpath('.//a:defRPr')[0]
        tick_labels = TickLabels(catAx)
        return tick_labels, Font_, defRPr, font_

    @pytest.fixture(params=[
        ('c:catAx',                              'General'),
        ('c:valAx/c:numFmt{formatCode=General}', 'General'),
    ])
    def number_format_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        tick_labels = TickLabels(element(xAx_cxml))
        return tick_labels, expected_value

    @pytest.fixture(params=[
        ('c:catAx', 'General',
         'c:catAx/c:numFmt{formatCode=General,sourceLinked=0}'),
        ('c:valAx/c:numFmt{formatCode=General}', '00.00',
         'c:valAx/c:numFmt{formatCode=00.00,sourceLinked=0}'),
    ])
    def number_format_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        tick_labels = TickLabels(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return tick_labels, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:catAx',                          False),
        ('c:valAx/c:numFmt',                 True),
        ('c:valAx/c:numFmt{sourceLinked=0}', False),
        ('c:catAx/c:numFmt{sourceLinked=1}', True),
    ])
    def number_format_is_linked_get_fixture(self, request):
        xAx_cxml, expected_value = request.param
        tick_labels = TickLabels(element(xAx_cxml))
        return tick_labels, expected_value

    @pytest.fixture(params=[
        ('c:valAx', True,  'c:valAx/c:numFmt{sourceLinked=1}'),
        ('c:catAx', False, 'c:catAx/c:numFmt{sourceLinked=0}'),
        ('c:valAx', None,  'c:valAx/c:numFmt'),
        ('c:catAx/c:numFmt', True, 'c:catAx/c:numFmt{sourceLinked=1}'),
        ('c:valAx/c:numFmt{sourceLinked=1}', False,
         'c:valAx/c:numFmt{sourceLinked=0}'),
    ])
    def number_format_is_linked_set_fixture(self, request):
        xAx_cxml, new_value, expected_xAx_cxml = request.param
        tick_labels = TickLabels(element(xAx_cxml))
        expected_xml = xml(expected_xAx_cxml)
        return tick_labels, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:catAx',                      100),
        ('c:catAx/c:lblOffset',          100),
        ('c:catAx/c:lblOffset{val=420}', 420),
        ('c:catAx/c:lblOffset{val=004}',   4),
        ('c:catAx/c:lblOffset{val=42%}',  42),
        ('c:catAx/c:lblOffset{val=02%}',   2),
    ])
    def offset_get_fixture(self, request):
        catAx_cxml, expected_value = request.param
        tick_labels = TickLabels(element(catAx_cxml))
        return tick_labels, expected_value

    @pytest.fixture(params=[
        ('c:catAx', 420, 'c:catAx/c:lblOffset{val=420}'),
        ('c:catAx/c:lblOffset{val=420}', 100, 'c:catAx'),
    ])
    def offset_set_fixture(self, request):
        catAx_cxml, new_value, expected_catAx_cxml = request.param
        tick_labels = TickLabels(element(catAx_cxml))
        expected_xml = xml(expected_catAx_cxml)
        return tick_labels, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:valAx{a:b=c}',
         'c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)'),
        ('c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p)',
         'c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)'),
        ('c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr)',
         'c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)'),
    ])
    def txPr_fixture(self, request):
        xAx_cxml, expected_cxml = request.param
        tick_labels = TickLabels(element(xAx_cxml))
        expected_xml = xml(expected_cxml)
        return tick_labels, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Font_(self, request, font_):
        return class_mock(
            request, 'pptx.chart.axis.Font', return_value=font_
        )

    @pytest.fixture
    def font_(self, request):
        return instance_mock(request, Font)


class DescribeValueAxis(object):

    def it_knows_its_major_unit(self, major_unit_get_fixture):
        value_axis, expected_value = major_unit_get_fixture
        assert value_axis.major_unit == expected_value

    def it_can_change_its_major_unit(self, major_unit_set_fixture):
        value_axis, new_value, expected_xml = major_unit_set_fixture
        value_axis.major_unit = new_value
        assert value_axis._element.xml == expected_xml

    def it_knows_its_minor_unit(self, minor_unit_get_fixture):
        value_axis, expected_value = minor_unit_get_fixture
        assert value_axis.minor_unit == expected_value

    def it_can_change_its_minor_unit(self, minor_unit_set_fixture):
        value_axis, new_value, expected_xml = minor_unit_set_fixture
        value_axis.minor_unit = new_value
        assert value_axis._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:valAx', None),
        ('c:valAx/c:majorUnit{val=4.2}', 4.2),
    ])
    def major_unit_get_fixture(self, request):
        valAx_cxml, expected_value = request.param
        value_axis = ValueAxis(element(valAx_cxml))
        return value_axis, expected_value

    @pytest.fixture(params=[
        ('c:valAx',                        42,
         'c:valAx/c:majorUnit{val=42.0}'),
        ('c:valAx',                        None,
         'c:valAx'),
        ('c:valAx/c:majorUnit{val=42.0}',  24.0,
         'c:valAx/c:majorUnit{val=24.0}'),
        ('c:valAx/c:majorUnit{val=42.0}',  None,
         'c:valAx'),
    ])
    def major_unit_set_fixture(self, request):
        valAx_cxml, new_value, expected_valAx_cxml = request.param
        value_axis = ValueAxis(element(valAx_cxml))
        expected_xml = xml(expected_valAx_cxml)
        return value_axis, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:valAx', None),
        ('c:valAx/c:minorUnit{val=2.4}', 2.4),
    ])
    def minor_unit_get_fixture(self, request):
        valAx_cxml, expected_value = request.param
        value_axis = ValueAxis(element(valAx_cxml))
        return value_axis, expected_value

    @pytest.fixture(params=[
        ('c:valAx',                        36,
         'c:valAx/c:minorUnit{val=36.0}'),
        ('c:valAx',                        None,
         'c:valAx'),
        ('c:valAx/c:minorUnit{val=36.0}',  12.6,
         'c:valAx/c:minorUnit{val=12.6}'),
        ('c:valAx/c:minorUnit{val=36.0}',  None,
         'c:valAx'),
    ])
    def minor_unit_set_fixture(self, request):
        valAx_cxml, new_value, expected_valAx_cxml = request.param
        value_axis = ValueAxis(element(valAx_cxml))
        expected_xml = xml(expected_valAx_cxml)
        return value_axis, new_value, expected_xml
