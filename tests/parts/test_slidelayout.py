# encoding: utf-8

"""
Test suite for pptx.parts.slidelayout module
"""

from __future__ import absolute_import

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.autoshape import CT_Shape
from pptx.parts.slidelayout import (
    _LayoutPlaceholder, _LayoutPlaceholders, _LayoutShapeFactory,
    _LayoutShapeTree, SlideLayout
)
from pptx.parts.slidemaster import _MasterPlaceholder, SlideMaster
from pptx.shapes.shape import BaseShape

from ..oxml.unitdata.shape import (
    a_ph, a_pic, an_ext, an_nvPr, an_nvSpPr, an_sp, an_spPr, an_xfrm
)
from ..unitutil import (
    function_mock, class_mock, instance_mock, method_mock, property_mock
)


class DescribeSlideLayout(object):

    def it_knows_the_slide_master_it_inherits_from(self, master_fixture):
        slide_layout, slide_master_ = master_fixture
        slide_master = slide_layout.slide_master
        slide_layout.part_related_by.assert_called_once_with(RT.SLIDE_MASTER)
        assert slide_master is slide_master_

    def it_provides_access_to_its_shapes(self, shapes_fixture):
        slide_layout, _LayoutShapeTree_, layout_shape_tree_ = shapes_fixture
        shapes = slide_layout.shapes
        _LayoutShapeTree_.assert_called_once_with(slide_layout)
        assert shapes is layout_shape_tree_

    def it_provides_access_to_its_placeholders(self, placeholders_fixture):
        slide_layout, _LayoutPlaceholders_, layout_placeholders_ = (
            placeholders_fixture
        )
        placeholders = slide_layout.placeholders
        _LayoutPlaceholders_.assert_called_once_with(slide_layout)
        assert placeholders is layout_placeholders_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def master_fixture(self, slide_master_, part_related_by_):
        slide_layout = SlideLayout(None, None, None, None)
        return slide_layout, slide_master_

    @pytest.fixture
    def placeholders_fixture(
            self, _LayoutPlaceholders_, layout_placeholders_):
        slide_layout = SlideLayout(None, None, None, None)
        return slide_layout, _LayoutPlaceholders_, layout_placeholders_

    @pytest.fixture
    def shapes_fixture(self, _LayoutShapeTree_, layout_shape_tree_):
        slide_layout = SlideLayout(None, None, None, None)
        return slide_layout, _LayoutShapeTree_, layout_shape_tree_

    # fixture components -----------------------------------

    @pytest.fixture
    def _LayoutPlaceholders_(self, request, layout_placeholders_):
        return class_mock(
            request, 'pptx.parts.slidelayout._LayoutPlaceholders',
            return_value=layout_placeholders_
        )

    @pytest.fixture
    def _LayoutShapeTree_(self, request, layout_shape_tree_):
        return class_mock(
            request, 'pptx.parts.slidelayout._LayoutShapeTree',
            return_value=layout_shape_tree_
        )

    @pytest.fixture
    def layout_placeholders_(self, request):
        return instance_mock(request, _LayoutPlaceholders)

    @pytest.fixture
    def layout_shape_tree_(self, request):
        return instance_mock(request, _LayoutShapeTree)

    @pytest.fixture
    def part_related_by_(self, request, slide_master_):
        return method_mock(
            request, SlideLayout, 'part_related_by',
            return_value=slide_master_
        )

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)


class Describe_LayoutShapeTree(object):

    def it_constructs_a_layout_placeholder_for_a_placeholder_shape(
            self, factory_fixture):
        layout_shapes, ph_elm_, _LayoutShapeFactory_, layout_placeholder_ = (
            factory_fixture
        )
        layout_placeholder = layout_shapes._shape_factory(ph_elm_)
        _LayoutShapeFactory_.assert_called_once_with(ph_elm_, layout_shapes)
        assert layout_placeholder is layout_placeholder_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def factory_fixture(
            self, ph_elm_, _LayoutShapeFactory_, layout_placeholder_):
        layout_shapes = _LayoutShapeTree(None)
        return (
            layout_shapes, ph_elm_, _LayoutShapeFactory_, layout_placeholder_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def layout_placeholder_(self, request):
        return instance_mock(request, _LayoutPlaceholder)

    @pytest.fixture
    def _LayoutShapeFactory_(self, request, layout_placeholder_):
        return function_mock(
            request, 'pptx.parts.slidelayout._LayoutShapeFactory',
            return_value=layout_placeholder_
        )

    @pytest.fixture
    def ph_elm_(self, request):
        return instance_mock(request, CT_Shape)


class Describe_LayoutShapeFactory(object):

    def it_constructs_a_layout_placeholder_for_a_shape_element(
            self, factory_fixture):
        shape_elm, parent_, ShapeConstructor_, shape_ = factory_fixture
        shape = _LayoutShapeFactory(shape_elm, parent_)
        ShapeConstructor_.assert_called_once_with(shape_elm, parent_)
        assert shape is shape_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=['ph', 'sp', 'pic'])
    def factory_fixture(
            self, request, ph_bldr, slide_layout_, _LayoutPlaceholder_,
            layout_placeholder_, BaseShapeFactory_, base_shape_):
        shape_bldr, ShapeConstructor_, shape_mock = {
            'ph':  (ph_bldr, _LayoutPlaceholder_, layout_placeholder_),
            'sp':  (an_sp(), BaseShapeFactory_,   base_shape_),
            'pic': (a_pic(), BaseShapeFactory_,   base_shape_),
        }[request.param]
        shape_elm = shape_bldr.with_nsdecls().element
        return shape_elm, slide_layout_, ShapeConstructor_, shape_mock

    # fixture components -----------------------------------

    @pytest.fixture
    def BaseShapeFactory_(self, request, base_shape_):
        return function_mock(
            request, 'pptx.parts.slidelayout.BaseShapeFactory',
            return_value=base_shape_
        )

    @pytest.fixture
    def base_shape_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def _LayoutPlaceholder_(self, request, layout_placeholder_):
        return class_mock(
            request, 'pptx.parts.slidelayout._LayoutPlaceholder',
            return_value=layout_placeholder_
        )

    @pytest.fixture
    def layout_placeholder_(self, request):
        return instance_mock(request, _LayoutPlaceholder)

    @pytest.fixture
    def ph_bldr(self):
        return (
            an_sp().with_child(
                an_nvSpPr().with_child(
                    an_nvPr().with_child(
                        a_ph().with_idx(1))))
        )

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)


class Describe_LayoutPlaceholders(object):

    def it_constructs_a_layout_placeholder_for_a_placeholder_shape(
            self, factory_fixture):
        layout_placeholders, ph_elm_ = factory_fixture[:2]
        _LayoutShapeFactory_, layout_placeholder_ = factory_fixture[2:]
        layout_placeholder = layout_placeholders._shape_factory(ph_elm_)
        _LayoutShapeFactory_.assert_called_once_with(
            ph_elm_, layout_placeholders
        )
        assert layout_placeholder is layout_placeholder_

    def it_can_find_a_placeholder_by_idx_value(self, get_fixture):
        layout_placeholders, ph_idx, placeholder_ = get_fixture
        placeholder = layout_placeholders.get(idx=ph_idx)
        assert placeholder is placeholder_

    def it_returns_default_if_placeholder_having_idx_not_found(
            self, default_fixture):
        layout_placeholders = default_fixture
        default = 'barfoo'
        placeholder = layout_placeholders.get('foobar', default)
        assert placeholder is default

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def default_fixture(self, _iter_):
        layout_placeholders = _LayoutPlaceholders(None)
        return layout_placeholders

    @pytest.fixture
    def factory_fixture(
            self, ph_elm_, _LayoutShapeFactory_, layout_placeholder_):
        layout_placeholders = _LayoutPlaceholders(None)
        return (
            layout_placeholders, ph_elm_, _LayoutShapeFactory_,
            layout_placeholder_
        )

    @pytest.fixture(params=[0, 1])
    def get_fixture(self, request, _iter_, placeholder_, placeholder_2_):
        layout_placeholders = _LayoutPlaceholders(None)
        ph_idx = request.param
        ph_shape_ = {0: placeholder_, 1: placeholder_2_}[request.param]
        return layout_placeholders, ph_idx, ph_shape_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _iter_(self, request, placeholder_, placeholder_2_):
        return method_mock(
            request, _LayoutPlaceholders, '__iter__',
            return_value=iter([placeholder_, placeholder_2_])
        )

    @pytest.fixture
    def layout_placeholder_(self, request):
        return instance_mock(request, _LayoutPlaceholder)

    @pytest.fixture
    def _LayoutShapeFactory_(self, request, layout_placeholder_):
        return function_mock(
            request, 'pptx.parts.slidelayout._LayoutShapeFactory',
            return_value=layout_placeholder_
        )

    @pytest.fixture
    def ph_elm_(self, request):
        return instance_mock(request, CT_Shape)

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, _LayoutPlaceholder, idx=0)

    @pytest.fixture
    def placeholder_2_(self, request):
        return instance_mock(request, _LayoutPlaceholder, idx=1)


class Describe_LayoutPlaceholder(object):

    def it_considers_inheritance_when_computing_pos_and_size(
            self, xfrm_fixture):
        layout_placeholder, _direct_or_inherited_value_ = xfrm_fixture[:2]
        attr_name, expected_value = xfrm_fixture[2:]
        value = getattr(layout_placeholder, attr_name)
        _direct_or_inherited_value_.assert_called_once_with(attr_name)
        assert value == expected_value

    def it_provides_direct_property_values_when_they_exist(
            self, direct_fixture):
        layout_placeholder, expected_width = direct_fixture
        width = layout_placeholder.width
        assert width == expected_width

    def it_provides_inherited_property_values_when_no_direct_value(
            self, inherited_fixture):
        layout_placeholder, _inherited_value_, inherited_left_ = (
            inherited_fixture
        )
        left = layout_placeholder.left
        _inherited_value_.assert_called_once_with('left')
        assert left == inherited_left_

    def it_knows_how_to_get_a_property_value_from_its_master(
            self, mstr_val_fixture):
        layout_placeholder, attr_name, expected_value = mstr_val_fixture
        value = layout_placeholder._inherited_value(attr_name)
        assert value == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def direct_fixture(self, sp, width):
        layout_placeholder = _LayoutPlaceholder(sp, None)
        return layout_placeholder, width

    @pytest.fixture
    def inherited_fixture(self, sp, _inherited_value_, int_value_):
        layout_placeholder = _LayoutPlaceholder(sp, None)
        return layout_placeholder, _inherited_value_, int_value_

    @pytest.fixture(params=[(True, 42), (False, None)])
    def mstr_val_fixture(
            self, request, _master_placeholder_, master_placeholder_):
        has_master_placeholder, expected_value = request.param
        layout_placeholder = _LayoutPlaceholder(None, None)
        attr_name = 'width'
        if has_master_placeholder:
            setattr(master_placeholder_, attr_name, expected_value)
            _master_placeholder_.return_value = master_placeholder_
        else:
            _master_placeholder_.return_value = None
        return layout_placeholder, attr_name, expected_value

    @pytest.fixture(params=['left', 'top', 'width', 'height'])
    def xfrm_fixture(self, request, _direct_or_inherited_value_, int_value_):
        attr_name = request.param
        layout_placeholder = _LayoutPlaceholder(None, None)
        _direct_or_inherited_value_.return_value = int_value_
        return (
            layout_placeholder, _direct_or_inherited_value_, attr_name,
            int_value_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _direct_or_inherited_value_(self, request):
        return method_mock(
            request, _LayoutPlaceholder, '_direct_or_inherited_value'
        )

    @pytest.fixture
    def _inherited_value_(self, request, int_value_):
        return method_mock(
            request, _LayoutPlaceholder, '_inherited_value',
            return_value=int_value_
        )

    @pytest.fixture
    def int_value_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def _master_placeholder_(self, request):
        return property_mock(
            request, _LayoutPlaceholder, '_master_placeholder'
        )

    @pytest.fixture
    def master_placeholder_(self, request):
        return instance_mock(request, _MasterPlaceholder)

    @pytest.fixture
    def sp(self, width):
        return (
            an_sp().with_nsdecls('p', 'a').with_child(
                an_spPr().with_child(
                    an_xfrm().with_child(
                        an_ext().with_cx(width))))
        ).element

    @pytest.fixture
    def width(self):
        return 31416
