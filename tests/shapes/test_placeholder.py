# encoding: utf-8

"""
Test suite for pptx.shapes.placeholder module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.shapes.shared import (
    BaseShapeElement, ST_Direction, ST_PlaceholderSize
)
from pptx.parts.slide import Slide, _SlideShapeTree
from pptx.parts.slidelayout import SlideLayout
from pptx.parts.slidemaster import SlideMaster
from pptx.shapes.placeholder import (
    BasePlaceholder, BasePlaceholders, LayoutPlaceholder, MasterPlaceholder,
    SlidePlaceholder
)
from pptx.shapes.shapetree import BaseShapeTree

from ..oxml.unitdata.shape import (
    an_ext, a_graphicFrame, a_ph, an_nvGraphicFramePr, an_nvPicPr, an_nvPr,
    an_nvSpPr, an_sp, an_spPr, an_xfrm
)
from ..unitutil.mock import instance_mock, method_mock, property_mock


class DescribeBasePlaceholders(object):

    def it_contains_only_placeholder_shapes(self, member_fixture):
        shape_elm_, is_ph_shape = member_fixture
        _is_ph_shape = BasePlaceholders._is_member_elm(shape_elm_)
        assert _is_ph_shape == is_ph_shape

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[True, False])
    def member_fixture(self, request, shape_elm_):
        is_ph_shape = request.param
        shape_elm_.has_ph_elm = is_ph_shape
        return shape_elm_, is_ph_shape

    # fixture components ---------------------------------------------

    @pytest.fixture
    def shape_elm_(self, request):
        return instance_mock(request, BaseShapeElement)


class DescribeBasePlaceholder(object):

    def it_knows_its_idx_value(self, idx_fixture):
        placeholder, idx = idx_fixture
        assert placeholder.idx == idx

    def it_knows_its_orient_value(self, orient_fixture):
        placeholder, orient = orient_fixture
        assert placeholder.orient == orient

    def it_raises_on_ph_orient_when_not_a_placeholder(
            self, orient_raise_fixture):
        shape = orient_raise_fixture
        with pytest.raises(ValueError):
            shape.orient

    def it_knows_its_sz_value(self, sz_fixture):
        placeholder, sz = sz_fixture
        assert placeholder.sz == sz

    def it_knows_its_placeholder_type(self, type_fixture):
        placeholder, ph_type = type_fixture
        assert placeholder.ph_type == ph_type

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('sp',           None, 'title', 0),
        ('sp',           0,    'title', 0),
        ('sp',           3,    'body',  3),
        ('pic',          6,    'pic',   6),
        ('graphicFrame', 9,    'tbl',   9),
    ])
    def idx_fixture(self, request):
        tagname, idx, ph_type, expected_idx = request.param
        shape_elm = self.shape_elm_factory(tagname, ph_type, idx)
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_idx

    @pytest.fixture(params=[
        (None, ST_Direction.HORZ),
        (ST_Direction.VERT, ST_Direction.VERT),
    ])
    def orient_fixture(self, request):
        orient, expected_orient = request.param
        ph_bldr = a_ph()
        if orient is not None:
            ph_bldr.with_orient(orient)
        shape_elm = (
            an_sp().with_nsdecls('p').with_child(
                an_nvSpPr().with_child(
                    an_nvPr().with_child(
                        ph_bldr)))
        ).element
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_orient

    @pytest.fixture
    def orient_raise_fixture(self):
        shape_elm = (
            an_sp().with_nsdecls('p').with_child(
                an_nvSpPr().with_child(
                    an_nvPr()))
        ).element
        shape = BasePlaceholder(shape_elm, None)
        return shape

    @pytest.fixture(params=[
        (None, ST_PlaceholderSize.FULL),
        (ST_PlaceholderSize.HALF, ST_PlaceholderSize.HALF),
    ])
    def sz_fixture(self, request):
        sz, expected_sz = request.param
        ph_bldr = a_ph()
        if sz is not None:
            ph_bldr.with_sz(sz)
        shape_elm = (
            an_sp().with_nsdecls('p').with_child(
                an_nvSpPr().with_child(
                    an_nvPr().with_child(
                        ph_bldr)))
        ).element
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_sz

    @pytest.fixture(params=[
        ('sp',           None,    1, PP_PLACEHOLDER.OBJECT),
        ('sp',           'title', 0, PP_PLACEHOLDER.TITLE),
        ('pic',          'pic',   6, PP_PLACEHOLDER.PICTURE),
        ('graphicFrame', 'tbl',   9, PP_PLACEHOLDER.TABLE),
    ])
    def type_fixture(self, request):
        tagname, ph_type, idx, expected_ph_type = request.param
        shape_elm = self.shape_elm_factory(tagname, ph_type, idx)
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_ph_type

    # fixture components ---------------------------------------------

    @staticmethod
    def shape_elm_factory(tagname, ph_type, idx):
        root_bldr, nvXxPr_bldr = {
            'sp':           (an_sp().with_nsdecls('p'), an_nvSpPr()),
            'pic':          (an_sp().with_nsdecls('p'), an_nvPicPr()),
            'graphicFrame': (a_graphicFrame().with_nsdecls('p'),
                             an_nvGraphicFramePr()),
        }[tagname]
        ph_bldr = {
            None:    a_ph().with_idx(idx),
            'obj':   a_ph().with_idx(idx),
            'title': a_ph().with_type('title'),
            'body':  a_ph().with_type('body').with_idx(idx),
            'pic':   a_ph().with_type('pic').with_idx(idx),
            'tbl':   a_ph().with_type('tbl').with_idx(idx),
        }[ph_type]
        return (
            root_bldr.with_child(
                nvXxPr_bldr.with_child(
                    an_nvPr().with_child(
                        ph_bldr)))
        ).element


class DescribeLayoutPlaceholder(object):

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

    def it_finds_its_corresponding_master_placeholder_to_help_inherit(
            self, mstr_ph_fixture):
        layout_placeholder, master_, mstr_ph_type, master_placeholder_ = (
            mstr_ph_fixture
        )
        master_placeholder = layout_placeholder._master_placeholder
        master_.placeholders.get.assert_called_once_with(mstr_ph_type, None)
        assert master_placeholder is master_placeholder_

    def it_finds_its_slide_master_to_help_inherit(self, slide_master_fixture):
        layout_placeholder, slide_master_ = slide_master_fixture
        slide_master = layout_placeholder._slide_master
        assert slide_master == slide_master_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def direct_fixture(self, sp, width):
        layout_placeholder = LayoutPlaceholder(sp, None)
        return layout_placeholder, width

    @pytest.fixture
    def inherited_fixture(self, sp, _inherited_value_, int_value_):
        layout_placeholder = LayoutPlaceholder(sp, None)
        return layout_placeholder, _inherited_value_, int_value_

    @pytest.fixture(params=[
        (PP_PLACEHOLDER.TABLE,  PP_PLACEHOLDER.BODY),
        (PP_PLACEHOLDER.BODY,   PP_PLACEHOLDER.BODY),
        (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY),
    ])
    def mstr_ph_fixture(
            self, request, ph_type_, _slide_master_, slide_master_,
            master_placeholder_):
        layout_placeholder = LayoutPlaceholder(None, None)
        ph_type, mstr_ph_type = request.param
        ph_type_.return_value = ph_type
        slide_master_.placeholders.get.return_value = master_placeholder_
        return (
            layout_placeholder, slide_master_, mstr_ph_type,
            master_placeholder_
        )

    @pytest.fixture(params=[(True, 42), (False, None)])
    def mstr_val_fixture(
            self, request, _master_placeholder_, master_placeholder_):
        has_master_placeholder, expected_value = request.param
        layout_placeholder = LayoutPlaceholder(None, None)
        attr_name = 'width'
        if has_master_placeholder:
            setattr(master_placeholder_, attr_name, expected_value)
            _master_placeholder_.return_value = master_placeholder_
        else:
            _master_placeholder_.return_value = None
        return layout_placeholder, attr_name, expected_value

    @pytest.fixture
    def slide_master_fixture(self, parent_, slide_master_):
        layout_placeholder = LayoutPlaceholder(None, parent_)
        return layout_placeholder, slide_master_

    @pytest.fixture(params=['left', 'top', 'width', 'height'])
    def xfrm_fixture(self, request, _direct_or_inherited_value_, int_value_):
        attr_name = request.param
        layout_placeholder = LayoutPlaceholder(None, None)
        _direct_or_inherited_value_.return_value = int_value_
        return (
            layout_placeholder, _direct_or_inherited_value_, attr_name,
            int_value_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _direct_or_inherited_value_(self, request):
        return method_mock(
            request, LayoutPlaceholder, '_direct_or_inherited_value'
        )

    @pytest.fixture
    def _inherited_value_(self, request, int_value_):
        return method_mock(
            request, LayoutPlaceholder, '_inherited_value',
            return_value=int_value_
        )

    @pytest.fixture
    def int_value_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def _master_placeholder_(self, request):
        return property_mock(
            request, LayoutPlaceholder, '_master_placeholder'
        )

    @pytest.fixture
    def master_placeholder_(self, request):
        return instance_mock(request, MasterPlaceholder)

    @pytest.fixture
    def parent_(self, request, slide_layout_):
        parent_ = instance_mock(request, BaseShapeTree)
        parent_.part = slide_layout_
        return parent_

    @pytest.fixture
    def ph_type_(self, request):
        return property_mock(request, LayoutPlaceholder, 'ph_type')

    @pytest.fixture
    def slide_layout_(self, request, slide_master_):
        slide_layout_ = instance_mock(request, SlideLayout)
        slide_layout_.slide_master = slide_master_
        return slide_layout_

    @pytest.fixture
    def _slide_master_(self, request, slide_master_):
        return property_mock(
            request, LayoutPlaceholder, '_slide_master',
            return_value=slide_master_
        )

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

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


class DescribeSlidePlaceholder(object):

    def it_considers_inheritance_when_computing_pos_and_size(
            self, xfrm_fixture):
        slide_placeholder, _direct_or_inherited_value_ = xfrm_fixture[:2]
        attr_name, expected_value = xfrm_fixture[2:]
        value = getattr(slide_placeholder, attr_name)
        _direct_or_inherited_value_.assert_called_once_with(attr_name)
        assert value == expected_value

    def it_provides_direct_property_values_when_they_exist(
            self, direct_fixture):
        slide_placeholder, expected_width = direct_fixture
        width = slide_placeholder.width
        assert width == expected_width

    def it_provides_inherited_property_values_when_no_direct_value(
            self, inherited_fixture):
        slide_placeholder, _inherited_value_, inherited_left_ = (
            inherited_fixture
        )
        left = slide_placeholder.left
        _inherited_value_.assert_called_once_with('left')
        assert left == inherited_left_

    def it_knows_how_to_get_a_property_value_from_its_layout(
            self, layout_val_fixture):
        slide_placeholder, attr_name, expected_value = layout_val_fixture
        value = slide_placeholder._inherited_value(attr_name)
        assert value == expected_value

    def it_finds_its_corresponding_layout_placeholder_to_help_inherit(
            self, layout_ph_fixture):
        slide_placeholder, layout_, idx, layout_placeholder_ = (
            layout_ph_fixture
        )
        layout_placeholder = slide_placeholder._layout_placeholder
        layout_.placeholders.get.assert_called_once_with(idx=idx)
        assert layout_placeholder is layout_placeholder_

    def it_finds_its_slide_layout_to_help_inherit(
            self, slide_layout_fixture):
        slide_placeholder, slide_layout_ = slide_layout_fixture
        slide_layout = slide_placeholder._slide_layout
        assert slide_layout == slide_layout_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def direct_fixture(self, sp, width):
        slide_placeholder = SlidePlaceholder(sp, None)
        return slide_placeholder, width

    @pytest.fixture
    def inherited_fixture(self, sp, _inherited_value_, int_value_):
        slide_placeholder = SlidePlaceholder(sp, None)
        return slide_placeholder, _inherited_value_, int_value_

    @pytest.fixture
    def layout_ph_fixture(
            self, request, idx_, int_value_, _slide_layout_, slide_layout_,
            layout_placeholder_):
        slide_placeholder = SlidePlaceholder(None, None)
        idx_.return_value = int_value_
        slide_layout_.placeholders.get.return_value = layout_placeholder_
        return (
            slide_placeholder, slide_layout_, int_value_, layout_placeholder_
        )

    @pytest.fixture(params=[(True, 42), (False, None)])
    def layout_val_fixture(
            self, request, _layout_placeholder_, layout_placeholder_):
        has_layout_placeholder, expected_value = request.param
        slide_placeholder = SlidePlaceholder(None, None)
        attr_name = 'width'
        if has_layout_placeholder:
            setattr(layout_placeholder_, attr_name, expected_value)
            _layout_placeholder_.return_value = layout_placeholder_
        else:
            _layout_placeholder_.return_value = None
        return slide_placeholder, attr_name, expected_value

    @pytest.fixture
    def slide_layout_fixture(self, parent_, slide_layout_):
        slide_placeholder = SlidePlaceholder(None, parent_)
        return slide_placeholder, slide_layout_

    @pytest.fixture(params=['left', 'top', 'width', 'height'])
    def xfrm_fixture(self, request, _effective_value_, int_value_):
        attr_name = request.param
        slide_placeholder = SlidePlaceholder(None, None)
        _effective_value_.return_value = int_value_
        return (
            slide_placeholder, _effective_value_, attr_name,
            int_value_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _effective_value_(self, request):
        return method_mock(
            request, SlidePlaceholder, '_effective_value'
        )

    @pytest.fixture
    def idx_(self, request):
        return property_mock(request, SlidePlaceholder, 'idx')

    @pytest.fixture
    def _inherited_value_(self, request, int_value_):
        return method_mock(
            request, SlidePlaceholder, '_inherited_value',
            return_value=int_value_
        )

    @pytest.fixture
    def int_value_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def _layout_placeholder_(self, request):
        return property_mock(
            request, SlidePlaceholder, '_layout_placeholder'
        )

    @pytest.fixture
    def layout_placeholder_(self, request):
        return instance_mock(request, LayoutPlaceholder)

    @pytest.fixture
    def parent_(self, request, slide_):
        parent_ = instance_mock(request, _SlideShapeTree)
        parent_.part = slide_
        return parent_

    @pytest.fixture
    def slide_(self, request, slide_layout_):
        slide_ = instance_mock(request, Slide)
        slide_.slide_layout = slide_layout_
        return slide_

    @pytest.fixture
    def _slide_layout_(self, request, slide_layout_):
        return property_mock(
            request, SlidePlaceholder, '_slide_layout',
            return_value=slide_layout_
        )

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

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
