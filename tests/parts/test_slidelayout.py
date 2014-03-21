# encoding: utf-8

"""
Test suite for pptx.parts.slidelayout module
"""

from __future__ import absolute_import

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.autoshape import CT_Shape
from pptx.parts.slidelayout import (
    _LayoutPlaceholder, _LayoutShapeTree, SlideLayout
)
from pptx.parts.slidemaster import SlideMaster

from ..unitutil import class_mock, function_mock, instance_mock, method_mock


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

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def master_fixture(self, slide_master_, part_related_by_):
        slide_layout = SlideLayout(None, None, None, None)
        return slide_layout, slide_master_

    @pytest.fixture
    def shapes_fixture(self, _LayoutShapeTree_, layout_shape_tree_):
        slide_layout = SlideLayout(None, None, None, None)
        return slide_layout, _LayoutShapeTree_, layout_shape_tree_

    # fixture components -----------------------------------

    @pytest.fixture
    def _LayoutShapeTree_(self, request, layout_shape_tree_):
        return class_mock(
            request, 'pptx.parts.slidelayout._LayoutShapeTree',
            return_value=layout_shape_tree_
        )

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
