# encoding: utf-8

"""
Test suite for pptx.shapes.shapetree module
"""

from __future__ import absolute_import

import pytest

from pptx.parts.slide import Slide
from pptx.shapes.autoshape import Shape
from pptx.shapes.base import BaseShape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.placeholder import SlidePlaceholder
from pptx.shapes.shapetree import (
    BaseShapeTree, BaseShapeFactory, SlideShapeFactory
)

from ..oxml.unitdata.shape import (
    a_cNvPr, a_ph, a_pic, an_nvPr, an_nvSpPr, an_sp, an_spTree
)
from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, function_mock, instance_mock


class DescribeBaseShapeFactory(object):

    def it_constructs_the_appropriate_shape_instance_for_a_shape_element(
            self, factory_fixture):
        shape_elm, parent_, ShapeClass_, shape_ = factory_fixture
        shape = BaseShapeFactory(shape_elm, parent_)
        ShapeClass_.assert_called_once_with(shape_elm, parent_)
        assert shape is shape_

    def it_finds_an_unused_shape_id_to_help_add_shape(self, next_id_fixture):
        shapes, next_available_shape_id = next_id_fixture
        shape_id = shapes._next_shape_id
        assert shape_id == next_available_shape_id

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'])
    def factory_fixture(
            self, request, slide_, Shape_, shape_, Picture_, picture_,
            GraphicFrame_, graphic_frame_, BaseShape_, base_shape_):
        shape_cxml, ShapeClass_, shape_mock = {
            'sp':           ('p:sp',           Shape_,        shape_),
            'pic':          ('p:pic',          Picture_,      picture_),
            'graphicFrame': ('p:graphicFrame', GraphicFrame_, graphic_frame_),
            'grpSp':        ('p:grpSp',        BaseShape_,    base_shape_),
            'cxnSp':        ('p:cxnSp',        BaseShape_,    base_shape_),
        }[request.param]
        shape_elm = element(shape_cxml)
        return shape_elm, slide_, ShapeClass_, shape_mock

    @pytest.fixture(params=[
        ((), 1), ((0,), 1), ((1,), 2), ((2,), 1), ((1, 3,), 2),
        (('foobar', 0, 1, 7), 2), (('1foo', 2, 2, 2), 1), ((1, 1, 1, 4), 2),
    ])
    def next_id_fixture(self, request, slide_):
        used_ids, next_available_shape_id = request.param
        nvSpPr_bldr = an_nvSpPr()
        for used_id in used_ids:
            nvSpPr_bldr.with_child(a_cNvPr().with_id(used_id))
        spTree = an_spTree().with_nsdecls().with_child(nvSpPr_bldr).element
        print(spTree.xml)
        slide_.spTree = spTree
        shapes = BaseShapeTree(slide_)
        return shapes, next_available_shape_id

    # fixture components -----------------------------------

    @pytest.fixture
    def BaseShape_(self, request, base_shape_):
        return class_mock(
            request, 'pptx.shapes.factory.BaseShape',
            return_value=base_shape_
        )

    @pytest.fixture
    def base_shape_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def GraphicFrame_(self, request, graphic_frame_):
        return class_mock(
            request, 'pptx.shapes.factory.GraphicFrame',
            return_value=graphic_frame_
        )

    @pytest.fixture
    def graphic_frame_(self, request):
        return instance_mock(request, GraphicFrame)

    @pytest.fixture
    def Picture_(self, request, picture_):
        return class_mock(
            request, 'pptx.shapes.factory.Picture', return_value=picture_
        )

    @pytest.fixture
    def picture_(self, request):
        return instance_mock(request, Picture)

    @pytest.fixture
    def Shape_(self, request, shape_):
        return class_mock(
            request, 'pptx.shapes.factory.Shape', return_value=shape_
        )

    @pytest.fixture
    def shape_(self, request):
        return instance_mock(request, Shape)

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)


class DescribeSlideShapeFactory(object):

    def it_constructs_a_slide_placeholder_for_a_shape_element(
            self, factory_fixture):
        shape_elm, parent_, ShapeConstructor_, shape_ = factory_fixture
        shape = SlideShapeFactory(shape_elm, parent_)
        ShapeConstructor_.assert_called_once_with(shape_elm, parent_)
        assert shape is shape_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=['ph', 'sp', 'pic'])
    def factory_fixture(
            self, request, ph_bldr, slide_, SlidePlaceholder_,
            slide_placeholder_, BaseShapeFactory_, base_shape_):
        shape_bldr, ShapeConstructor_, shape_mock = {
            'ph':  (ph_bldr, SlidePlaceholder_, slide_placeholder_),
            'sp':  (an_sp(), BaseShapeFactory_, base_shape_),
            'pic': (a_pic(), BaseShapeFactory_, base_shape_),
        }[request.param]
        shape_elm = shape_bldr.with_nsdecls().element
        return shape_elm, slide_, ShapeConstructor_, shape_mock

    # fixture components -----------------------------------

    @pytest.fixture
    def BaseShapeFactory_(self, request, base_shape_):
        return function_mock(
            request, 'pptx.shapes.factory.BaseShapeFactory',
            return_value=base_shape_
        )

    @pytest.fixture
    def base_shape_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def SlidePlaceholder_(self, request, slide_placeholder_):
        return class_mock(
            request, 'pptx.shapes.factory.SlidePlaceholder',
            return_value=slide_placeholder_
        )

    @pytest.fixture
    def slide_placeholder_(self, request):
        return instance_mock(request, SlidePlaceholder)

    @pytest.fixture
    def ph_bldr(self):
        return (
            an_sp().with_child(
                an_nvSpPr().with_child(
                    an_nvPr().with_child(
                        a_ph().with_idx(1))))
        )

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)
