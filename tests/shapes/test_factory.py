# encoding: utf-8

"""
Test suite for pptx.shapes.shapetree module
"""

from __future__ import absolute_import

import pytest

from pptx.parts.slide import Slide
from pptx.shapes.autoshape import Shape
from pptx.shapes.base import BaseShape
from pptx.shapes.factory import (
    BaseShapeFactory, _SlidePlaceholderFactory, SlideShapeFactory
)
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.placeholder import _BaseSlidePlaceholder
from pptx.shapes.shapetree import BaseShapeTree

from ..oxml.unitdata.shape import a_cNvPr, an_nvSpPr, an_spTree
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

    def it_constructs_the_right_type_of_shape(self, factory_fixture):
        element, parent_, Constructor_, shape_ = factory_fixture
        shape = SlideShapeFactory(element, parent_)
        Constructor_.assert_called_once_with(element, parent_)
        assert shape is shape_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('p:sp/p:nvSpPr/p:nvPr/p:ph{type=title}',      'spf'),
        ('p:pic/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=1}', 'spf'),
        ('p:sp',                                       'bsf'),
        ('p:pic',                                      'bsf'),
    ])
    def factory_fixture(self, request, _SlidePlaceholderFactory_,
                        placeholder_, BaseShapeFactory_, base_shape_):
        shape_cxml, shape_type = request.param
        shape_elm = element(shape_cxml)
        Constructor_, shape_ = {
            'spf': (_SlidePlaceholderFactory_, placeholder_),
            'bsf': (BaseShapeFactory_,         base_shape_),
        }[shape_type]
        return shape_elm, 42, Constructor_, shape_

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
    def placeholder_(self, request):
        return instance_mock(request, _BaseSlidePlaceholder)

    @pytest.fixture
    def _SlidePlaceholderFactory_(self, request, placeholder_):
        return function_mock(
            request, 'pptx.shapes.factory._SlidePlaceholderFactory',
            return_value=placeholder_
        )


class Describe_SlidePlaceholderFactory(object):

    def it_constructs_the_right_type_of_placeholder(self, factory_fixture):
        element, parent_, Constructor_, placeholder_ = factory_fixture
        placeholder = _SlidePlaceholderFactory(element, parent_)
        Constructor_.assert_called_once_with(element, parent_)
        assert placeholder is placeholder_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('p:sp/p:nvSpPr/p:nvPr/p:ph{type=title}',                 'sp'),
        ('p:sp/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=1}',             'pph'),
        ('p:sp/p:nvSpPr/p:nvPr/p:ph{type=clipArt,idx=1}',         'pph'),
        ('p:sp/p:nvSpPr/p:nvPr/p:ph{type=tbl,idx=1}',             'tph'),
        ('p:sp/p:nvSpPr/p:nvPr/p:ph{type=chart,idx=10}',          'cph'),
        ('p:pic/p:nvPicPr/p:nvPr/p:ph{type=pic,idx=1}',           'php'),
        ('p:graphicFrame/p:nvSpPr/p:nvPr/p:ph{type=tbl,idx=2}',   'phgf'),
        ('p:graphicFrame/p:nvSpPr/p:nvPr/p:ph{type=chart,idx=2}', 'phgf'),
        ('p:graphicFrame/p:nvSpPr/p:nvPr/p:ph{type=dgm,idx=2}',   'phgf'),
    ])
    def factory_fixture(
            self, request, SlidePlaceholder_, ChartPlaceholder_,
            PicturePlaceholder_, TablePlaceholder_, PlaceholderGraphicFrame_,
            PlaceholderPicture_, placeholder_):
        shape_cxml, constructor_key = request.param
        shape_elm = element(shape_cxml)
        Constructor_, shape_ = {
            'sp':   (SlidePlaceholder_,        placeholder_),
            'cph':  (ChartPlaceholder_,        placeholder_),
            'pph':  (PicturePlaceholder_,      placeholder_),
            'tph':  (TablePlaceholder_,        placeholder_),
            'phgf': (PlaceholderGraphicFrame_, placeholder_),
            'php':  (PlaceholderPicture_,      placeholder_),
        }[constructor_key]
        return shape_elm, 42, Constructor_, shape_

    # fixture components -----------------------------------

    @pytest.fixture
    def ChartPlaceholder_(self, request, placeholder_):
        return class_mock(
            request, 'pptx.shapes.factory.ChartPlaceholder',
            return_value=placeholder_
        )

    @pytest.fixture
    def PicturePlaceholder_(self, request, placeholder_):
        return class_mock(
            request, 'pptx.shapes.factory.PicturePlaceholder',
            return_value=placeholder_
        )

    @pytest.fixture
    def PlaceholderGraphicFrame_(self, request, placeholder_):
        return class_mock(
            request, 'pptx.shapes.factory.PlaceholderGraphicFrame',
            return_value=placeholder_
        )

    @pytest.fixture
    def PlaceholderPicture_(self, request, placeholder_):
        return class_mock(
            request, 'pptx.shapes.factory.PlaceholderPicture',
            return_value=placeholder_
        )

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, _BaseSlidePlaceholder)

    @pytest.fixture
    def SlidePlaceholder_(self, request, placeholder_):
        return class_mock(
            request, 'pptx.shapes.factory.SlidePlaceholder',
            return_value=placeholder_
        )

    @pytest.fixture
    def TablePlaceholder_(self, request, placeholder_):
        return class_mock(
            request, 'pptx.shapes.factory.TablePlaceholder',
            return_value=placeholder_
        )
