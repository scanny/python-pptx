# encoding: utf-8

"""
Test suite for pptx.shapes.shapetree module
"""

from __future__ import absolute_import

import pytest

from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.parts.slide import Slide
from pptx.shapes.autoshape import Shape
from pptx.shapes.shape import BaseShape
from pptx.shapes.picture import Picture
from pptx.shapes.table import Table
from pptx.shapes.shapetree import BaseShapeTree, BaseShapeFactory

from ..oxml.unitdata.shape import (
    a_cNvPr, a_graphic, a_graphicData, a_graphicFrame, a_grpSp, a_pic,
    an_nvSpPr, an_sp, an_spPr, an_spTree
)
from ..oxml.unitdata.slides import a_sld, a_cSld
from ..unitutil.mock import (
    call, class_mock, function_mock, instance_mock, method_mock
)


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

    @pytest.fixture(params=['sp', 'pic', 'tbl', 'chart', 'grpSp'])
    def factory_fixture(
            self, request, slide_, Shape_, shape_, Picture_, picture_,
            tbl_bldr, Table_, table_, chart_bldr, BaseShape_, base_shape_):
        shape_bldr, ShapeClass_, shape_mock = {
            'sp':    (an_sp(),    Shape_,     shape_),
            'pic':   (a_pic(),    Picture_,   picture_),
            'tbl':   (tbl_bldr,   Table_,     table_),
            'chart': (chart_bldr, BaseShape_, base_shape_),
            'grpSp': (a_grpSp(),  BaseShape_, base_shape_),
        }[request.param]
        shape_elm = shape_bldr.with_nsdecls().element
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
            request, 'pptx.shapes.shapetree.BaseShape',
            return_value=base_shape_
        )

    @pytest.fixture
    def base_shape_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def chart_bldr(self):
        chart_uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        return (
            a_graphicFrame().with_child(
                a_graphic().with_child(
                    a_graphicData().with_uri(chart_uri)))
        )

    @pytest.fixture
    def Picture_(self, request, picture_):
        return class_mock(
            request, 'pptx.shapes.shapetree.Picture', return_value=picture_
        )

    @pytest.fixture
    def picture_(self, request):
        return instance_mock(request, Picture)

    @pytest.fixture
    def Shape_(self, request, shape_):
        return class_mock(
            request, 'pptx.shapes.shapetree.Shape', return_value=shape_
        )

    @pytest.fixture
    def shape_(self, request):
        return instance_mock(request, Shape)

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def Table_(self, request, table_):
        return class_mock(
            request, 'pptx.shapes.shapetree.Table', return_value=table_
        )

    @pytest.fixture
    def table_(self, request):
        return instance_mock(request, Table)

    @pytest.fixture
    def tbl_bldr(self):
        tbl_uri = 'http://schemas.openxmlformats.org/drawingml/2006/table'
        return (
            a_graphicFrame().with_child(
                a_graphic().with_child(
                    a_graphicData().with_uri(tbl_uri)))
        )


class DescribeBaseShapeTree(object):

    def it_knows_how_many_shapes_it_contains(self, len_fixture):
        shapes, expected_count = len_fixture
        shape_count = len(shapes)
        assert shape_count == expected_count

    def it_can_iterate_over_the_shapes_it_contains(self, iter_fixture):
        shapes, BaseShapeFactory_, sp_, sp_2_, shape_, shape_2_ = (
            iter_fixture
        )
        iter_vals = [s for s in shapes]
        assert BaseShapeFactory_.call_args_list == [
            call(sp_, shapes),
            call(sp_2_, shapes)
        ]
        assert iter_vals == [shape_, shape_2_]

    def it_iterates_over_spTree_shape_elements_to_help__iter__(
            self, iter_elms_fixture):
        shapes, expected_elm_count = iter_elms_fixture
        shape_elms = [elm for elm in shapes._iter_member_elms()]
        assert len(shape_elms) == expected_elm_count
        for elm in shape_elms:
            assert isinstance(elm, CT_Shape)

    def it_supports_indexed_access(self, getitem_fixture):
        shapes, idx, BaseShapeFactory_, shape_elm_, shape_ = getitem_fixture
        shape = shapes[idx]
        BaseShapeFactory_.assert_called_once_with(shape_elm_, shapes)
        assert shape is shape_

    def it_raises_on_shape_index_out_of_range(self, getitem_fixture):
        shapes = getitem_fixture[0]
        with pytest.raises(IndexError):
            shapes[2]

    def it_knows_the_part_it_belongs_to(self, slide):
        shapes = BaseShapeTree(slide)
        assert shapes.part is slide

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getitem_fixture(
            self, _iter_member_elms_, BaseShapeFactory_, sp_2_, shape_):
        shapes = BaseShapeTree(None)
        idx = 1
        return shapes, idx, BaseShapeFactory_, sp_2_, shape_

    @pytest.fixture
    def iter_fixture(
            self, _iter_member_elms_, BaseShapeFactory_, sp_, sp_2_,
            shape_, shape_2_):
        shapes = BaseShapeTree(None)
        return shapes, BaseShapeFactory_, sp_, sp_2_, shape_, shape_2_

    @pytest.fixture
    def iter_elms_fixture(self, slide):
        shapes = BaseShapeTree(slide)
        expected_elm_count = 2
        return shapes, expected_elm_count

    @pytest.fixture
    def len_fixture(self, slide):
        shapes = BaseShapeTree(slide)
        expected_count = 2
        return shapes, expected_count

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _iter_member_elms_(self, request, sp_, sp_2_):
        return method_mock(
            request, BaseShapeTree, '_iter_member_elms',
            return_value=iter([sp_, sp_2_])
        )

    @pytest.fixture
    def BaseShapeFactory_(self, request, shape_, shape_2_):
        return function_mock(
            request, 'pptx.shapes.shapetree.BaseShapeFactory',
            side_effect=[shape_, shape_2_]
        )

    @pytest.fixture
    def shape_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def shape_2_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def sld(self):
        sld_bldr = (
            a_sld().with_nsdecls().with_child(
                a_cSld().with_child(
                    an_spTree().with_child(
                        an_spPr()).with_child(
                        an_sp()).with_child(
                        an_sp())))
        )
        return sld_bldr.element

    @pytest.fixture
    def slide(self, sld):
        return Slide(None, None, sld, None)

    @pytest.fixture
    def sp_(self, request):
        return instance_mock(request, CT_Shape)

    @pytest.fixture
    def sp_2_(self, request):
        return instance_mock(request, CT_Shape)
