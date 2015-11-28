# encoding: utf-8

"""
Test suite for pptx.shapes.shapetree module
"""

from __future__ import absolute_import

import pytest

from pptx.chart.data import ChartData
from pptx.enum.base import EnumValue
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from pptx.oxml.shapes.groupshape import CT_GroupShape
from pptx.oxml.shapes.picture import CT_Picture
from pptx.oxml.shapes.shared import BaseShapeElement, ST_Direction
from pptx.parts.image import ImagePart
from pptx.parts.slide import Slide
from pptx.parts.slidelayout import SlideLayout
from pptx.shapes.autoshape import AutoShapeType, Shape
from pptx.shapes.base import BaseShape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.placeholder import (
    BasePlaceholder, LayoutPlaceholder, SlidePlaceholder
)
from pptx.shapes.shapetree import (
    BasePlaceholders, BaseShapeTree, SlideShapeTree
)
from pptx.shapes.table import Table

from ..oxml.unitdata.shape import a_cNvPr, an_sp, an_spPr, an_spTree
from ..oxml.unitdata.slides import a_sld, a_cSld
from ..unitutil.cxml import element
from ..unitutil.mock import (
    call, class_mock, function_mock, instance_mock, method_mock,
    property_mock
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


class DescribeSlideShapeTree(object):

    def it_can_add_a_chart(self, add_chart_fixture):
        shape_tree, chart_type_, x, y, cx, cy = add_chart_fixture[:6]
        chart_data_, rId_, graphic_frame_ = add_chart_fixture[6:]

        graphic_frame = shape_tree.add_chart(
            chart_type_, x, y, cx, cy, chart_data_
        )

        shape_tree.part.add_chart_part.assert_called_once_with(
            chart_type_, chart_data_
        )
        shape_tree._add_chart_graphic_frame.assert_called_once_with(
            shape_tree, rId_, x, y, cx, cy
        )
        assert graphic_frame is graphic_frame_

    def it_constructs_a_slide_placeholder_for_a_placeholder_shape(
            self, factory_fixture):
        shapes, ph_elm_, SlideShapeFactory_, slide_placeholder_ = (
            factory_fixture
        )
        slide_placeholder = shapes._shape_factory(ph_elm_)
        SlideShapeFactory_.assert_called_once_with(ph_elm_, shapes)
        assert slide_placeholder is slide_placeholder_

    def it_can_find_the_title_placeholder(self, title_fixture):
        shapes, _shape_factory_, calls, expected_value = title_fixture

        title_placeholder = shapes.title

        assert _shape_factory_.call_args_list == calls
        assert title_placeholder == expected_value

    def it_can_add_an_autoshape(self, autoshape_fixture):
        # fixture ----------------------
        shapes, autoshape_type_id_, x_, y_, cx_, cy_ = autoshape_fixture[:6]
        AutoShapeType_, _add_sp_from_autoshape_type_ = autoshape_fixture[6:8]
        autoshape_type_, _shape_factory_, sp_ = autoshape_fixture[8:11]
        shape_ = autoshape_fixture[11]
        # exercise ---------------------
        shape = shapes.add_shape(autoshape_type_id_, x_, y_, cx_, cy_)
        # verify -----------------------
        AutoShapeType_.assert_called_once_with(autoshape_type_id_)
        _add_sp_from_autoshape_type_.assert_called_once_with(
            autoshape_type_, x_, y_, cx_, cy_
        )
        _shape_factory_.assert_called_once_with(sp_)
        assert shape is shape_

    def it_can_add_a_picture_shape(self, picture_fixture):
        shapes, image_file_, x_, y_, cx_, cy_ = picture_fixture[:6]
        image_part_, rId_, pic_, picture_ = picture_fixture[6:]

        picture = shapes.add_picture(image_file_, x_, y_, cx_, cy_)

        shapes._slide.get_or_add_image_part.assert_called_once_with(
            image_file_
        )
        shapes._add_pic_from_image_part.assert_called_once_with(
            image_part_, rId_, x_, y_, cx_, cy_
        )
        shapes._shape_factory.assert_called_once_with(pic_)
        assert picture is picture_

    def it_can_add_a_table(self, table_fixture):
        # fixture ----------------------
        shapes, rows_, cols_, x_, y_, cx_, cy_ = table_fixture[:7]
        _add_graphicFrame_containing_table_ = table_fixture[7]
        _shape_factory_, graphicFrame_, table_ = table_fixture[8:]
        # exercise ---------------------
        table = shapes.add_table(rows_, cols_, x_, y_, cx_, cy_)
        # verify -----------------------
        _add_graphicFrame_containing_table_.assert_called_once_with(
            rows_, cols_, x_, y_, cx_, cy_
        )
        _shape_factory_.assert_called_once_with(graphicFrame_)
        assert table is table_

    def it_can_add_a_textbox(self, textbox_fixture):
        shapes, x_, y_, cx_, cy_, _add_textbox_sp_ = textbox_fixture[:6]
        _shape_factory_, sp_, textbox_ = textbox_fixture[6:]
        textbox = shapes.add_textbox(x_, y_, cx_, cy_)
        _add_textbox_sp_.assert_called_once_with(x_, y_, cx_, cy_)
        _shape_factory_.assert_called_once_with(sp_)
        assert textbox is textbox_

    def it_can_clone_placeholder_shapes_from_a_layout(self, clone_fixture):
        shapes, slide_layout_, placeholder_, _clone_layout_placeholder_ = (
            clone_fixture
        )
        shapes.clone_layout_placeholders(slide_layout_)
        _clone_layout_placeholder_.assert_called_once_with(placeholder_)

    def it_knows_the_index_of_each_shape(self, index_fixture):
        shapes, shape_, expected_idx = index_fixture
        idx = shapes.index(shape_)
        assert idx == expected_idx

    def it_raises_on_index_where_shape_not_found(self, index_fixture):
        shapes, shape_, expected_idx = index_fixture
        shapes._spTree.iter_shape_elms.return_value = []
        with pytest.raises(ValueError):
            shapes.index(shape_)

    def it_adds_a_chart_graphic_frame_to_help_add_chart(
            self, add_chart_graph_frm_fixture):
        shape_tree, rId_, x, y, cx, cy = add_chart_graph_frm_fixture[:6]
        _add_chart_graphicFrame_ = add_chart_graph_frm_fixture[6]
        SlideShapeFactory_, graphicFrame_ = add_chart_graph_frm_fixture[7:9]
        graphic_frame_ = add_chart_graph_frm_fixture[9]

        graphic_frame = shape_tree._add_chart_graphic_frame(
            rId_, x, y, cx, cy
        )

        _add_chart_graphicFrame_.assert_called_once_with(
            shape_tree, rId_, x, y, cx, cy
        )
        SlideShapeFactory_.assert_called_once_with(
            graphicFrame_, shape_tree
        )
        assert graphic_frame is graphic_frame_

    def it_adds_a_chart_graphicFrame_to_help_add_chart(
            self, add_chart_graphicFrame_fixture):
        shape_tree, rId, x, y, cx, cy, expected_xml = (
            add_chart_graphicFrame_fixture
        )
        graphicFrame = shape_tree._add_chart_graphicFrame(rId, x, y, cx, cy)
        assert graphicFrame.xml == expected_xml
        shape_tree._spTree.append.assert_called_once_with(graphicFrame)

    def it_adds_a_graphicFrame_to_help_add_table(self, graphicFrame_fixture):
        # fixture ----------------------
        shapes, rows_, cols_, x_, y_, cx_, cy_ = graphicFrame_fixture[:7]
        spTree_, id_, name, graphicFrame_ = graphicFrame_fixture[7:]
        # exercise ---------------------
        graphicFrame = shapes._add_graphicFrame_containing_table(
            rows_, cols_, x_, y_, cx_, cy_
        )
        # verify -----------------------
        spTree_.add_table.assert_called_once_with(
            id_, name, rows_, cols_, x_, y_, cx_, cy_
        )
        assert graphicFrame is graphicFrame_

    def it_adds_a_pic_to_help_add_picture(self, pic_fixture):
        # fixture ----------------------
        shapes, image_part_, rId_, x_, y_, cx_, cy_ = pic_fixture[:7]
        spTree_, id_, name, desc_ = pic_fixture[7:11]
        scaled_cx_, scaled_cy_, pic_ = pic_fixture[11:]
        # exercise ---------------------
        pic = shapes._add_pic_from_image_part(
            image_part_, rId_, x_, y_, cx_, cy_
        )
        # verify -----------------------
        image_part_.scale.assert_called_once_with(cx_, cy_)
        spTree_.add_pic.assert_called_once_with(
            id_, name, desc_, rId_, x_, y_, scaled_cx_, scaled_cy_
        )
        assert pic is pic_

    def it_adds_an_sp_to_help_add_shape(self, sp_fixture):
        # fixture ----------------------
        shapes, autoshape_type_, x_, y_, cx_, cy_ = sp_fixture[:6]
        spTree_, id_, name, prst_, sp_ = sp_fixture[6:]
        # exercise ---------------------
        sp = shapes._add_sp_from_autoshape_type(
            autoshape_type_, x_, y_, cx_, cy_
        )
        # verify -----------------------
        spTree_.add_autoshape.assert_called_once_with(
            id_, name, prst_, x_, y_, cx_, cy_
        )
        assert sp is sp_

    def it_adds_an_sp_to_help_add_textbox(self, textbox_sp_fixture):
        shapes, x_, y_, cx_, cy_, spTree_, id_, name, sp_ = (
            textbox_sp_fixture
        )
        sp = shapes._add_textbox_sp(x_, y_, cx_, cy_)
        spTree_.add_textbox.assert_called_once_with(
            id_, name, x_, y_, cx_, cy_
        )
        assert sp is sp_

    def it_clones_a_placeholder_to_help_clone_placeholders(
            self, clone_ph_fixture):
        shapes, layout_placeholder_, spTree_ = clone_ph_fixture[:3]
        id_, name_, ph_type_, orient_, sz_, idx_ = clone_ph_fixture[3:]
        shapes._clone_layout_placeholder(layout_placeholder_)
        spTree_.add_placeholder.assert_called_once_with(
            id_, name_, ph_type_, orient_, sz_, idx_
        )

    def it_can_find_the_next_placeholder_name_to_help_clone_placeholder(
            self, ph_name_fixture):
        shapes, ph_type, id_, orient, expected_name = ph_name_fixture
        name = shapes._next_ph_name(ph_type, id_, orient)
        print(shapes._spTree.xml)
        assert name == expected_name

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_chart_fixture(
            self, slide_, chart_type_, chart_data_, rId_,
            _add_chart_graphic_frame_, graphic_frame_):
        shape_tree = SlideShapeTree(slide_)
        x, y, cx, cy = 1, 2, 3, 4
        return (
            shape_tree, chart_type_, x, y, cx, cy, chart_data_, rId_,
            graphic_frame_
        )

    @pytest.fixture
    def add_chart_graph_frm_fixture(
            self, rId_, x_, y_, cx_, cy_, _add_chart_graphicFrame_,
            SlideShapeFactory_, graphicFrame_, graphic_frame_):
        shape_tree = SlideShapeTree(None)
        SlideShapeFactory_.return_value = graphic_frame_
        return (
            shape_tree, rId_, x_, y_, cx_, cy_, _add_chart_graphicFrame_,
            SlideShapeFactory_, graphicFrame_, graphic_frame_
        )

    @pytest.fixture
    def add_chart_graphicFrame_fixture(self, slide_, _next_shape_id_, id_):
        rId, x, y, cx, cy = 'rId42', 1, 2, 3, 4
        shape_tree = SlideShapeTree(slide_)
        expected_xml = (
            '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/pre'
            'sentationml/2006/main">\n  <p:nvGraphicFramePr>\n    <p:cNvPr i'
            'd="42" name="Chart 41"/>\n    <p:cNvGraphicFramePr>\n      <a:g'
            'raphicFrameLocks noGrp="1"/>\n    </p:cNvGraphicFramePr>\n    <'
            'p:nvPr/>\n  </p:nvGraphicFramePr>\n  <p:xfrm>\n    <a:off x="1"'
            ' y="2"/>\n    <a:ext cx="3" cy="4"/>\n  </p:xfrm>\n  <a:graphic'
            '>\n    <a:graphicData uri="http://schemas.openxmlformats.org/dr'
            'awingml/2006/chart">\n      <c:chart xmlns:c="http://schemas.op'
            'enxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.'
            'openxmlformats.org/officeDocument/2006/relationships" r:id="rId'
            '42"/>\n    </a:graphicData>\n  </a:graphic>\n</p:graphicFrame>'
            '\n'
        )
        return (
            shape_tree, rId, x, y, cx, cy, expected_xml
        )

    @pytest.fixture
    def autoshape_fixture(
            self, autoshape_type_id_, x_, y_, cx_, cy_, AutoShapeType_,
            _add_sp_from_autoshape_type_, autoshape_type_, _shape_factory_,
            sp_, shape_):
        shapes = SlideShapeTree(None)
        _shape_factory_.return_value = shape_
        return (
            shapes, autoshape_type_id_, x_, y_, cx_, cy_, AutoShapeType_,
            _add_sp_from_autoshape_type_, autoshape_type_, _shape_factory_,
            sp_, shape_
        )

    @pytest.fixture
    def clone_fixture(
            self, slide_layout_, placeholder_, _clone_layout_placeholder_):
        shapes = SlideShapeTree(None)
        slide_layout_.iter_cloneable_placeholders.return_value = (
            iter([placeholder_])
        )
        return (
            shapes, slide_layout_, placeholder_, _clone_layout_placeholder_
        )

    @pytest.fixture
    def clone_ph_fixture(
            self, slide_, layout_placeholder_, spTree_, _next_shape_id_, id_,
            _next_ph_name_, name_, ph_type_, orient_, sz_, idx_):
        shapes = SlideShapeTree(slide_)
        return (
            shapes, layout_placeholder_, spTree_, id_, name_, ph_type_,
            orient_, sz_, idx_
        )

    @pytest.fixture
    def factory_fixture(
            self, ph_elm_, SlideShapeFactory_, slide_placeholder_):
        shapes = SlideShapeTree(None)
        SlideShapeFactory_.return_value = slide_placeholder_
        return shapes, ph_elm_, SlideShapeFactory_, slide_placeholder_

    @pytest.fixture
    def graphicFrame_fixture(
            self, slide_, rows_, cols_, x_, y_, cx_, cy_, spTree_,
            _next_shape_id_, id_, graphicFrame_):
        shapes = SlideShapeTree(slide_)
        name = 'Table 41'
        return (
            shapes, rows_, cols_, x_, y_, cx_, cy_, spTree_, id_, name,
            graphicFrame_
        )

    @pytest.fixture
    def index_fixture(self, slide_, shape_):
        shapes = SlideShapeTree(slide_)
        expected_idx = 1
        return shapes, shape_, expected_idx

    @pytest.fixture(params=[
        (PP_PLACEHOLDER.OBJECT, 3, ST_Direction.HORZ,
         'Content Placeholder 2'),
        (PP_PLACEHOLDER.TABLE,  4, ST_Direction.HORZ,
         'Table Placeholder 4'),
        (PP_PLACEHOLDER.TABLE,  7, ST_Direction.VERT,
         'Vertical Table Placeholder 6'),
        (PP_PLACEHOLDER.TITLE,  2, ST_Direction.HORZ,
         'Title 2'),
    ])
    def ph_name_fixture(self, request, slide_):
        ph_type, id_, orient, expected_name = request.param
        slide_.spTree = (
            an_spTree().with_nsdecls().with_child(
                a_cNvPr().with_name('Title 1')).with_child(
                a_cNvPr().with_name('Table Placeholder 3'))
        ).element
        shapes = SlideShapeTree(slide_)
        return shapes, ph_type, id_, orient, expected_name

    @pytest.fixture
    def pic_fixture(
            self, slide_, image_part_, rId_, x_, y_, cx_, cy_, spTree_,
            _next_shape_id_, id_, name, desc_, scaled_cx_, scaled_cy_,
            pic_):
        shapes = SlideShapeTree(slide_)
        return (
            shapes, image_part_, rId_, x_, y_, cx_, cy_, spTree_, id_,
            name, desc_, scaled_cx_, scaled_cy_, pic_
        )

    @pytest.fixture
    def picture_fixture(
            self, slide_, image_file_, x_, y_, cx_, cy_, image_part_, rId_,
            _add_pic_from_image_part_, pic_, _shape_factory_, picture_):
        shapes = SlideShapeTree(slide_)
        slide_.get_or_add_image_part.return_value = (image_part_, rId_)
        _shape_factory_.return_value = picture_
        return (
            shapes, image_file_, x_, y_, cx_, cy_, image_part_, rId_, pic_,
            picture_
        )

    @pytest.fixture
    def sp_fixture(
            self, slide_, autoshape_type_, x_, y_, cx_, cy_, spTree_,
            _next_shape_id_, id_, prst_, sp_):
        shapes = SlideShapeTree(slide_)
        name = 'Foobar 41'
        return (
            shapes, autoshape_type_, x_, y_, cx_, cy_, spTree_, id_, name,
            prst_, sp_
        )

    @pytest.fixture
    def table_fixture(
            self, rows_, cols_, x_, y_, cx_, cy_,
            _add_graphicFrame_containing_table_, _shape_factory_,
            graphicFrame_, table_):
        shapes = SlideShapeTree(None)
        _shape_factory_.return_value = table_
        return (
            shapes, rows_, cols_, x_, y_, cx_, cy_,
            _add_graphicFrame_containing_table_, _shape_factory_,
            graphicFrame_, table_
        )

    @pytest.fixture
    def textbox_fixture(
            self, x_, y_, cx_, cy_, _add_textbox_sp_, _shape_factory_,
            sp_, textbox_):
        shapes = SlideShapeTree(None)
        _shape_factory_.return_value = textbox_
        return (
            shapes, x_, y_, cx_, cy_, _add_textbox_sp_, _shape_factory_,
            sp_, textbox_
        )

    @pytest.fixture
    def textbox_sp_fixture(
            self, slide_, x_, y_, cx_, cy_, spTree_, _next_shape_id_,
            id_, sp_):
        shapes = SlideShapeTree(slide_)
        name = 'TextBox 41'
        return shapes, x_, y_, cx_, cy_, spTree_, id_, name, sp_

    @pytest.fixture(params=[
        ('p:spTree/(p:sp,p:sp/p:nvSpPr/p:nvPr/p:ph{idx=0})', True),
        ('p:spTree/(p:sp,p:sp)',                             False),
    ])
    def title_fixture(self, request, slide_, _shape_factory_, shape_):
        spTree_cxml, found = request.param
        spTree = element(spTree_cxml)
        sp = spTree.xpath('p:sp')[1]
        slide_.spTree = spTree

        shapes = SlideShapeTree(slide_)
        _shape_factory_.return_value = shape_
        calls = [call(sp)] if found else []
        expected_value = shape_ if found else None
        return shapes, _shape_factory_, calls, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _add_chart_graphicFrame_(self, request, graphicFrame_):
        return method_mock(
            request, SlideShapeTree, '_add_chart_graphicFrame',
            autospec=True, return_value=graphicFrame_
        )

    @pytest.fixture
    def _add_chart_graphic_frame_(self, request, graphic_frame_):
        return method_mock(
            request, SlideShapeTree, '_add_chart_graphic_frame',
            autospec=True, return_value=graphic_frame_
        )

    @pytest.fixture
    def _add_graphicFrame_containing_table_(self, request, graphicFrame_):
        return method_mock(
            request, SlideShapeTree, '_add_graphicFrame_containing_table',
            return_value=graphicFrame_
        )

    @pytest.fixture
    def _add_pic_from_image_part_(self, request, pic_):
        return method_mock(
            request, SlideShapeTree, '_add_pic_from_image_part',
            return_value=pic_
        )

    @pytest.fixture
    def _add_sp_from_autoshape_type_(self, request, sp_):
        return method_mock(
            request, SlideShapeTree, '_add_sp_from_autoshape_type',
            return_value=sp_
        )

    @pytest.fixture
    def _add_textbox_sp_(self, request, sp_):
        return method_mock(
            request, SlideShapeTree, '_add_textbox_sp', return_value=sp_
        )

    @pytest.fixture
    def AutoShapeType_(self, request, autoshape_type_):
        return class_mock(
            request, 'pptx.shapes.shapetree.AutoShapeType',
            return_value=autoshape_type_
        )

    @pytest.fixture
    def autoshape_type_(self, request, prst_):
        return instance_mock(
            request, AutoShapeType, basename='Foobar', prst=prst_
        )

    @pytest.fixture
    def autoshape_type_id_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, ChartData)

    @pytest.fixture
    def chart_type_(self, request):
        return instance_mock(request, EnumValue)

    @pytest.fixture
    def _clone_layout_placeholder_(self, request):
        return method_mock(
            request, SlideShapeTree, '_clone_layout_placeholder'
        )

    @pytest.fixture
    def cols_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def cx_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def cy_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def desc_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def graphicFrame_(self, request):
        return instance_mock(request, CT_GraphicalObjectFrame)

    @pytest.fixture
    def graphic_frame_(self, request):
        return instance_mock(request, GraphicFrame)

    @pytest.fixture
    def id_(self, request):
        return 42

    @pytest.fixture
    def idx_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def image_part_(self, request, desc_, scaled_cx_, scaled_cy_):
        image_part_ = instance_mock(request, ImagePart)
        image_part_.desc = desc_
        image_part_.scale.return_value = scaled_cx_, scaled_cy_
        return image_part_

    @pytest.fixture
    def image_file_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def layout_placeholder_(self, request, ph_type_, orient_, sz_, idx_):
        return instance_mock(
            request, LayoutPlaceholder, ph_type=ph_type_, orient=orient_,
            sz=sz_, idx=idx_
        )

    @pytest.fixture
    def name(self):
        return 'Picture 41'

    @pytest.fixture
    def name_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def _next_ph_name_(self, request, name_):
        return method_mock(
            request, SlideShapeTree, '_next_ph_name', return_value=name_
        )

    @pytest.fixture
    def _next_shape_id_(self, request, id_):
        return property_mock(
            request, SlideShapeTree, '_next_shape_id', return_value=id_
        )

    @pytest.fixture
    def orient_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def ph_elm_(self, request):
        return instance_mock(request, CT_Shape)

    @pytest.fixture
    def ph_type_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def pic_(self, request):
        return instance_mock(request, CT_Picture)

    @pytest.fixture
    def picture_(self, request):
        return instance_mock(request, Picture)

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, BasePlaceholder)

    @pytest.fixture
    def prst_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def rId_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def rows_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def scaled_cx_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def scaled_cy_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def shape_(self, request, sp_2_):
        return instance_mock(request, Shape, element=sp_2_)

    @pytest.fixture
    def _shape_factory_(self, request):
        return method_mock(request, SlideShapeTree, '_shape_factory')

    @pytest.fixture
    def slide_(self, request, spTree_, image_part_, rId_):
        slide_ = instance_mock(request, Slide)
        slide_.spTree = spTree_
        slide_.add_chart_part.return_value = rId_
        return slide_

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_placeholder_(self, request):
        return instance_mock(request, SlidePlaceholder)

    @pytest.fixture
    def SlideShapeFactory_(self, request):
        return function_mock(
            request, 'pptx.shapes.shapetree.SlideShapeFactory'
        )

    @pytest.fixture
    def sp_(self, request):
        return instance_mock(request, CT_Shape)

    @pytest.fixture
    def sp_2_(self, request):
        return instance_mock(request, CT_Shape)

    @pytest.fixture
    def spTree_(self, request, pic_, sp_, sp_2_, graphicFrame_):
        spTree_ = instance_mock(request, CT_GroupShape)
        spTree_.add_pic.return_value = pic_
        spTree_.add_autoshape.return_value = sp_
        spTree_.add_table.return_value = graphicFrame_
        spTree_.add_textbox.return_value = sp_
        spTree_.iter_shape_elms.return_value = [sp_, sp_2_]
        return spTree_

    @pytest.fixture
    def sz_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def table_(self, request):
        return instance_mock(request, Table)

    @pytest.fixture
    def textbox_(self, request):
        return instance_mock(request, Shape)

    @pytest.fixture
    def x_(self, request):
        return instance_mock(request, int)

    @pytest.fixture
    def y_(self, request):
        return instance_mock(request, int)
