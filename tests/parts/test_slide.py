# encoding: utf-8

"""
Test suite for pptx.parts.slide module
"""

from __future__ import absolute_import

import pytest

from pptx.chart.data import ChartData
from pptx.enum.base import EnumValue
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.opc.package import Part, _Relationship
from pptx.oxml.parts.presentation import CT_SlideId, CT_SlideIdList
from pptx.oxml.parts.slide import CT_Slide
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from pptx.oxml.shapes.groupshape import CT_GroupShape
from pptx.oxml.shapes.picture import CT_Picture
from pptx.oxml.shapes.shared import ST_Direction
from pptx.package import Package
from pptx.parts.chart import ChartPart
from pptx.parts.image import ImagePart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import (
    BaseSlide, Slide, SlideCollection, SlidePlaceholder, _SlidePlaceholders,
    _SlideShapeFactory, _SlideShapeTree
)
from pptx.parts.slidelayout import _LayoutPlaceholder, SlideLayout
from pptx.shapes.autoshape import AutoShapeType, Shape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.placeholder import BasePlaceholder
from pptx.shapes.shape import BaseShape
from pptx.shapes.table import Table

from ..oxml.unitdata.shape import (
    a_cNvPr, a_ph, a_pic, an_nvPr, an_nvSpPr, an_sp, an_spTree
)
from ..unitutil.cxml import element
from ..unitutil.file import absjoin, test_file_dir

from ..unitutil.mock import (
    ANY, call, class_mock, function_mock, initializer_mock, instance_mock,
    MagicMock, method_mock, property_mock
)


class DescribeBaseSlide(object):

    def it_provides_access_to_its_spTree_element_to_help_ShapeTree(
            self, slide):
        spTree = slide.spTree
        assert isinstance(spTree, CT_GroupShape)

    def it_knows_its_name(self, name_fixture):
        base_slide, expected_value = name_fixture
        assert base_slide.name == expected_value

    def it_can_add_an_image_part(self, image_part_fixture):
        slide, image_file, image_part_, rId_ = image_part_fixture

        image_part, rId = slide.get_or_add_image_part(image_file)

        slide._package.get_or_add_image_part.assert_called_once_with(
            image_file
        )
        slide.relate_to.assert_called_once_with(image_part_, RT.IMAGE)
        assert image_part is image_part_
        assert rId is rId_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def image_part_fixture(
            self, partname_, package_, image_part_, relate_to_):
        slide = BaseSlide(partname_, None, None, package_)
        image_file, rId = 'foobar.png', 'rId6'
        package_.get_or_add_image_part.return_value = image_part_
        relate_to_.return_value = rId
        return slide, image_file, image_part_, rId

    @pytest.fixture
    def name_fixture(self):
        sld_cxml, expected_value = 'p:sld/p:cSld{name=Foobar}', 'Foobar'
        sld = element(sld_cxml)
        base_slide = BaseSlide(None, None, sld, None)
        return base_slide, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def partname_(self):
        return PackURI('/foo/bar.xml')

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(request, BaseSlide, 'relate_to')

    @pytest.fixture
    def slide(self):
        sld = element('p:sld/p:cSld/p:spTree')
        return BaseSlide(None, None, sld, None)


class DescribeSlide(object):

    def it_can_add_a_chart_part(self, add_chart_part_fixture):
        slide, chart_type_, chart_data_ = add_chart_part_fixture[:3]
        ChartPart_, chart_part_, package_, rId = add_chart_part_fixture[3:]

        _rId = slide.add_chart_part(chart_type_, chart_data_)

        ChartPart_.new.assert_called_once_with(
            chart_type_, chart_data_, package_
        )
        slide.relate_to.assert_called_once_with(slide, chart_part_, RT.CHART)
        assert _rId is rId

    def it_provides_access_to_the_shapes_on_the_slide(self, shapes_fixture):
        slide, _SlideShapeTree_, slide_shape_tree_ = shapes_fixture
        shapes = slide.shapes
        _SlideShapeTree_.assert_called_once_with(slide)
        assert shapes is slide_shape_tree_

    def it_provides_access_to_its_placeholders(self, placeholders_fixture):
        slide, _SlidePlaceholders_, slide_placeholders_ = (
            placeholders_fixture
        )
        placeholders = slide.placeholders
        _SlidePlaceholders_.assert_called_once_with(slide)
        assert placeholders is slide_placeholders_

    def it_can_create_a_new_slide(self, new_fixture):
        slide_layout_, partname_, package_ = new_fixture[:3]
        Slide_init_, slide_elm_, shapes_, relate_to_ = new_fixture[3:]

        slide = Slide.new(slide_layout_, partname_, package_)

        Slide_init_.assert_called_once_with(
            partname_, CT.PML_SLIDE, slide_elm_, package_
        )
        shapes_.clone_layout_placeholders.assert_called_once_with(
            slide_layout_
        )
        relate_to_.assert_called_once_with(
            slide, slide_layout_, RT.SLIDE_LAYOUT
        )
        assert isinstance(slide, Slide)

    def it_knows_the_slide_layout_it_inherits_from(self, layout_fixture):
        slide, slide_layout_ = layout_fixture
        slide_layout = slide.slide_layout
        slide.part_related_by.assert_called_once_with(RT.SLIDE_LAYOUT)
        assert slide_layout is slide_layout_

    def it_knows_the_minimal_element_xml_for_a_slide(self, slide):
        path = absjoin(test_file_dir, 'minimal_slide.xml')
        sld = CT_Slide.new()
        with open(path, 'r') as f:
            expected_xml = f.read()
        assert sld.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_chart_part_fixture(
            self, package_, chart_type_, chart_data_, ChartPart_,
            chart_part_, rId, relate_to_):
        slide = Slide(None, None, None, package_)
        return (
            slide, chart_type_, chart_data_, ChartPart_, chart_part_,
            package_, rId
        )

    @pytest.fixture
    def layout_fixture(self, slide_layout_, part_related_by_):
        slide = Slide(None, None, None, None)
        return slide, slide_layout_

    @pytest.fixture
    def placeholders_fixture(
            self, _SlidePlaceholders_, slide_placeholders_):
        slide = Slide(None, None, None, None)
        return slide, _SlidePlaceholders_, slide_placeholders_

    @pytest.fixture
    def shapes_fixture(self, _SlideShapeTree_, slide_shape_tree_):
        slide = Slide(None, None, None, None)
        return slide, _SlideShapeTree_, slide_shape_tree_

    @pytest.fixture
    def new_fixture(
            self, slide_layout_, partname_, package_, Slide_init_,
            CT_Slide_, slide_elm_, shapes_prop_, shapes_,
            relate_to_):
        return (
            slide_layout_, partname_, package_, Slide_init_, slide_elm_,
            shapes_, relate_to_
        )

    # fixture components -----------------------------------

    @pytest.fixture
    def ChartPart_(self, request, chart_part_):
        ChartPart_ = class_mock(request, 'pptx.parts.slide.ChartPart')
        ChartPart_.new.return_value = chart_part_
        return ChartPart_

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, ChartData)

    @pytest.fixture
    def chart_part_(self, request):
        return instance_mock(request, ChartPart)

    @pytest.fixture
    def chart_type_(self, request):
        return instance_mock(request, EnumValue)

    @pytest.fixture
    def CT_Slide_(self, request, slide_elm_):
        CT_Slide_ = class_mock(request, 'pptx.parts.slide.CT_Slide')
        CT_Slide_.new.return_value = slide_elm_
        return CT_Slide_

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_related_by_(self, request, slide_layout_):
        return method_mock(
            request, Slide, 'part_related_by',
            return_value=slide_layout_
        )

    @pytest.fixture
    def partname_(self, request):
        return instance_mock(request, PackURI)

    @pytest.fixture
    def relate_to_(self, request, rId):
        return method_mock(
            request, Part, 'relate_to', autospec=True, return_value=rId
        )

    @pytest.fixture
    def rId(self):
        return 'rId42'

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, _SlideShapeTree)

    @pytest.fixture
    def shapes_prop_(self, request, shapes_):
        return property_mock(request, Slide, 'shapes', return_value=shapes_)

    @pytest.fixture
    def slide(self):
        return Slide(None, None, None, None)

    @pytest.fixture
    def _SlidePlaceholders_(self, request, slide_placeholders_):
        return class_mock(
            request, 'pptx.parts.slide._SlidePlaceholders',
            return_value=slide_placeholders_
        )

    @pytest.fixture
    def _SlideShapeTree_(self, request, slide_shape_tree_):
        return class_mock(
            request, 'pptx.parts.slide._SlideShapeTree',
            return_value=slide_shape_tree_
        )

    @pytest.fixture
    def slide_elm_(self, request):
        return instance_mock(request, CT_Slide)

    @pytest.fixture
    def Slide_init_(self, request):
        return initializer_mock(request, Slide)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_placeholders_(self, request):
        return instance_mock(request, _SlidePlaceholders)

    @pytest.fixture
    def slide_shape_tree_(self, request):
        return instance_mock(request, _SlideShapeTree)


class DescribeSlideCollection(object):

    def it_supports_indexed_access(self, slides_with_slide_parts_, rIds_):
        slides, slide_, slide_2_ = slides_with_slide_parts_
        rId_, rId_2_ = rIds_
        # verify -----------------------
        assert slides[0] is slide_
        assert slides[1] is slide_2_
        slides._sldIdLst.__getitem__.assert_has_calls(
            [call(0), call(1)]
        )
        slides._prs.related_parts.__getitem__.assert_has_calls(
            [call(rId_), call(rId_2_)]
        )

    def it_raises_on_slide_index_out_of_range(self, slides):
        with pytest.raises(IndexError):
            slides[2]

    def it_can_iterate_over_the_slides(self, slides, slide_, slide_2_):
        assert [s for s in slides] == [slide_, slide_2_]

    def it_supports_len(self, slides):
        assert len(slides) == 2

    def it_can_add_a_new_slide(self, slides, slidelayout_, Slide_, slide_):
        slide = slides.add_slide(slidelayout_)
        Slide_.new.assert_called_once_with(
            slidelayout_, PackURI('/ppt/slides/slide3.xml'),
            slides._prs.package
        )
        slides._prs.relate_to.assert_called_once_with(slide_, RT.SLIDE)
        slides._sldIdLst.add_sldId.assert_called_once_with(ANY)
        assert slide is slide_

    def it_knows_the_next_available_slide_partname(
            self, slides_with_slide_parts_):
        slides = slides_with_slide_parts_[0]
        expected_partname = PackURI('/ppt/slides/slide3.xml')
        partname = slides._next_partname
        assert isinstance(partname, PackURI)
        assert partname == expected_partname

    def it_can_assign_partnames_to_the_slides(
            self, slides, slide_, slide_2_):
        slides.rename_slides()
        assert slide_.partname == '/ppt/slides/slide1.xml'
        assert slide_2_.partname == '/ppt/slides/slide2.xml'

    # fixtures -------------------------------------------------------
    #
    #   slides
    #   |
    #   +- ._sldIdLst = [sldId_, sldId_2_]
    #   |                |       |
    #   |                |       +- .rId = rId_2_
    #   |                |
    #   |                +- .rId = rId_
    #   +- ._prs
    #       |
    #       +- .related_parts = {rId_: slide_, rId_2_: slide_2_}
    #
    # ----------------------------------------------------------------

    @pytest.fixture
    def prs_(self, request, rel_, related_parts_):
        prs_ = instance_mock(request, PresentationPart)
        prs_.load_rel.return_value = rel_
        prs_.related_parts = related_parts_
        return prs_

    @pytest.fixture
    def rel_(self, request, rId_):
        return instance_mock(request, _Relationship, rId=rId_)

    @pytest.fixture
    def related_parts_(self, request, rIds_, slide_parts_):
        """
        Return pass-thru mock dict that both operates as a dict an records
        calls to __getitem__ for call asserts.
        """
        rId_, rId_2_ = rIds_
        slide_, slide_2_ = slide_parts_
        slide_rId_map = {rId_: slide_, rId_2_: slide_2_}

        def getitem(key):
            return slide_rId_map[key]

        related_parts_ = MagicMock()
        related_parts_.__getitem__.side_effect = getitem
        return related_parts_

    @pytest.fixture
    def rename_slides_(self, request):
        return method_mock(request, SlideCollection, 'rename_slides')

    @pytest.fixture
    def rId_(self, request):
        return 'rId1'

    @pytest.fixture
    def rId_2_(self, request):
        return 'rId2'

    @pytest.fixture
    def rIds_(self, request, rId_, rId_2_):
        return rId_, rId_2_

    @pytest.fixture
    def Slide_(self, request, slide_):
        Slide_ = class_mock(request, 'pptx.parts.slide.Slide')
        Slide_.new.return_value = slide_
        return Slide_

    @pytest.fixture
    def sldId_(self, request, rId_):
        return instance_mock(request, CT_SlideId, rId=rId_)

    @pytest.fixture
    def sldId_2_(self, request, rId_2_):
        return instance_mock(request, CT_SlideId, rId=rId_2_)

    @pytest.fixture
    def sldIdLst_(self, request, sldId_, sldId_2_):
        sldIdLst_ = instance_mock(request, CT_SlideIdList)
        sldIdLst_.__getitem__.side_effect = [sldId_, sldId_2_]
        sldIdLst_.__iter__.return_value = iter([sldId_, sldId_2_])
        sldIdLst_.__len__.return_value = 2
        return sldIdLst_

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_2_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_parts_(self, request, slide_, slide_2_):
        return slide_, slide_2_

    @pytest.fixture
    def slidelayout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slides(self, sldIdLst_, prs_):
        return SlideCollection(sldIdLst_, prs_)

    @pytest.fixture
    def slides_with_slide_parts_(self, sldIdLst_, prs_, slide_parts_):
        slide_, slide_2_ = slide_parts_
        slides = SlideCollection(sldIdLst_, prs_)
        return slides, slide_, slide_2_


class Describe_SlideShapeTree(object):

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
        shapes, ph_elm_, _SlideShapeFactory_, slide_placeholder_ = (
            factory_fixture
        )
        slide_placeholder = shapes._shape_factory(ph_elm_)
        _SlideShapeFactory_.assert_called_once_with(ph_elm_, shapes)
        assert slide_placeholder is slide_placeholder_

    def it_can_find_the_title_placeholder(self, title_fixture):
        shapes, _shape_factory_, sp_2_, title_placeholder_ = title_fixture
        title_placeholder = shapes.title
        _shape_factory_.assert_called_once_with(sp_2_)
        assert title_placeholder == title_placeholder_

    def it_returns_None_when_slide_has_no_title_ph(self, no_title_fixture):
        shapes = no_title_fixture
        title_placeholder = shapes.title
        assert title_placeholder is None

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
        _SlideShapeFactory_, graphicFrame_ = add_chart_graph_frm_fixture[7:9]
        graphic_frame_ = add_chart_graph_frm_fixture[9]

        graphic_frame = shape_tree._add_chart_graphic_frame(
            rId_, x, y, cx, cy
        )

        _add_chart_graphicFrame_.assert_called_once_with(
            shape_tree, rId_, x, y, cx, cy
        )
        _SlideShapeFactory_.assert_called_once_with(
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
        shape_tree = _SlideShapeTree(slide_)
        x, y, cx, cy = 1, 2, 3, 4
        return (
            shape_tree, chart_type_, x, y, cx, cy, chart_data_, rId_,
            graphic_frame_
        )

    @pytest.fixture
    def add_chart_graph_frm_fixture(
            self, rId_, x_, y_, cx_, cy_, _add_chart_graphicFrame_,
            _SlideShapeFactory_, graphicFrame_, graphic_frame_):
        shape_tree = _SlideShapeTree(None)
        _SlideShapeFactory_.return_value = graphic_frame_
        return (
            shape_tree, rId_, x_, y_, cx_, cy_, _add_chart_graphicFrame_,
            _SlideShapeFactory_, graphicFrame_, graphic_frame_
        )

    @pytest.fixture
    def add_chart_graphicFrame_fixture(self, slide_, _next_shape_id_, id_):
        rId, x, y, cx, cy = 'rId42', 1, 2, 3, 4
        shape_tree = _SlideShapeTree(slide_)
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
        shapes = _SlideShapeTree(None)
        _shape_factory_.return_value = shape_
        return (
            shapes, autoshape_type_id_, x_, y_, cx_, cy_, AutoShapeType_,
            _add_sp_from_autoshape_type_, autoshape_type_, _shape_factory_,
            sp_, shape_
        )

    @pytest.fixture
    def clone_fixture(
            self, slide_layout_, placeholder_, _clone_layout_placeholder_):
        shapes = _SlideShapeTree(None)
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
        shapes = _SlideShapeTree(slide_)
        return (
            shapes, layout_placeholder_, spTree_, id_, name_, ph_type_,
            orient_, sz_, idx_
        )

    @pytest.fixture
    def factory_fixture(
            self, ph_elm_, _SlideShapeFactory_, slide_placeholder_):
        shapes = _SlideShapeTree(None)
        _SlideShapeFactory_.return_value = slide_placeholder_
        return shapes, ph_elm_, _SlideShapeFactory_, slide_placeholder_

    @pytest.fixture
    def graphicFrame_fixture(
            self, slide_, rows_, cols_, x_, y_, cx_, cy_, spTree_,
            _next_shape_id_, id_, graphicFrame_):
        shapes = _SlideShapeTree(slide_)
        name = 'Table 41'
        return (
            shapes, rows_, cols_, x_, y_, cx_, cy_, spTree_, id_, name,
            graphicFrame_
        )

    @pytest.fixture
    def index_fixture(self, slide_, shape_):
        shapes = _SlideShapeTree(slide_)
        expected_idx = 1
        return shapes, shape_, expected_idx

    @pytest.fixture
    def no_title_fixture(self, slide_, spTree_, sp_):
        shapes = _SlideShapeTree(slide_)
        spTree_.iter_shape_elms.return_value = [sp_, sp_]
        return shapes

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
        shapes = _SlideShapeTree(slide_)
        return shapes, ph_type, id_, orient, expected_name

    @pytest.fixture
    def pic_fixture(
            self, slide_, image_part_, rId_, x_, y_, cx_, cy_, spTree_,
            _next_shape_id_, id_, name, desc_, scaled_cx_, scaled_cy_,
            pic_):
        shapes = _SlideShapeTree(slide_)
        return (
            shapes, image_part_, rId_, x_, y_, cx_, cy_, spTree_, id_,
            name, desc_, scaled_cx_, scaled_cy_, pic_
        )

    @pytest.fixture
    def picture_fixture(
            self, slide_, image_file_, x_, y_, cx_, cy_, image_part_, rId_,
            _add_pic_from_image_part_, pic_, _shape_factory_, picture_):
        shapes = _SlideShapeTree(slide_)
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
        shapes = _SlideShapeTree(slide_)
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
        shapes = _SlideShapeTree(None)
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
        shapes = _SlideShapeTree(None)
        _shape_factory_.return_value = textbox_
        return (
            shapes, x_, y_, cx_, cy_, _add_textbox_sp_, _shape_factory_,
            sp_, textbox_
        )

    @pytest.fixture
    def textbox_sp_fixture(
            self, slide_, x_, y_, cx_, cy_, spTree_, _next_shape_id_,
            id_, sp_):
        shapes = _SlideShapeTree(slide_)
        name = 'TextBox 41'
        return shapes, x_, y_, cx_, cy_, spTree_, id_, name, sp_

    @pytest.fixture
    def title_fixture(self, slide_, sp_2_, _shape_factory_, shape_):
        shapes = _SlideShapeTree(slide_)
        _shape_factory_.return_value = shape_
        title_placeholder_ = shape_
        return shapes, _shape_factory_, sp_2_, title_placeholder_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _add_chart_graphicFrame_(self, request, graphicFrame_):
        return method_mock(
            request, _SlideShapeTree, '_add_chart_graphicFrame',
            autospec=True, return_value=graphicFrame_
        )

    @pytest.fixture
    def _add_chart_graphic_frame_(self, request, graphic_frame_):
        return method_mock(
            request, _SlideShapeTree, '_add_chart_graphic_frame',
            autospec=True, return_value=graphic_frame_
        )

    @pytest.fixture
    def _add_graphicFrame_containing_table_(self, request, graphicFrame_):
        return method_mock(
            request, _SlideShapeTree, '_add_graphicFrame_containing_table',
            return_value=graphicFrame_
        )

    @pytest.fixture
    def _add_pic_from_image_part_(self, request, pic_):
        return method_mock(
            request, _SlideShapeTree, '_add_pic_from_image_part',
            return_value=pic_
        )

    @pytest.fixture
    def _add_sp_from_autoshape_type_(self, request, sp_):
        return method_mock(
            request, _SlideShapeTree, '_add_sp_from_autoshape_type',
            return_value=sp_
        )

    @pytest.fixture
    def _add_textbox_sp_(self, request, sp_):
        return method_mock(
            request, _SlideShapeTree, '_add_textbox_sp', return_value=sp_
        )

    @pytest.fixture
    def AutoShapeType_(self, request, autoshape_type_):
        return class_mock(
            request, 'pptx.parts.slide.AutoShapeType',
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
            request, _SlideShapeTree, '_clone_layout_placeholder'
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
        image_part_._desc = desc_
        image_part_.scale.return_value = scaled_cx_, scaled_cy_
        return image_part_

    @pytest.fixture
    def image_file_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def layout_placeholder_(self, request, ph_type_, orient_, sz_, idx_):
        return instance_mock(
            request, _LayoutPlaceholder, ph_type=ph_type_, orient=orient_,
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
            request, _SlideShapeTree, '_next_ph_name', return_value=name_
        )

    @pytest.fixture
    def _next_shape_id_(self, request, id_):
        return property_mock(
            request, _SlideShapeTree, '_next_shape_id', return_value=id_
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
        return method_mock(request, _SlideShapeTree, '_shape_factory')

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
    def _SlideShapeFactory_(self, request):
        return function_mock(request, 'pptx.parts.slide._SlideShapeFactory')

    @pytest.fixture
    def sp_(self, request):
        return instance_mock(request, CT_Shape)

    @pytest.fixture
    def sp_2_(self, request):
        return instance_mock(request, CT_Shape, ph_idx=0)

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


class Describe_SlideShapeFactory(object):

    def it_constructs_a_slide_placeholder_for_a_shape_element(
            self, factory_fixture):
        shape_elm, parent_, ShapeConstructor_, shape_ = factory_fixture
        shape = _SlideShapeFactory(shape_elm, parent_)
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
            request, 'pptx.parts.slide.BaseShapeFactory',
            return_value=base_shape_
        )

    @pytest.fixture
    def base_shape_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def SlidePlaceholder_(self, request, slide_placeholder_):
        return class_mock(
            request, 'pptx.parts.slide.SlidePlaceholder',
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


class Describe_SlidePlaceholders(object):

    def it_constructs_a_slide_placeholder_for_a_placeholder_shape(
            self, factory_fixture):
        slide_placeholders, ph_elm_ = factory_fixture[:2]
        _SlideShapeFactory_, slide_placeholder_ = factory_fixture[2:]
        slide_placeholder = slide_placeholders._shape_factory(ph_elm_)
        _SlideShapeFactory_.assert_called_once_with(
            ph_elm_, slide_placeholders
        )
        assert slide_placeholder is slide_placeholder_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def factory_fixture(
            self, ph_elm_, _SlideShapeFactory_, slide_placeholder_):
        slide_placeholders = _SlidePlaceholders(None)
        return (
            slide_placeholders, ph_elm_, _SlideShapeFactory_,
            slide_placeholder_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def slide_placeholder_(self, request):
        return instance_mock(request, SlidePlaceholder)

    @pytest.fixture
    def _SlideShapeFactory_(self, request, slide_placeholder_):
        return function_mock(
            request, 'pptx.parts.slide._SlideShapeFactory',
            return_value=slide_placeholder_
        )

    @pytest.fixture
    def ph_elm_(self, request):
        return instance_mock(request, CT_Shape)
