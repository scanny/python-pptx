# encoding: utf-8

"""
Test suite for pptx.shapes.shapetree module
"""

from __future__ import absolute_import

import pytest

from pptx.oxml.autoshape import CT_Shape
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
from ..unitutil import (
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


# --------------------------------------------------------------------
# Legacy tests -------------------------------------------------------
# --------------------------------------------------------------------

import os

from hamcrest import assert_that, equal_to, is_
from mock import Mock, patch, PropertyMock

from pptx.constants import MSO_AUTO_SHAPE_TYPE as MAST
from pptx.oxml.ns import namespaces
from pptx.shapes.shapetree import ShapeCollection
from pptx.spec import (
    PH_TYPE_OBJ, PH_TYPE_TBL, PH_TYPE_TITLE, PH_ORIENT_HORZ, PH_ORIENT_VERT
)

from ..oxml.unitdata.shape import test_shape_elements, test_shapes
from ..unitutil import absjoin, parse_xml_file, TestCase, test_file_dir


test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
test_bmp_path = absjoin(test_file_dir, 'python.bmp')
new_image_path = absjoin(test_file_dir, 'monty-truth.png')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')
images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

nsmap = namespaces('a', 'r', 'p')


def _sldLayout1():
    path = os.path.join(test_file_dir, 'slideLayout1.xml')
    sldLayout = parse_xml_file(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = ShapeCollection(spTree)
    return shapes


class TestShapeCollection(TestCase):
    """Test ShapeCollection"""
    def setUp(self):
        path = absjoin(test_file_dir, 'slide1.xml')
        sld = parse_xml_file(path).getroot()
        spTree = sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        self.shapes = ShapeCollection(spTree)

    def test_construction_size(self):
        """ShapeCollection is expected size after construction"""
        # verify -----------------------
        self.assertLength(self.shapes, 9)

    @patch('pptx.shapes.shapetree.CT_Shape')
    @patch('pptx.shapes.shapetree.Shape')
    @patch('pptx.shapes.shapetree.ShapeCollection._next_sh'
           'ape_id', new_callable=PropertyMock)
    @patch('pptx.shapes.shapetree.AutoShapeType')
    def test_add_shape_collaboration(self, AutoShapeType, _next_shape_id,
                                     Shape, CT_Shape):
        """ShapeCollection.add_shape() calls the right collaborators"""
        # constant values -------------
        basename = 'Rounded Rectangle'
        prst = 'roundRect'
        id_, name = 9, '%s 8' % basename
        left, top, width, height = 111, 222, 333, 444
        autoshape_type_id = MAST.ROUNDED_RECTANGLE
        # setup mockery ---------------
        autoshape_type = Mock(name='autoshape_type')
        autoshape_type.basename = basename
        autoshape_type.prst = prst
        AutoShapeType.return_value = autoshape_type
        _next_shape_id.return_value = id_
        sp = Mock(name='sp')
        CT_Shape.new_autoshape_sp.return_value = sp
        _spTree = Mock(name='_spTree')
        _shapes = Mock(name='_shapes')
        shapes = test_shapes.empty_shape_collection
        shapes._spTree = _spTree
        shapes._shapes = _shapes
        shape = Mock('shape')
        Shape.return_value = shape
        # exercise ---------------------
        retval = shapes.add_shape(autoshape_type_id, left, top, width, height)
        # verify -----------------------
        AutoShapeType.assert_called_once_with(autoshape_type_id)
        CT_Shape.new_autoshape_sp.assert_called_once_with(
            id_, name, prst, left, top, width, height)
        Shape.assert_called_once_with(sp, shapes)
        _spTree.append.assert_called_once_with(sp)
        _shapes.append.assert_called_once_with(shape)
        assert_that(retval, is_(equal_to(shape)))

    @patch('pptx.shapes.shapetree.Picture')
    @patch('pptx.shapes.shapetree.CT_Picture')
    @patch('pptx.shapes.shapetree.ShapeCollection._next_sh'
           'ape_id', new_callable=PropertyMock)
    def test_add_picture_collaboration(self, next_shape_id, CT_Picture,
                                       Picture):
        """ShapeCollection.add_picture() calls the right collaborators"""
        # constant values -------------
        img_file = test_image_path
        left, top, width, height = 1, 2, 3, 4
        id_, name, desc = 12, 'Picture 11', 'image1.jpeg'
        rId = 'rId1'
        # setup mockery ---------------
        next_shape_id.return_value = id_
        image = Mock(name='image', _desc=desc)
        image._scale.return_value = width, height
        slide = Mock(name='slide')
        slide._add_image.return_value = image, rId
        _spTree = Mock(name='_spTree')
        _shapes = Mock(name='_shapes')
        shapes = ShapeCollection(test_shape_elements.empty_spTree, slide)
        shapes._spTree = _spTree
        shapes._shapes = _shapes
        pic = Mock(name='pic')
        CT_Picture.new_pic.return_value = pic
        picture = Mock(name='picture')
        Picture.return_value = picture
        # # exercise --------------------
        retval = shapes.add_picture(img_file, left, top, width, height)
        # verify -----------------------
        shapes._slide._add_image.assert_called_once_with(img_file)
        image._scale.assert_called_once_with(width, height)
        CT_Picture.new_pic.assert_called_once_with(
            id_, name, desc, rId, left, top, width, height)
        _spTree.append.assert_called_once_with(pic)
        Picture.assert_called_once_with(pic, shapes)
        _shapes.append.assert_called_once_with(picture)
        assert_that(retval, is_(equal_to(picture)))

    @patch('pptx.shapes.shapetree.Table')
    @patch('pptx.shapes.shapetree.CT_GraphicalObjectFrame')
    @patch('pptx.shapes.shapetree.ShapeCollection._next_sh'
           'ape_id', new_callable=PropertyMock)
    def test_add_table_collaboration(
            self, _next_shape_id, CT_GraphicalObjectFrame, Table):
        """ShapeCollection.add_table() calls the right collaborators"""
        # constant values -------------
        id_, name = 9, 'Table 8'
        rows, cols = 2, 3
        left, top, width, height = 111, 222, 333, 444
        # setup mockery ---------------
        _next_shape_id.return_value = id_
        graphicFrame = Mock(name='graphicFrame')
        CT_GraphicalObjectFrame.new_table.return_value = graphicFrame
        _spTree = Mock(name='_spTree')
        _shapes = Mock(name='_shapes')
        shapes = test_shapes.empty_shape_collection
        shapes._spTree = _spTree
        shapes._shapes = _shapes
        table = Mock('table')
        Table.return_value = table
        # exercise ---------------------
        retval = shapes.add_table(rows, cols, left, top, width, height)
        # verify -----------------------
        _next_shape_id.assert_called_once_with()
        CT_GraphicalObjectFrame.new_table.assert_called_once_with(
            id_, name, rows, cols, left, top, width, height)
        _spTree.append.assert_called_once_with(graphicFrame)
        Table.assert_called_once_with(graphicFrame, shapes)
        _shapes.append.assert_called_once_with(table)
        assert_that(retval, is_(equal_to(table)))

    @patch('pptx.shapes.shapetree.CT_Shape')
    @patch('pptx.shapes.shapetree.Shape')
    @patch('pptx.shapes.shapetree.ShapeCollection._next_sh'
           'ape_id', new_callable=PropertyMock)
    def test_add_textbox_collaboration(self, _next_shape_id, Shape,
                                       CT_Shape):
        """ShapeCollection.add_textbox() calls the right collaborators"""
        # constant values -------------
        id_, name = 9, 'TextBox 8'
        left, top, width, height = 111, 222, 333, 444
        # setup mockery ---------------
        sp = Mock(name='sp')
        shape = Mock('shape')
        _spTree = Mock(name='_spTree')
        shapes = test_shapes.empty_shape_collection
        shapes._spTree = _spTree
        _next_shape_id.return_value = id_
        CT_Shape.new_textbox_sp.return_value = sp
        Shape.return_value = shape
        # exercise ---------------------
        retval = shapes.add_textbox(left, top, width, height)
        # verify -----------------------
        CT_Shape.new_textbox_sp.assert_called_once_with(
            id_, name, left, top, width, height)
        Shape.assert_called_once_with(sp, shapes)
        _spTree.append.assert_called_once_with(sp)
        assert_that(shapes._shapes[0], is_(equal_to(shape)))
        assert_that(retval, is_(equal_to(shape)))

    def test_title_value(self):
        """ShapeCollection.title value is ref to correct shape"""
        # exercise ---------------------
        title_shape = self.shapes.title
        # verify -----------------------
        expected = 0
        actual = self.shapes.index(title_shape)
        msg = "expected shapes[%d], got shapes[%d]" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_title_is_none_on_no_title_placeholder(self):
        """ShapeCollection.title value is None when no title placeholder"""
        # setup ------------------------
        shapes = test_shapes.empty_shape_collection
        # verify -----------------------
        assert_that(shapes.title, is_(None))

    def test__next_ph_name_return_value(self):
        """
        ShapeCollection._next_ph_name() returns correct value

        * basename + 'Placeholder' + num, e.g. 'Table Placeholder 8'
        * numpart of name defaults to id-1, but increments until unique
        * prefix 'Vertical' if orient="vert"

        """
        cases = (
            (PH_TYPE_OBJ,   3, PH_ORIENT_HORZ, 'Content Placeholder 2'),
            (PH_TYPE_TBL,   4, PH_ORIENT_HORZ, 'Table Placeholder 4'),
            (PH_TYPE_TBL,   7, PH_ORIENT_VERT, 'Vertical Table Placeholder 6'),
            (PH_TYPE_TITLE, 2, PH_ORIENT_HORZ, 'Title 2'))
        # setup ------------------------
        shapes = _sldLayout1_shapes()
        for ph_type, id, orient, expected_name in cases:
            # exercise --------------------
            name = shapes._next_ph_name(ph_type, id, orient)
            # verify ----------------------
            expected = expected_name
            actual = name
            msg = ("expected placeholder name '%s', got '%s'" %
                   (expected, actual))
            self.assertEqual(expected, actual, msg)

    def test__next_shape_id_value(self):
        """ShapeCollection._next_shape_id value is correct"""
        # setup ------------------------
        shapes = _sldLayout1_shapes()
        # exercise ---------------------
        next_id = shapes._next_shape_id
        # verify -----------------------
        expected = 4
        actual = next_id
        msg = "expected %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
