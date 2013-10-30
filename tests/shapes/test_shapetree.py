# encoding: utf-8

"""Test suite for pptx.shapes module."""

from __future__ import absolute_import

import os

from hamcrest import assert_that, equal_to, is_
from mock import Mock, patch, PropertyMock

from pptx.constants import MSO_AUTO_SHAPE_TYPE as MAST
from pptx.oxml.core import SubElement
from pptx.oxml.ns import namespaces, nsdecls
from pptx.parts.slides import SlideLayout
from pptx.shapes.shapetree import Placeholder, ShapeCollection
from pptx.spec import (
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL, PH_TYPE_TITLE, PH_ORIENT_HORZ,
    PH_ORIENT_VERT
)

from ..unitdata import test_shape_elements, test_shapes
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

    def test_constructor_raises_on_contentPart_shape(self):
        """ShapeCollection() raises on contentPart shape"""
        # setup ------------------------
        spTree = test_shape_elements.empty_spTree
        SubElement(spTree, 'p:contentPart')
        # verify -----------------------
        with self.assertRaises(ValueError):
            ShapeCollection(spTree)

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
        Shape.assert_called_once_with(sp)
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
        file = test_image_path
        left, top, width, height = 1, 2, 3, 4
        id_, name, desc = 12, 'Picture 11', 'image1.jpeg'
        rId = 'rId1'
        # setup mockery ---------------
        next_shape_id.return_value = id_
        image = Mock(name='image', _desc=desc)
        image._scale.return_value = width, height
        rel = Mock(name='rel', _rId=rId)
        slide = Mock(name='slide')
        slide._add_image.return_value = image, rel
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
        retval = shapes.add_picture(file, left, top, width, height)
        # verify -----------------------
        shapes._slide._add_image.assert_called_once_with(file)
        image._scale.assert_called_once_with(width, height)
        CT_Picture.new_pic.assert_called_once_with(
            id_, name, desc, rId, left, top, width, height)
        _spTree.append.assert_called_once_with(pic)
        Picture.assert_called_once_with(pic)
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
        Table.assert_called_once_with(graphicFrame)
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
        Shape.assert_called_once_with(sp)
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

    def test_placeholders_values(self):
        """ShapeCollection.placeholders values are correct and sorted"""
        # setup ------------------------
        expected_values = (
            ('Title 1',                    PH_TYPE_CTRTITLE,  0),
            ('Vertical Subtitle 2',        PH_TYPE_SUBTITLE,  1),
            ('Date Placeholder 7',         PH_TYPE_DT,       10),
            ('Footer Placeholder 4',       PH_TYPE_FTR,      11),
            ('Slide Number Placeholder 5', PH_TYPE_SLDNUM,   12),
            ('Table Placeholder 3',        PH_TYPE_TBL,      14))
        shapes = _sldLayout1_shapes()
        # exercise ---------------------
        placeholders = shapes.placeholders
        # verify -----------------------
        for idx, ph in enumerate(placeholders):
            values = (ph.name, ph.type, ph.idx)
            expected = expected_values[idx]
            actual = values
            msg = ("expected placeholders[%d] values %s, got %s" %
                   (idx, expected, actual))
            self.assertEqual(expected, actual, msg)

    def test__clone_layout_placeholders_shapes(self):
        """ShapeCollection._clone_layout_placeholders clones shapes"""
        # setup ------------------------
        expected_values = (
            [2, 'Title 1',             PH_TYPE_CTRTITLE,  0],
            [3, 'Vertical Subtitle 2', PH_TYPE_SUBTITLE,  1],
            [4, 'Table Placeholder 3', PH_TYPE_TBL,      14])
        slidelayout = SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        shapes = test_shapes.empty_shape_collection
        # exercise ---------------------
        shapes._clone_layout_placeholders(slidelayout)
        # verify -----------------------
        for idx, sp in enumerate(shapes):
            # verify is placeholder ---
            is_placeholder = sp.is_placeholder
            msg = ("expected shapes[%d].is_placeholder == True %r"
                   % (idx, sp))
            self.assertTrue(is_placeholder, msg)
            # verify values -----------
            ph = Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = ("expected placeholder[%d] values %s, got %s"
                   % (idx, expected, actual))
            self.assertEqual(expected, actual, msg)

    def test__clone_layout_placeholder_values(self):
        """ShapeCollection._clone_layout_placeholder() values correct"""
        # setup ------------------------
        layout_shapes = _sldLayout1_shapes()
        layout_ph_shapes = [sp for sp in layout_shapes if sp.is_placeholder]
        shapes = test_shapes.empty_shape_collection
        expected_values = (
            [2, 'Title 1',                    PH_TYPE_CTRTITLE,  0],
            [3, 'Date Placeholder 2',         PH_TYPE_DT,       10],
            [4, 'Vertical Subtitle 3',        PH_TYPE_SUBTITLE,  1],
            [5, 'Table Placeholder 4',        PH_TYPE_TBL,      14],
            [6, 'Slide Number Placeholder 5', PH_TYPE_SLDNUM,   12],
            [7, 'Footer Placeholder 6',       PH_TYPE_FTR,      11])
        # exercise ---------------------
        for idx, layout_ph_sp in enumerate(layout_ph_shapes):
            layout_ph = Placeholder(layout_ph_sp)
            sp = shapes._clone_layout_placeholder(layout_ph)
            # verify ------------------
            ph = Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = "expected placeholder values %s, got %s" % (expected, actual)
            self.assertEqual(expected, actual, msg)

    def test__clone_layout_placeholder_xml(self):
        """ShapeCollection._clone_layout_placeholder() emits correct XML"""
        # setup ------------------------
        layout_shapes = _sldLayout1_shapes()
        layout_ph_shapes = [sp for sp in layout_shapes if sp.is_placeholder]
        shapes = test_shapes.empty_shape_collection
        expected_xml_tmpl = (
            '<p:sp %s>\n  <p:nvSpPr>\n    <p:cNvPr id="%s" name="%s"/>\n    <'
            'p:cNvSpPr>\n      <a:spLocks noGrp="1"/>\n    </p:cNvSpPr>\n    '
            '<p:nvPr>\n      <p:ph type="%s"%s/>\n    </p:nvPr>\n  </p:nvSpPr'
            '>\n  <p:spPr/>\n%s</p:sp>\n' %
            (nsdecls('p', 'a'), '%d', '%s', '%s', '%s', '%s')
        )
        txBody_snippet = (
            '  <p:txBody>\n    <a:bodyPr/>\n    <a:lstStyle/>\n    <a:p/>\n  '
            '</p:txBody>\n')
        expected_values = [
            (2, 'Title 1', PH_TYPE_CTRTITLE, '', txBody_snippet),
            (3, 'Date Placeholder 2', PH_TYPE_DT, ' sz="half" idx="10"', ''),
            (4, 'Vertical Subtitle 3', PH_TYPE_SUBTITLE,
                ' orient="vert" idx="1"', txBody_snippet),
            (5, 'Table Placeholder 4', PH_TYPE_TBL,
                ' sz="quarter" idx="14"', ''),
            (6, 'Slide Number Placeholder 5', PH_TYPE_SLDNUM,
                ' sz="quarter" idx="12"', ''),
            (7, 'Footer Placeholder 6', PH_TYPE_FTR,
                ' sz="quarter" idx="11"', '')]
                    # verify ----------------------
        for idx, layout_ph_sp in enumerate(layout_ph_shapes):
            layout_ph = Placeholder(layout_ph_sp)
            sp = shapes._clone_layout_placeholder(layout_ph)
            ph = Placeholder(sp)
            expected_xml = expected_xml_tmpl % expected_values[idx]
            self.assertEqualLineByLine(expected_xml, ph._element)

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
