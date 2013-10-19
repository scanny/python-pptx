# -*- coding: utf-8 -*-

"""Test suite for pptx.shapes module."""

import os

from hamcrest import assert_that, equal_to, is_, is_not
from mock import Mock, patch, PropertyMock

from pptx.constants import MSO_AUTO_SHAPE_TYPE as MAST, MSO
from pptx.oxml import _SubElement, nsdecls, oxml_fromstring, oxml_parse
from pptx.presentation import _SlideLayout
from pptx.shapes import (
    _Adjustment, _AdjustmentCollection, _AutoShapeType, _Placeholder, _Shape,
    _ShapeCollection
)
from pptx.spec import namespaces
from pptx.spec import (
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL, PH_TYPE_TITLE, PH_ORIENT_HORZ,
    PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER
)
from testdata import a_prstGeom, test_shape_elements, test_shapes
from testing import TestCase


# module globals -------------------------------------------------------------
def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir = os.path.split(__file__)[0]
test_file_dir = absjoin(thisdir, 'test_files')

test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
test_bmp_path = absjoin(test_file_dir, 'python.bmp')
new_image_path = absjoin(test_file_dir, 'monty-truth.png')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')
images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

nsmap = namespaces('a', 'r', 'p')


def _sldLayout1():
    path = os.path.join(thisdir, 'test_files/slideLayout1.xml')
    sldLayout = oxml_parse(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = _ShapeCollection(spTree)
    return shapes


class Test_Adjustment(TestCase):
    """Test _Adjustment"""
    def test_it_should_have_correct_effective_value(self):
        """_Adjustment.effective_value is correct"""
        # setup ------------------------
        name = "don't care"
        cases = (
            # no actual, effective should be determined by default value
            (50000, None, 0.5),
            # actual matches default
            (50000, 50000, 0.5),
            # actual is different than default
            (50000, 12500, 0.125),
            # actual is zero
            (50000, 0, 0.0),
            # negative default
            (-20833, None, -0.20833),
            # negative actual
            (-20833, -5678901, -56.78901),
        )
        # verify -----------------------
        for def_val, actual, expected in cases:
            adjustment = _Adjustment(name, def_val, actual)
            assert_that(adjustment.effective_value, is_(equal_to(expected)))


class Test_AdjustmentCollection(TestCase):
    """Test _AdjustmentCollection"""
    def test_it_should_load_default_adjustment_values(self):
        """_AdjustmentCollection() loads default adjustment values"""
        # setup ------------------------
        cases = (
            ('rect', ()),
            ('chevron', (('adj', 50000),)),
            ('accentBorderCallout1',
             (('adj1', 18750), ('adj2', -8333), ('adj3', 112500),
              ('adj4', -38333))),
            ('wedgeRoundRectCallout',
             (('adj1', -20833), ('adj2', 62500), ('adj3', 16667))),
            ('circularArrow',
             (('adj1', 12500), ('adj2', 1142319), ('adj3', 20457681),
              ('adj4', 10800000), ('adj5', 12500))),
        )
        for prst, expected_values in cases:
            prstGeom = a_prstGeom(prst).with_avLst.element
            adjustments = _AdjustmentCollection(prstGeom)._adjustments
            # verify -------------------
            reason = ("\n     failed for prst: '%s'" % prst)
            assert_that(len(adjustments),
                        is_(equal_to(len(expected_values))),
                        reason)
            actuals = tuple([(adj.name, adj.def_val) for adj in adjustments])
            assert_that(actuals, is_(equal_to(expected_values)), reason)

    def test_it_should_load_adj_val_actuals_from_xml(self):
        """_AdjustmentCollection() loads adjustment value actuals from XML"""
        # setup ------------------------
        def expected_actual(adj_vals, adjustment_name):
            for name, actual in adj_vals:
                if name == adjustment_name:
                    return actual
            return None

        cases = (
            # no adjustment values in xml or spec
            (a_prstGeom('rect').with_avLst, ()),
            # no adjustment values in xml, but some in spec
            (a_prstGeom('circularArrow'), ()),
            # adjustment value in xml but none in spec
            (a_prstGeom('rect').with_gd(), ()),
            # middle adjustment value in xml
            (a_prstGeom('mathDivide').with_gd(name='adj2'),
             (('adj2', 25000),)),
            # all adjustment values in xml
            (a_prstGeom('wedgeRoundRectCallout').with_gd(111, 'adj1')
                                                .with_gd(222, 'adj2')
                                                .with_gd(333, 'adj3'),
             (('adj1', 111), ('adj2', 222), ('adj3', 333))),
        )
        # verify -----------------------
        for prstGeom_builder, adj_vals in cases:
            prstGeom = prstGeom_builder.element
            adjustments = _AdjustmentCollection(prstGeom)._adjustments
            for adjustment in adjustments:
                expected_value = expected_actual(adj_vals, adjustment.name)
                reason = (
                    "failed for adj val name '%s', for this XML:\n\n%s" %
                    (adjustment.name, prstGeom_builder.xml))
                assert_that(adjustment.actual, is_(expected_value), reason)

    def test_it_should_return_effective_value_on_indexed_access(self):
        """_AdjustmentCollection[n] is normalized effective value of nth"""
        # setup ------------------------
        cases = (
            ('rect', ()),
            ('chevron', (0.5,)),
            ('circularArrow', (0.125, 11.42319, 204.57681, 108.0, 0.125)),
        )
        for prst, expected_values in cases:
            prstGeom = a_prstGeom(prst).element
            # exercise -----------------
            adjustments = _AdjustmentCollection(prstGeom)
            # verify -------------------
            reason = "failed on case: prst=='%s'" % prst
            assert_that(len(adjustments), is_(len(expected_values)), reason)
            retvals = tuple([adj for adj in adjustments])
            assert_that(retvals, is_(equal_to(expected_values)), reason)

    def test_it_should_update_actual_value_on_indexed_assignment(self):
        """Assignment to _AdjustmentCollection[n] updates nth actual"""
        # setup ------------------------
        cases = (
            ('chevron', 0, 0.5, 50000),
            ('circularArrow', 2, 99.99, 9999000),
        )
        prst = 'chevron'
        for prst, idx, new_value, expected in cases:
            prstGeom = a_prstGeom(prst).element
            adjustments = _AdjustmentCollection(prstGeom)
            # exercise -----------------
            adjustments[idx] = new_value
            # verify -------------------
            reason = "failed on case: prst=='%s'" % prst
            assert_that(adjustments._adjustments[idx].actual,
                        is_(equal_to(expected)),
                        reason)

    def test_it_should_round_trip_indexed_assignment(self):
        """Assignment to _AdjustmentCollection[n] round-trips"""
        # setup ------------------------
        new_value = 0.375
        prstGeom = a_prstGeom('chevron').element
        adjustments = _AdjustmentCollection(prstGeom)
        assert_that(adjustments[0], is_not(equal_to(new_value)))
        # exercise ---------------------
        adjustments[0] = new_value
        # verify -----------------------
        assert_that(adjustments[0], is_(equal_to(new_value)))

    def test_it_should_raise_on_bad_key(self):
        """_AdjustmentCollection[idx] raises on invalid idx"""
        # setup ------------------------
        prstGeom = a_prstGeom('chevron').element
        adjustments = _AdjustmentCollection(prstGeom)
        # verify -----------------------
        with self.assertRaises(IndexError):
            adjustments[-6]
        with self.assertRaises(IndexError):
            adjustments[6]
        with self.assertRaises(TypeError):
            adjustments[0.0]
        with self.assertRaises(TypeError):
            adjustments['0']
        with self.assertRaises(IndexError):
            adjustments[-6] = 1.0
        with self.assertRaises(IndexError):
            adjustments[6] = 1.0
        with self.assertRaises(TypeError):
            adjustments[0.0] = 1.0
        with self.assertRaises(TypeError):
            adjustments['0'] = 1.0

    def test_it_should_raise_on_assigned_bad_value(self):
        """_AdjustmentCollection[n] = val raises on val is not number"""
        # setup ------------------------
        prstGeom = a_prstGeom('chevron').element
        adjustments = _AdjustmentCollection(prstGeom)
        # verify -----------------------
        with self.assertRaises(ValueError):
            adjustments[0] = 'foobar'

    def test_writes_adj_vals_to_xml_on_assignment(self):
        """_AdjustmentCollection writes adj vals to XML on assignment"""
        # setup ------------------------
        prstGeom = a_prstGeom('chevron').element
        adjustments = _AdjustmentCollection(prstGeom)
        __prstGeom = Mock(name='__prstGeom')
        adjustments._AdjustmentCollection__prstGeom = __prstGeom
        # exercise ---------------------
        adjustments[0] = 0.999
        # verify -----------------------
        assert_that(__prstGeom.rewrite_guides.call_count, is_(1))


class Test_AutoShapeType(TestCase):
    """Test _AutoShapeType"""
    def test_construction_return_values(self):
        """_AutoShapeType() returns instance with correct property values"""
        # setup ------------------------
        id_ = MAST.ROUNDED_RECTANGLE
        prst = 'roundRect'
        basename = 'Rounded Rectangle'
        # exercise ---------------------
        autoshape_type = _AutoShapeType(id_)
        # verify -----------------------
        assert_that(autoshape_type.autoshape_type_id, is_(equal_to(id_)))
        assert_that(autoshape_type.prst, is_(equal_to(prst)))
        assert_that(autoshape_type.basename, is_(equal_to(basename)))

    def test_default_adjustment_values_return_value(self):
        """_AutoShapeType.default_adjustment_values() return val correct"""
        # setup ------------------------
        cases = (
            ('rect', ()),
            ('chevron', (('adj', 50000),)),
            ('leftCircularArrow',
             (('adj1', 12500), ('adj2', -1142319), ('adj3', 1142319),
              ('adj4', 10800000), ('adj5', 12500))),
        )
        # verify -----------------------
        for prst, expected_vals in cases:
            def_adj_vals = _AutoShapeType.default_adjustment_values(prst)
            assert_that(def_adj_vals, is_(equal_to(expected_vals)))

    def test__lookup_id_by_prst_return_value(self):
        """_AutoShapeType._lookup_id_by_prst() return value is correct"""
        # setup ------------------------
        autoshape_type_id = MAST.ROUNDED_RECTANGLE
        prst = 'roundRect'
        # exercise ---------------------
        retval = _AutoShapeType._lookup_id_by_prst(prst)
        # verify -----------------------
        assert_that(retval, is_(equal_to(autoshape_type_id)))

    def test__lookup_id_raises_on_bad_prst(self):
        """_AutoShapeType._lookup_id_by_prst() raises on bad prst"""
        # setup ------------------------
        prst = 'badPrst'
        # verify -----------------------
        with self.assertRaises(KeyError):
            _AutoShapeType._lookup_id_by_prst(prst)

    def test_second_construction_returns_cached_instance(self):
        """_AutoShapeType() returns cached instance on duplicate call"""
        # setup ------------------------
        id_ = MAST.ROUNDED_RECTANGLE
        ast1 = _AutoShapeType(id_)
        # exercise ---------------------
        ast2 = _AutoShapeType(id_)
        # verify -----------------------
        assert_that(ast2, is_(equal_to(ast1)))

    def test_construction_raises_on_bad_autoshape_type_id(self):
        """_AutoShapeType() raises on bad auto shape type id"""
        # setup ------------------------
        autoshape_type_id = 9999
        # verify -----------------------
        with self.assertRaises(KeyError):
            _AutoShapeType(autoshape_type_id)


class Test_Placeholder(TestCase):
    """Test _Placeholder"""
    def test_property_values(self):
        """_Placeholder property values are correct"""
        # setup ------------------------
        expected_values = (
            (PH_TYPE_CTRTITLE, PH_ORIENT_HORZ, PH_SZ_FULL,     0),
            (PH_TYPE_DT,       PH_ORIENT_HORZ, PH_SZ_HALF,    10),
            (PH_TYPE_SUBTITLE, PH_ORIENT_VERT, PH_SZ_FULL,     1),
            (PH_TYPE_TBL,      PH_ORIENT_HORZ, PH_SZ_QUARTER, 14),
            (PH_TYPE_SLDNUM,   PH_ORIENT_HORZ, PH_SZ_QUARTER, 12),
            (PH_TYPE_FTR,      PH_ORIENT_HORZ, PH_SZ_QUARTER, 11))
        shapes = _sldLayout1_shapes()
        # exercise ---------------------
        for idx, sp in enumerate(shapes):
            ph = _Placeholder(sp)
            values = (ph.type, ph.orient, ph.sz, ph.idx)
            # verify ----------------------
            expected = expected_values[idx]
            actual = values
            msg = ("expected shapes[%d] values %s, got %s"
                   % (idx, expected, actual))
            self.assertEqual(expected, actual, msg)


class Test_Picture(TestCase):
    """Test _Picture"""
    def test_shape_type_value_correct_for_picture(self):
        """_Shape.shape_type value is correct for picture"""
        # setup ------------------------
        picture = test_shapes.picture
        # verify -----------------------
        assert_that(picture.shape_type, is_(equal_to(MSO.PICTURE)))


class Test_Shape(TestCase):
    """Test _Shape"""
    @patch('pptx.shapes._BaseShape.__init__')
    @patch('pptx.shapes._AdjustmentCollection')
    def test_it_initializes_adjustments_on_construction(
            self, _AdjustmentCollection, _BaseShape__init__):
        """_Shape() initializes adjustments on construction"""
        # setup ------------------------
        adjustments = Mock(name='adjustments')
        _AdjustmentCollection.return_value = adjustments
        sp = Mock(name='sp')
        # exercise ---------------------
        shape = _Shape(sp)
        # verify -----------------------
        _BaseShape__init__.assert_called_once_with(sp)
        _AdjustmentCollection.assert_called_once_with(sp.prstGeom)
        assert_that(shape.adjustments, is_(adjustments))

    def test_auto_shape_type_value_correct(self):
        """_Shape.auto_shape_type value is correct"""
        # setup ------------------------
        rounded_rectangle = test_shapes.rounded_rectangle
        # verify -----------------------
        assert_that(rounded_rectangle.auto_shape_type,
                    is_(equal_to(MAST.ROUNDED_RECTANGLE)))

    def test_auto_shape_type_raises_on_non_auto_shape(self):
        """_Shape.auto_shape_type raises on non auto shape"""
        # setup ------------------------
        textbox = test_shapes.textbox
        # verify -----------------------
        with self.assertRaises(ValueError):
            textbox.auto_shape_type

    def test_shape_type_value_correct(self):
        """_Shape.shape_type value is correct for all sub-types"""
        # setup ------------------------
        autoshape = test_shapes.autoshape
        placeholder = test_shapes.placeholder
        textbox = test_shapes.textbox
        # verify -----------------------
        assert_that(autoshape.shape_type, is_(equal_to(MSO.AUTO_SHAPE)))
        assert_that(placeholder.shape_type, is_(equal_to(MSO.PLACEHOLDER)))
        assert_that(textbox.shape_type, is_(equal_to(MSO.TEXT_BOX)))

    def test_shape_type_raises_on_unrecognized_shape_type(self):
        """_Shape.shape_type raises on unrecognized shape type"""
        # setup ------------------------
        xml = (
            '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/'
            '2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/'
            '2006/main"><p:nvSpPr><p:cNvPr id="9" name="Unknown Shape Type 8"'
            '/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>'
        )
        sp = oxml_fromstring(xml)
        shape = _Shape(sp)
        # verify -----------------------
        with self.assertRaises(NotImplementedError):
            shape.shape_type


class Test_ShapeCollection(TestCase):
    """Test _ShapeCollection"""
    def setUp(self):
        path = absjoin(test_file_dir, 'slide1.xml')
        sld = oxml_parse(path).getroot()
        spTree = sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        self.shapes = _ShapeCollection(spTree)

    def test_construction_size(self):
        """_ShapeCollection is expected size after construction"""
        # verify -----------------------
        self.assertLength(self.shapes, 9)

    def test_constructor_raises_on_contentPart_shape(self):
        """_ShapeCollection() raises on contentPart shape"""
        # setup ------------------------
        spTree = test_shape_elements.empty_spTree
        _SubElement(spTree, 'p:contentPart')
        # verify -----------------------
        with self.assertRaises(ValueError):
            _ShapeCollection(spTree)

    @patch('pptx.shapes.CT_Shape')
    @patch('pptx.shapes._Shape')
    @patch('pptx.shapes._ShapeCollection._ShapeCollection__next_shape_id',
           new_callable=PropertyMock)
    @patch('pptx.shapes._AutoShapeType')
    def test_add_shape_collaboration(self, _AutoShapeType, __next_shape_id,
                                     _Shape, CT_Shape):
        """_ShapeCollection.add_shape() calls the right collaborators"""
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
        _AutoShapeType.return_value = autoshape_type
        __next_shape_id.return_value = id_
        sp = Mock(name='sp')
        CT_Shape.new_autoshape_sp.return_value = sp
        __spTree = Mock(name='__spTree')
        __shapes = Mock(name='__shapes')
        shapes = test_shapes.empty_shape_collection
        shapes._ShapeCollection__spTree = __spTree
        shapes._ShapeCollection__shapes = __shapes
        shape = Mock('shape')
        _Shape.return_value = shape
        # exercise ---------------------
        retval = shapes.add_shape(autoshape_type_id, left, top, width, height)
        # verify -----------------------
        _AutoShapeType.assert_called_once_with(autoshape_type_id)
        CT_Shape.new_autoshape_sp.assert_called_once_with(
            id_, name, prst, left, top, width, height)
        _Shape.assert_called_once_with(sp)
        __spTree.append.assert_called_once_with(sp)
        __shapes.append.assert_called_once_with(shape)
        assert_that(retval, is_(equal_to(shape)))

    @patch('pptx.shapes._Picture')
    @patch('pptx.shapes.CT_Picture')
    @patch('pptx.shapes._ShapeCollection._ShapeCollection__next_shape_id',
           new_callable=PropertyMock)
    def test_add_picture_collaboration(self, next_shape_id, CT_Picture,
                                       _Picture):
        """_ShapeCollection.add_picture() calls the right collaborators"""
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
        __spTree = Mock(name='__spTree')
        __shapes = Mock(name='__shapes')
        shapes = _ShapeCollection(test_shape_elements.empty_spTree, slide)
        shapes._ShapeCollection__spTree = __spTree
        shapes._ShapeCollection__shapes = __shapes
        pic = Mock(name='pic')
        CT_Picture.new_pic.return_value = pic
        picture = Mock(name='picture')
        _Picture.return_value = picture
        # # exercise --------------------
        retval = shapes.add_picture(file, left, top, width, height)
        # verify -----------------------
        shapes._ShapeCollection__slide._add_image.assert_called_once_with(file)
        image._scale.assert_called_once_with(width, height)
        CT_Picture.new_pic.assert_called_once_with(
            id_, name, desc, rId, left, top, width, height)
        __spTree.append.assert_called_once_with(pic)
        _Picture.assert_called_once_with(pic)
        __shapes.append.assert_called_once_with(picture)
        assert_that(retval, is_(equal_to(picture)))

    @patch('pptx.shapes._Table')
    @patch('pptx.shapes.CT_GraphicalObjectFrame')
    @patch('pptx.shapes._ShapeCollection._ShapeCollection__next_shape_id',
           new_callable=PropertyMock)
    def test_add_table_collaboration(
            self, __next_shape_id, CT_GraphicalObjectFrame, _Table):
        """_ShapeCollection.add_table() calls the right collaborators"""
        # constant values -------------
        id_, name = 9, 'Table 8'
        rows, cols = 2, 3
        left, top, width, height = 111, 222, 333, 444
        # setup mockery ---------------
        __next_shape_id.return_value = id_
        graphicFrame = Mock(name='graphicFrame')
        CT_GraphicalObjectFrame.new_table.return_value = graphicFrame
        __spTree = Mock(name='__spTree')
        __shapes = Mock(name='__shapes')
        shapes = test_shapes.empty_shape_collection
        shapes._ShapeCollection__spTree = __spTree
        shapes._ShapeCollection__shapes = __shapes
        table = Mock('table')
        _Table.return_value = table
        # exercise ---------------------
        retval = shapes.add_table(rows, cols, left, top, width, height)
        # verify -----------------------
        __next_shape_id.assert_called_once_with()
        CT_GraphicalObjectFrame.new_table.assert_called_once_with(
            id_, name, rows, cols, left, top, width, height)
        __spTree.append.assert_called_once_with(graphicFrame)
        _Table.assert_called_once_with(graphicFrame)
        __shapes.append.assert_called_once_with(table)
        assert_that(retval, is_(equal_to(table)))

    @patch('pptx.shapes.CT_Shape')
    @patch('pptx.shapes._Shape')
    @patch('pptx.shapes._ShapeCollection._ShapeCollection__next_shape_id',
           new_callable=PropertyMock)
    def test_add_textbox_collaboration(self, __next_shape_id, _Shape,
                                       CT_Shape):
        """_ShapeCollection.add_textbox() calls the right collaborators"""
        # constant values -------------
        id_, name = 9, 'TextBox 8'
        left, top, width, height = 111, 222, 333, 444
        # setup mockery ---------------
        sp = Mock(name='sp')
        shape = Mock('shape')
        __spTree = Mock(name='__spTree')
        shapes = test_shapes.empty_shape_collection
        shapes._ShapeCollection__spTree = __spTree
        __next_shape_id.return_value = id_
        CT_Shape.new_textbox_sp.return_value = sp
        _Shape.return_value = shape
        # exercise ---------------------
        retval = shapes.add_textbox(left, top, width, height)
        # verify -----------------------
        CT_Shape.new_textbox_sp.assert_called_once_with(
            id_, name, left, top, width, height)
        _Shape.assert_called_once_with(sp)
        __spTree.append.assert_called_once_with(sp)
        assert_that(shapes._ShapeCollection__shapes[0], is_(equal_to(shape)))
        assert_that(retval, is_(equal_to(shape)))

    def test_title_value(self):
        """_ShapeCollection.title value is ref to correct shape"""
        # exercise ---------------------
        title_shape = self.shapes.title
        # verify -----------------------
        expected = 0
        actual = self.shapes.index(title_shape)
        msg = "expected shapes[%d], got shapes[%d]" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_title_is_none_on_no_title_placeholder(self):
        """_ShapeCollection.title value is None when no title placeholder"""
        # setup ------------------------
        shapes = test_shapes.empty_shape_collection
        # verify -----------------------
        assert_that(shapes.title, is_(None))

    def test_placeholders_values(self):
        """_ShapeCollection.placeholders values are correct and sorted"""
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
        """_ShapeCollection._clone_layout_placeholders clones shapes"""
        # setup ------------------------
        expected_values = (
            [2, 'Title 1',             PH_TYPE_CTRTITLE,  0],
            [3, 'Vertical Subtitle 2', PH_TYPE_SUBTITLE,  1],
            [4, 'Table Placeholder 3', PH_TYPE_TBL,      14])
        slidelayout = _SlideLayout()
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
            ph = _Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = ("expected placeholder[%d] values %s, got %s"
                   % (idx, expected, actual))
            self.assertEqual(expected, actual, msg)

    def test___clone_layout_placeholder_values(self):
        """_ShapeCollection.__clone_layout_placeholder() values correct"""
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
            layout_ph = _Placeholder(layout_ph_sp)
            sp = shapes._ShapeCollection__clone_layout_placeholder(layout_ph)
            # verify ------------------
            ph = _Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = "expected placeholder values %s, got %s" % (expected, actual)
            self.assertEqual(expected, actual, msg)

    def test___clone_layout_placeholder_xml(self):
        """_ShapeCollection.__clone_layout_placeholder() emits correct XML"""
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
            layout_ph = _Placeholder(layout_ph_sp)
            sp = shapes._ShapeCollection__clone_layout_placeholder(layout_ph)
            ph = _Placeholder(sp)
            expected_xml = expected_xml_tmpl % expected_values[idx]
            self.assertEqualLineByLine(expected_xml, ph._element)

    def test___next_ph_name_return_value(self):
        """
        _ShapeCollection.__next_ph_name() returns correct value

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
            name = shapes._ShapeCollection__next_ph_name(ph_type, id, orient)
            # verify ----------------------
            expected = expected_name
            actual = name
            msg = ("expected placeholder name '%s', got '%s'" %
                   (expected, actual))
            self.assertEqual(expected, actual, msg)

    def test___next_shape_id_value(self):
        """_ShapeCollection.__next_shape_id value is correct"""
        # setup ------------------------
        shapes = _sldLayout1_shapes()
        # exercise ---------------------
        next_id = shapes._ShapeCollection__next_shape_id
        # verify -----------------------
        expected = 4
        actual = next_id
        msg = "expected %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
