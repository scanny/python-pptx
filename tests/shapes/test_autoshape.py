# encoding: utf-8

"""Test suite for pptx.autoshape module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_, is_not
from mock import Mock, patch

from pptx.constants import MSO_AUTO_SHAPE_TYPE as MAST, MSO
from pptx.shapes.autoshape import (
    _Adjustment, _AdjustmentCollection, _AutoShapeType, _Shape
)
from pptx.oxml import oxml_fromstring

from ..unitdata import a_prstGeom, test_shapes
from ..unitutil import TestCase


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


class Test_Shape(TestCase):
    """Test _Shape"""
    @patch('pptx.shapes.autoshape._BaseShape.__init__')
    @patch('pptx.shapes.autoshape._AdjustmentCollection')
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
