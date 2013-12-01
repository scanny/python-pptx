# encoding: utf-8

"""Test suite for pptx.autoshape module."""

from __future__ import absolute_import

import pytest

from hamcrest import assert_that, equal_to, is_, is_not
from mock import Mock

from pptx.constants import MSO_AUTO_SHAPE_TYPE as MAST, MSO
from pptx.shapes.autoshape import (
    Adjustment, AdjustmentCollection, AutoShapeType, Shape
)
from pptx.oxml import parse_xml_bytes
from pptx.oxml.autoshape import CT_Shape

from ..oxml.unitdata.autoshape import a_prstGeom
from ..unitutil import class_mock, instance_mock, property_mock, TestCase


class DescribeAdjustment(object):

    def it_knows_its_effective_value(self, effective_val_fixture_):
        adjustment, expected_effective_value = effective_val_fixture_
        assert adjustment.effective_value == expected_effective_value

    # fixture --------------------------------------------------------

    def _effective_adj_val_cases():
        return [
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
        ]

    @pytest.fixture(params=_effective_adj_val_cases())
    def effective_val_fixture_(self, request):
        name = None
        def_val, actual, expected_effective_value = request.param
        adjustment = Adjustment(name, def_val, actual)
        return adjustment, expected_effective_value


class DescribeAdjustmentCollection(object):

    def it_should_load_default_adjustment_values(self, prstGeom_cases_):
        prstGeom, prst, expected = prstGeom_cases_
        adjustments = AdjustmentCollection(prstGeom)._adjustments
        actuals = tuple([(adj.name, adj.def_val) for adj in adjustments])
        reason = ("\n     failed for prst: '%s'" % prst)
        assert len(adjustments) == len(expected), reason
        assert actuals == expected, reason

    def it_should_load_adj_val_actuals_from_xml(
            self, load_adj_actuals_fixture_):
        prstGeom, expected_actuals, prstGeom_xml = load_adj_actuals_fixture_
        adjustments = AdjustmentCollection(prstGeom)._adjustments
        # trace ------------------------
        print("\nfailed on case: '%s'\n\nXML:\n%s" %
              (prstGeom.prst, prstGeom_xml))
        # verify -----------------------
        actual_actuals = dict([(a.name, a.actual) for a in adjustments])
        assert actual_actuals == expected_actuals

    # fixture --------------------------------------------------------

    def _adj_actuals_cases():
        return [
            # no adjustment values in xml or spec
            (a_prstGeom('rect').with_avLst, ()),
            # no adjustment values in xml, but some in spec
            (a_prstGeom('circularArrow'),
             (('adj1', None), ('adj2', None), ('adj3', None), ('adj4', None),
              ('adj5', None))),
            # adjustment value in xml but none in spec
            (a_prstGeom('rect').with_gd(), ()),
            # middle adjustment value in xml
            (a_prstGeom('mathDivide').with_gd(name='adj2'),
             (('adj1', None), ('adj2', 25000), ('adj3', None))),
            # all adjustment values in xml
            (a_prstGeom('wedgeRoundRectCallout')
                .with_gd(111, 'adj1')
                .with_gd(222, 'adj2')
                .with_gd(333, 'adj3'),
             (('adj1', 111), ('adj2', 222), ('adj3', 333))),
        ]

    @pytest.fixture(params=_adj_actuals_cases())
    def load_adj_actuals_fixture_(self, request):
        prstGeom_bldr, adj_vals = request.param
        prstGeom = prstGeom_bldr.element
        expected = dict(adj_vals)
        prstGeom_xml = prstGeom_bldr.xml
        return prstGeom, expected, prstGeom_xml

    def _prstGeom_cases():
        return [
            # rect has no adjustments
            ('rect', ()),
            # chevron has one simple one
            ('chevron', (('adj', 50000),)),
            # one with several and some negative
            ('accentBorderCallout1',
             (('adj1', 18750), ('adj2', -8333), ('adj3', 112500),
              ('adj4', -38333))),
            # another one with some negative
            ('wedgeRoundRectCallout',
             (('adj1', -20833), ('adj2', 62500), ('adj3', 16667))),
            # one with values outside normal range
            ('circularArrow',
             (('adj1', 12500), ('adj2', 1142319), ('adj3', 20457681),
              ('adj4', 10800000), ('adj5', 12500))),
        ]

    @pytest.fixture(params=_prstGeom_cases())
    def prstGeom_cases_(self, request):
        prst, expected_values = request.param
        prstGeom = a_prstGeom(prst).with_avLst.element
        return prstGeom, prst, expected_values


class TestAdjustmentCollection(TestCase):

    def test_it_should_return_effective_value_on_indexed_access(self):
        """AdjustmentCollection[n] is normalized effective value of nth"""
        # setup ------------------------
        cases = (
            ('rect', ()),
            ('chevron', (0.5,)),
            ('circularArrow', (0.125, 11.42319, 204.57681, 108.0, 0.125)),
        )
        for prst, expected_values in cases:
            prstGeom = a_prstGeom(prst).element
            # exercise -----------------
            adjustments = AdjustmentCollection(prstGeom)
            # verify -------------------
            reason = "failed on case: prst=='%s'" % prst
            assert_that(len(adjustments), is_(len(expected_values)), reason)
            retvals = tuple([adj for adj in adjustments])
            assert_that(retvals, is_(equal_to(expected_values)), reason)

    def test_it_should_update_actual_value_on_indexed_assignment(self):
        """Assignment to AdjustmentCollection[n] updates nth actual"""
        # setup ------------------------
        cases = (
            ('chevron', 0, 0.5, 50000),
            ('circularArrow', 2, 99.99, 9999000),
        )
        prst = 'chevron'
        for prst, idx, new_value, expected in cases:
            prstGeom = a_prstGeom(prst).element
            adjustments = AdjustmentCollection(prstGeom)
            # exercise -----------------
            adjustments[idx] = new_value
            # verify -------------------
            reason = "failed on case: prst=='%s'" % prst
            assert_that(adjustments._adjustments[idx].actual,
                        is_(equal_to(expected)),
                        reason)

    def test_it_should_round_trip_indexed_assignment(self):
        """Assignment to AdjustmentCollection[n] round-trips"""
        # setup ------------------------
        new_value = 0.375
        prstGeom = a_prstGeom('chevron').element
        adjustments = AdjustmentCollection(prstGeom)
        assert_that(adjustments[0], is_not(equal_to(new_value)))
        # exercise ---------------------
        adjustments[0] = new_value
        # verify -----------------------
        assert_that(adjustments[0], is_(equal_to(new_value)))

    def test_it_should_raise_on_bad_key(self):
        """AdjustmentCollection[idx] raises on invalid idx"""
        # setup ------------------------
        prstGeom = a_prstGeom('chevron').element
        adjustments = AdjustmentCollection(prstGeom)
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
        """AdjustmentCollection[n] = val raises on val is not number"""
        # setup ------------------------
        prstGeom = a_prstGeom('chevron').element
        adjustments = AdjustmentCollection(prstGeom)
        # verify -----------------------
        with self.assertRaises(ValueError):
            adjustments[0] = 'foobar'

    def test_writes_adj_vals_to_xml_on_assignment(self):
        """AdjustmentCollection writes adj vals to XML on assignment"""
        # setup ------------------------
        prstGeom = a_prstGeom('chevron').element
        adjustments = AdjustmentCollection(prstGeom)
        _prstGeom = Mock(name='_prstGeom')
        adjustments._prstGeom = _prstGeom
        # exercise ---------------------
        adjustments[0] = 0.999
        # verify -----------------------
        assert_that(_prstGeom.rewrite_guides.call_count, is_(1))


class TestAutoShapeType(TestCase):
    """Test AutoShapeType"""
    def test_construction_return_values(self):
        """AutoShapeType() returns instance with correct property values"""
        # setup ------------------------
        id_ = MAST.ROUNDED_RECTANGLE
        prst = 'roundRect'
        basename = 'Rounded Rectangle'
        # exercise ---------------------
        autoshape_type = AutoShapeType(id_)
        # verify -----------------------
        assert_that(autoshape_type.autoshape_type_id, is_(equal_to(id_)))
        assert_that(autoshape_type.prst, is_(equal_to(prst)))
        assert_that(autoshape_type.basename, is_(equal_to(basename)))

    def test_default_adjustment_values_return_value(self):
        """AutoShapeType.default_adjustment_values() return val correct"""
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
            def_adj_vals = AutoShapeType.default_adjustment_values(prst)
            assert_that(def_adj_vals, is_(equal_to(expected_vals)))

    def test_from_prst_return_value(self):
        """AutoShapeType.from_prst() return value is correct"""
        # setup ------------------------
        autoshape_type_id = MAST.ROUNDED_RECTANGLE
        prst = 'roundRect'
        # exercise ---------------------
        retval = AutoShapeType.from_prst(prst)
        # verify -----------------------
        assert_that(retval, is_(equal_to(autoshape_type_id)))

    def test_from_prst_raises_on_bad_prst(self):
        # setup ------------------------
        prst = 'badPrst'
        # verify -----------------------
        with self.assertRaises(KeyError):
            AutoShapeType.from_prst(prst)

    def test_second_construction_returns_cached_instance(self):
        """AutoShapeType() returns cached instance on duplicate call"""
        # setup ------------------------
        id_ = MAST.ROUNDED_RECTANGLE
        ast1 = AutoShapeType(id_)
        # exercise ---------------------
        ast2 = AutoShapeType(id_)
        # verify -----------------------
        assert_that(ast2, is_(equal_to(ast1)))

    def test_construction_raises_on_bad_autoshape_type_id(self):
        """AutoShapeType() raises on bad auto shape type id"""
        # setup ------------------------
        autoshape_type_id = 9999
        # verify -----------------------
        with self.assertRaises(KeyError):
            AutoShapeType(autoshape_type_id)


class DescribeShape(object):

    def it_initializes_adjustments_on_first_ref(self, init_adjs_fixture_):
        shape, adjs_, AdjustmentCollection_, sp_ = init_adjs_fixture_
        assert shape.adjustments is adjs_
        AdjustmentCollection_.assert_called_once_with(sp_.prstGeom)

    def it_knows_its_autoshape_type(self, autoshape_type_fixture_):
        shape, autoshape_type, AutoShapeType_, prst = autoshape_type_fixture_
        assert shape.auto_shape_type == autoshape_type
        AutoShapeType_.from_prst.assert_called_once_with(prst)

    def it_raises_when_auto_shape_type_called_on_non_autoshape(
            self, non_autoshape_shape_):
        with pytest.raises(ValueError):
            non_autoshape_shape_.auto_shape_type

    def it_knows_its_shape_type_when_its_a_placeholder(
            self, placeholder_shape_):
        assert placeholder_shape_.shape_type == MSO.PLACEHOLDER

    def it_knows_its_shape_type_when_its_not_a_placeholder(
            self, non_placeholder_shapes_):
        autoshape_shape_, textbox_shape_ = non_placeholder_shapes_
        assert autoshape_shape_.shape_type == MSO.AUTO_SHAPE
        assert textbox_shape_.shape_type == MSO.TEXT_BOX

    def it_raises_when_shape_type_called_on_unrecognized_shape(self):
        xml = (
            '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/'
            '2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/'
            '2006/main"><p:nvSpPr><p:cNvPr id="9" name="Unknown Shape Type 8"'
            '/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>'
        )
        sp = parse_xml_bytes(xml)
        shape = Shape(sp, None)
        # verify -----------------------
        with pytest.raises(NotImplementedError):
            shape.shape_type

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def autoshape_type_fixture_(
            self, request, sp_, autoshape_type, AutoShapeType_, prst):
        shape = Shape(sp_, None)
        return shape, autoshape_type, AutoShapeType_, prst

    @pytest.fixture
    def init_adjs_fixture_(
            self, request, sp_, adjustments_, AdjustmentCollection_):
        shape = Shape(sp_, None)
        return shape, adjustments_, AdjustmentCollection_, sp_

    @pytest.fixture
    def AdjustmentCollection_(self, request, adjustments_):
        return class_mock(
            request, 'pptx.shapes.autoshape.AdjustmentCollection',
            return_value=adjustments_
        )

    @pytest.fixture
    def adjustments_(self, request):
        return instance_mock(request, AdjustmentCollection)

    @pytest.fixture
    def AutoShapeType_(self, request, autoshape_type):
        AutoShapeType_ = class_mock(
            request, 'pptx.shapes.autoshape.AutoShapeType',
            return_value=autoshape_type
        )
        AutoShapeType_.from_prst.return_value = autoshape_type
        return AutoShapeType_

    @pytest.fixture
    def autoshape_type(self):
        return 66

    @pytest.fixture
    def non_autoshape_shape_(self, request, sp_):
        sp_.is_autoshape = False
        return Shape(sp_, None)

    @pytest.fixture
    def non_placeholder_shapes_(self, request):
        autoshape_sp_ = instance_mock(
            request, CT_Shape, name='autoshape_sp_', is_autoshape=True,
            is_textbox=False
        )
        autoshape_shape_ = Shape(autoshape_sp_, None)
        textbox_sp_ = instance_mock(
            request, CT_Shape, name='textbox_sp_', is_autoshape=False,
            is_textbox=True
        )
        textbox_shape_ = Shape(textbox_sp_, None)
        property_mock(request, Shape, 'is_placeholder', return_value=False)
        return autoshape_shape_, textbox_shape_

    @pytest.fixture
    def placeholder_shape_(self, request, sp_):
        placeholder_shape_ = Shape(sp_, None)
        property_mock(request, Shape, 'is_placeholder', return_value=True)
        return placeholder_shape_

    @pytest.fixture
    def prst(self):
        return 'foobar'

    @pytest.fixture
    def sp_(self, request, prst):
        return instance_mock(request, CT_Shape, prst=prst, is_autoshape=True)
