# encoding: utf-8

"""Test suite for pptx.part module."""

from hamcrest import assert_that, equal_to, is_

from pptx.rels import _Relationship
from pptx.opc_constants import RELATIONSHIP_TYPE as RT

from testing import TestCase


class Test_Relationship(TestCase):
    """Test _Relationship"""
    def setUp(self):
        rId = 'rId1'
        reltype = RT.SLIDE
        target_part = None
        self.rel = _Relationship(rId, reltype, target_part)

    def test_constructor_raises_on_bad_rId(self):
        """_Relationship constructor raises on non-standard rId"""
        with self.assertRaises(AssertionError):
            _Relationship('Non-std14', None, None)

    def test__num_value(self):
        """_Relationship._num value is correct"""
        # setup ------------------------
        num = 91
        rId = 'rId%d' % num
        rel = _Relationship(rId, None, None)
        # verify -----------------------
        assert_that(rel._num, is_(equal_to(num)))

    def test__num_value_on_non_standard_rId(self):
        """_Relationship._num value is correct for non-standard rId"""
        # setup ------------------------
        rel = _Relationship('rIdSm', None, None)
        # verify -----------------------
        assert_that(rel._num, is_(equal_to(9999)))

    def test__rId_setter(self):
        """Relationship._rId setter stores passed value"""
        # setup ------------------------
        rId = 'rId9'
        # exercise ----------------
        self.rel._rId = rId
        # verify ------------------
        expected = rId
        actual = self.rel._rId
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
