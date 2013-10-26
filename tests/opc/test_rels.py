# encoding: utf-8

"""Test suite for pptx.part module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_, same_instance
from mock import Mock

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.rels import Relationship, RelationshipCollection
from pptx.parts.part import BasePart

from ..unitutil import TestCase


class PartBuilder(object):
    """Builder class for test Parts"""
    def __init__(self):
        self.partname = '/ppt/slides/slide1.xml'

    def with_partname(self, partname):
        self.partname = partname
        return self

    def build(self):
        p = BasePart()
        p.partname = self.partname
        return p


class TestRelationship(TestCase):
    """Test Relationship"""
    def setUp(self):
        rId = 'rId1'
        reltype = RT.SLIDE
        target_part = None
        self.rel = Relationship(rId, reltype, target_part)

    def test_constructor_raises_on_bad_rId(self):
        """Relationship constructor raises on non-standard rId"""
        with self.assertRaises(AssertionError):
            Relationship('Non-std14', None, None)

    def test__num_value(self):
        """Relationship._num value is correct"""
        # setup ------------------------
        num = 91
        rId = 'rId%d' % num
        rel = Relationship(rId, None, None)
        # verify -----------------------
        assert_that(rel._num, is_(equal_to(num)))

    def test__num_value_on_non_standard_rId(self):
        """Relationship._num value is correct for non-standard rId"""
        # setup ------------------------
        rel = Relationship('rIdSm', None, None)
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


class TestRelationshipCollection(TestCase):
    """Test RelationshipCollection"""
    def setUp(self):
        self.relationships = RelationshipCollection()

    def _reltype_ordering_mock(self):
        """
        Return RelationshipCollection instance with mocked-up contents
        suitable for testing _reltype_ordering.
        """
        # setup ------------------------
        partnames = ['/ppt/slides/slide4.xml',
                     '/ppt/slideLayouts/slideLayout1.xml',
                     '/ppt/slideMasters/slideMaster1.xml',
                     '/ppt/slides/slide1.xml',
                     '/ppt/presProps.xml']
        part1 = Mock(name='part1')
        part1.partname = partnames[0]
        part2 = Mock(name='part2')
        part2.partname = partnames[1]
        part3 = Mock(name='part3')
        part3.partname = partnames[2]
        part4 = Mock(name='part4')
        part4.partname = partnames[3]
        part5 = Mock(name='part5')
        part5.partname = partnames[4]
        rel1 = Relationship('rId1', RT.SLIDE,        part1)
        rel2 = Relationship('rId2', RT.SLIDE_LAYOUT, part2)
        rel3 = Relationship('rId3', RT.SLIDE_MASTER, part3)
        rel4 = Relationship('rId4', RT.SLIDE,        part4)
        rel5 = Relationship('rId5', RT.PRES_PROPS,   part5)
        relationships = RelationshipCollection()
        relationships._additem(rel1)
        relationships._additem(rel2)
        relationships._additem(rel3)
        relationships._additem(rel4)
        relationships._additem(rel5)
        return (relationships, partnames)

    def test_it_can_find_related_part(self):
        """RelationshipCollection can find related part"""
        # setup ------------------------
        reltype = RT.CORE_PROPERTIES
        part = Mock(name='part')
        relationship = Relationship('rId1', reltype, part)
        relationships = RelationshipCollection()
        relationships._additem(relationship)
        # exercise ---------------------
        retval = relationships.related_part(reltype)
        # verify -----------------------
        assert_that(retval, same_instance(part))

    def test_it_raises_if_it_cant_find_a_related_part(self):
        """RelationshipCollection raises if it can't find a related part"""
        # setup ------------------------
        relationships = RelationshipCollection()
        # exercise ---------------------
        with self.assertRaises(KeyError):
            relationships.related_part('foobar')

    def test__additem_raises_on_dup_rId(self):
        """RelationshipCollection._additem raises on duplicate rId"""
        # setup ------------------------
        part1 = BasePart()
        part2 = BasePart()
        rel1 = Relationship('rId9', None, part1)
        rel2 = Relationship('rId9', None, part2)
        self.relationships._additem(rel1)
        # verify -----------------------
        with self.assertRaises(ValueError):
            self.relationships._additem(rel2)

    def test__additem_maintains_rId_ordering(self):
        """RelationshipCollection maintains rId ordering on additem()"""
        # setup ------------------------
        part1 = BasePart()
        part2 = BasePart()
        part3 = BasePart()
        rel1 = Relationship('rId1', None, part1)
        rel2 = Relationship('rId2', None, part2)
        rel3 = Relationship('rId3', None, part3)
        # exercise ---------------------
        self.relationships._additem(rel2)
        self.relationships._additem(rel1)
        self.relationships._additem(rel3)
        # verify -----------------------
        expected = ['rId1', 'rId2', 'rId3']
        actual = [rel._rId for rel in self.relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__additem_maintains_reltype_ordering(self):
        """RelationshipCollection maintains reltype ordering on additem()"""
        # setup ------------------------
        relationships, partnames = self._reltype_ordering_mock()
        ordering = (RT.SLIDE_MASTER, RT.SLIDE_LAYOUT, RT.SLIDE)
        relationships._reltype_ordering = ordering
        partname = '/ppt/slides/slide2.xml'
        part = Mock(name='new_part')
        part.partname = partname
        rId = relationships._next_rId
        rel = Relationship(rId, RT.SLIDE, part)
        # exercise ---------------------
        relationships._additem(rel)
        # verify ordering -------------
        expected = [partnames[2], partnames[1], partnames[3],
                    partname, partnames[0], partnames[4]]
        actual = [r._target.partname for r in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_rels_of_reltype_return_value(self):
        """RelationshipCollection._rels_of_reltype returns correct rels"""
        # setup ------------------------
        relationships, partnames = self._reltype_ordering_mock()
        # exercise ---------------------
        retval = relationships.rels_of_reltype(RT.SLIDE)
        # verify ordering -------------
        expected = ['rId1', 'rId4']
        actual = [rel._rId for rel in retval]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__reltype_ordering_sorts_rels(self):
        """RelationshipCollection._reltype_ordering sorts rels"""
        # setup ------------------------
        relationships, partnames = self._reltype_ordering_mock()
        ordering = (RT.SLIDE_MASTER, RT.SLIDE_LAYOUT, RT.SLIDE)
        # exercise ---------------------
        relationships._reltype_ordering = ordering
        # verify ordering -------------
        assert_that(relationships._reltype_ordering, is_(equal_to(ordering)))
        expected = [partnames[2], partnames[1], partnames[3], partnames[0],
                    partnames[4]]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__reltype_ordering_renumbers_rels(self):
        """RelationshipCollection._reltype_ordering renumbers rels"""
        # setup ------------------------
        relationships, partnames = self._reltype_ordering_mock()
        ordering = (RT.SLIDE_MASTER, RT.SLIDE_LAYOUT, RT.SLIDE)
        # exercise ---------------------
        relationships._reltype_ordering = ordering
        # verify renumbering ----------
        expected = ['rId1', 'rId2', 'rId3', 'rId4', 'rId5']
        actual = [rel._rId for rel in relationships]
        msg = "expected numbering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__next_rId_fills_gap(self):
        """RelationshipCollection._next_rId fills gap in rId sequence"""
        # setup ------------------------
        part1 = BasePart()
        part2 = BasePart()
        part3 = BasePart()
        part4 = BasePart()
        rel1 = Relationship('rId1', None, part1)
        rel2 = Relationship('rId2', None, part2)
        rel3 = Relationship('rId3', None, part3)
        rel4 = Relationship('rId4', None, part4)
        cases = (('rId1', (rel2, rel3, rel4)),
                 ('rId2', (rel1, rel3, rel4)),
                 ('rId3', (rel1, rel2, rel4)),
                 ('rId4', (rel1, rel2, rel3)))
        # exercise ---------------------
        expected_rIds = []
        actual_rIds = []
        for expected_rId, rels in cases:
            expected_rIds.append(expected_rId)
            relationships = RelationshipCollection()
            for rel in rels:
                relationships._additem(rel)
            actual_rIds.append(relationships._next_rId)
        # verify -----------------------
        expected = expected_rIds
        actual = actual_rIds
        msg = "expected rIds %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_reorders_on_partname_change(self):
        """RelationshipCollection reorders on partname change"""
        # setup ------------------------
        partname1 = '/ppt/slides/slide1.xml'
        partname2 = '/ppt/slides/slide2.xml'
        partname3 = '/ppt/slides/slide3.xml'
        part1 = PartBuilder().with_partname(partname1).build()
        part2 = PartBuilder().with_partname(partname2).build()
        rel1 = Relationship('rId1', RT.SLIDE, part1)
        rel2 = Relationship('rId2', RT.SLIDE, part2)
        relationships = RelationshipCollection()
        relationships._reltype_ordering = (RT.SLIDE)
        relationships._additem(rel1)
        relationships._additem(rel2)
        # exercise ---------------------
        part1.partname = partname3
        # verify -----------------------
        expected = [partname2, partname3]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
