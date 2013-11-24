# encoding: utf-8

"""Test suite for pptx.part module."""

from __future__ import absolute_import

from mock import Mock

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.parts.part import BasePart, PartCollection

from ..unitutil import TestCase


class TestBasePart(TestCase):
    """Test BasePart"""
    def setUp(self):
        partname = PackURI('/foo/bar.xml')
        self.basepart = BasePart(partname, None, None)
        self.cls = BasePart

    def test__add_relationship_adds_specified_relationship(self):
        """BasePart._add_relationship adds specified relationship"""
        # setup ------------------------
        reltype = RT.IMAGE
        target_part = Mock(name='image')
        rId = 'rId1'
        # exercise ---------------------
        rel = self.basepart._add_relationship(reltype, target_part, rId)
        # verify -----------------------
        expected = (rId, reltype, target_part)
        actual = (rel.rId, rel.reltype, rel.target_part)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    # def test__add_relationship_reuses_matching_relationship(self):
    #     """BasePart._add_relationship reuses matching relationship"""
    #     # setup ------------------------
    #     reltype = RT.IMAGE
    #     target = Mock(name='image')
    #     # exercise ---------------------
    #     rel1 = self.basepart._add_relationship(reltype, target)
    #     rel2 = self.basepart._add_relationship(reltype, target)
    #     # verify -----------------------
    #     assert_that(rel1, is_(equal_to(rel2)))

    def test__blob_value_for_binary_part(self):
        """BasePart._blob value is correct for binary part"""
        # setup ------------------------
        partname = PackURI('/docProps/thumbnail.jpeg')
        blob = '0123456789'
        basepart = BasePart(partname, None, blob)
        # exercise ---------------------
        retval = basepart.blob
        # verify -----------------------
        expected = blob
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_partname_setter(self):
        """BasePart.partname setter stores passed value"""
        # setup ------------------------
        partname = '/ppt/presentation.xml'
        # exercise ----------------
        self.basepart.partname = partname
        # verify ------------------
        expected = partname
        actual = self.basepart.partname
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class TestPartCollection(TestCase):

    def test_add_part_preserves_sort_order(self):
        partname1 = PackURI('/ppt/slides/slide1.xml')
        partname2 = PackURI('/ppt/slides/slide2.xml')
        partname3 = PackURI('/ppt/slides/slide3.xml')
        part1 = Mock(name='part1')
        part1.partname = partname1
        part2 = Mock(name='part2')
        part2.partname = partname2
        part3 = Mock(name='part3')
        part3.partname = partname3
        parts = PartCollection()
        # exercise ---------------------
        parts.add_part(part2)
        parts.add_part(part3)
        parts.add_part(part1)
        # verify -----------------------
        expected = [partname1, partname2, partname3]
        actual = [part.partname for part in parts]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
