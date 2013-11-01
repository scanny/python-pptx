# encoding: utf-8

"""Test suite for pptx.part module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_
from mock import Mock

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml_bytes
from pptx.parts.part import BasePart, PartCollection

from ..unitutil import serialize_xml, TestCase


class TestBasePart(TestCase):
    """Test BasePart"""
    def setUp(self):
        self.basepart = BasePart()
        self.cls = BasePart

    def test__add_relationship_adds_specified_relationship(self):
        """BasePart._add_relationship adds specified relationship"""
        # setup ------------------------
        reltype = RT.IMAGE
        target = Mock(name='image')
        # exercise ---------------------
        rel = self.basepart._add_relationship(reltype, target)
        # verify -----------------------
        expected = ('rId1', reltype, target)
        actual = (rel._rId, rel._reltype, rel.target)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__add_relationship_reuses_matching_relationship(self):
        """BasePart._add_relationship reuses matching relationship"""
        # setup ------------------------
        reltype = RT.IMAGE
        target = Mock(name='image')
        # exercise ---------------------
        rel1 = self.basepart._add_relationship(reltype, target)
        rel2 = self.basepart._add_relationship(reltype, target)
        # verify -----------------------
        assert_that(rel1, is_(equal_to(rel2)))

    def test__blob_value_for_binary_part(self):
        """BasePart._blob value is correct for binary part"""
        # setup ------------------------
        blob = '0123456789'
        self.basepart._load_blob = blob
        self.basepart.partname = '/docProps/thumbnail.jpeg'
        # exercise ---------------------
        retval = self.basepart._blob
        # verify -----------------------
        expected = blob
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__blob_value_for_xml_part(self):
        """BasePart._blob value is correct for XML part"""
        # setup ------------------------
        elm = parse_xml_bytes('<root><elm1 attr="one"/></root>')
        self.basepart._element = elm
        self.basepart.partname = '/ppt/presentation.xml'
        # exercise ---------------------
        retval = self.basepart._blob
        # verify -----------------------
        expected = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<root><elm1 attr="one"/></root>'
        )
        actual = retval
        msg = "expected: \n'%s'\n, got \n'%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__load_sets__element_for_xml_part(self):
        """BasePart._load() sets _element for xml part"""
        # setup ------------------------
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/presentation.xml'
        pkgpart.blob = '<root><elm1   attr="spam"/></root>'
        pkgpart.relationships = []
        part_dict = {}
        part = self.basepart._load(pkgpart, part_dict)
        # exercise ---------------------
        elm = part._element
        # verify -----------------------
        expected = '<root><elm1 attr="spam"/></root>'
        actual = serialize_xml(elm)
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
    """Test PartCollection"""
    def test__loadpart_sorts_loaded_parts(self):
        """PartCollection._loadpart sorts loaded parts"""
        # setup ------------------------
        partname1 = '/ppt/slides/slide1.xml'
        partname2 = '/ppt/slides/slide2.xml'
        partname3 = '/ppt/slides/slide3.xml'
        part1 = Mock(name='part1')
        part1.partname = partname1
        part2 = Mock(name='part2')
        part2.partname = partname2
        part3 = Mock(name='part3')
        part3.partname = partname3
        parts = PartCollection()
        # exercise ---------------------
        parts._loadpart(part2)
        parts._loadpart(part3)
        parts._loadpart(part1)
        # verify -----------------------
        expected = [partname1, partname2, partname3]
        actual = [part.partname for part in parts]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
