# encoding: utf-8

"""Test suite for pptx.part module."""

from mock import Mock

from pptx.part import _PartCollection

from testing import TestCase


class Test_PartCollection(TestCase):
    """Test _PartCollection"""
    def test__loadpart_sorts_loaded_parts(self):
        """_PartCollection._loadpart sorts loaded parts"""
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
        parts = _PartCollection()
        # exercise ---------------------
        parts._loadpart(part2)
        parts._loadpart(part3)
        parts._loadpart(part1)
        # verify -----------------------
        expected = [partname1, partname2, partname3]
        actual = [part.partname for part in parts]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
