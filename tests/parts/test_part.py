# encoding: utf-8

"""
Test suite for pptx.part module
"""

from __future__ import absolute_import

from mock import Mock

from pptx.opc.packuri import PackURI
from pptx.parts.part import PartCollection

from ..unitutil.legacy import TestCase


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
