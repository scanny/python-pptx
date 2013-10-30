# encoding: utf-8

"""Test suite for pptx.placeholder module."""

from __future__ import absolute_import

from pptx.oxml.ns import namespaces
from pptx.shapes.placeholder import Placeholder
from pptx.shapes.shapetree import ShapeCollection
from pptx.spec import (
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL, PH_ORIENT_HORZ,
    PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER
)

from ..unitutil import absjoin, parse_xml_file, TestCase, test_file_dir

nsmap = namespaces('a', 'r', 'p')


def _sldLayout1():
    path = absjoin(test_file_dir, 'slideLayout1.xml')
    sldLayout = parse_xml_file(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = ShapeCollection(spTree)
    return shapes


class Test_Placeholder(TestCase):
    """Test Placeholder"""
    def test_property_values(self):
        """Placeholder property values are correct"""
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
            ph = Placeholder(sp)
            values = (ph.type, ph.orient, ph.sz, ph.idx)
            # verify ----------------------
            expected = expected_values[idx]
            actual = values
            msg = ("expected shapes[%d] values %s, got %s"
                   % (idx, expected, actual))
            self.assertEqual(expected, actual, msg)
