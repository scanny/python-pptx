# encoding: utf-8

"""Test suite for pptx.oxml.table module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_

from pptx.constants import TEXT_ANCHORING_TYPE as TANC
from pptx.oxml.ns import nsdecls
from pptx.oxml.table import CT_Table

from ..unitdata import a_tbl, test_table_elements, test_table_xml
from ..unitutil import TestCase


class TestCT_Table(TestCase):
    """Test CT_Table"""
    boolprops = ('bandRow', 'firstRow', 'lastRow',
                 'bandCol', 'firstCol', 'lastCol')

    def test_new_tbl_generates_correct_xml(self):
        """CT_Table._new_tbl() returns correct XML"""
        # setup ------------------------
        rows, cols = 2, 3
        width, height = 334, 445
        xml = (
            '<a:tbl %s>\n  <a:tblPr firstRow="1" bandRow="1">\n    <a:tableSt'
            'yleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>\n '
            ' </a:tblPr>\n  <a:tblGrid>\n    <a:gridCol w="111"/>\n    <a:gri'
            'dCol w="111"/>\n    <a:gridCol w="112"/>\n  </a:tblGrid>\n  <a:t'
            'r h="222">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n '
            '       <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <'
            'a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:'
            'bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBod'
            'y>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n'
            '        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n    '
            '  </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n  <a:tr '
            'h="223">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n   '
            '     <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:'
            'tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bo'
            'dyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>'
            '\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n  '
            '      <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      '
            '</a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n</a:tbl>\n'
            % nsdecls('a')
        )
        # exercise ---------------------
        tbl = CT_Table.new_tbl(rows, cols, width, height)
        # verify -----------------------
        self.assertEqualLineByLine(xml, tbl)

    def test_boolean_property_value_is_correct(self):
        """CT_Table boolean property value is correct"""
        def getter_cases(propname):
            """Test cases for boolean property getter tests"""
            return (
                # defaults to False if no tblPr element present
                (a_tbl(), False),
                # defaults to False if tblPr element is empty
                (a_tbl().with_tblPr, False),
                # returns True if firstCol is valid truthy value
                (a_tbl().with_prop(propname, '1'), True),
                (a_tbl().with_prop(propname, 'true'), True),
                # returns False if firstCol has valid False value
                (a_tbl().with_prop(propname, '0'), False),
                (a_tbl().with_prop(propname, 'false'), False),
                # returns False if firstCol is not valid xsd:boolean value
                (a_tbl().with_prop(propname, 'foobar'), False),
            )
        for propname in self.boolprops:
            cases = getter_cases(propname)
            for tbl_builder, expected_property_value in cases:
                reason = (
                    'tbl.%s did not return %s for this XML:\n\n%s' %
                    (propname, expected_property_value, tbl_builder.xml)
                )
                assert_that(
                    getattr(tbl_builder.element, propname),
                    is_(equal_to(expected_property_value)),
                    reason
                )

    def test_assignment_to_boolean_property_produces_correct_xml(self):
        """Assignment to boolean property of CT_Table produces correct XML"""
        def xml_check_cases(propname):
            return (
                # => True: tblPr and attribute should be added
                (a_tbl(), True, a_tbl().with_prop(propname, '1')),
                # => False: attribute should be removed if false
                (a_tbl().with_prop(propname, '1'), False, a_tbl().with_tblPr),
                # => False: attribute should not be added if false
                (a_tbl(), False, a_tbl()),
            )
        for propname in self.boolprops:
            cases = xml_check_cases(propname)
            for tc_builder, assigned_value, expected_tc_builder in cases:
                tc = tc_builder.element
                setattr(tc, propname, assigned_value)
                self.assertEqualLineByLine(expected_tc_builder.xml, tc)


class TestCT_TableCell(TestCase):
    """Test CT_TableCell"""
    def test_anchor_property_value_is_correct(self):
        """CT_TableCell.anchor property value is correct"""
        # setup ------------------------
        cases = (
            (test_table_elements.cell, None),
            (test_table_elements.top_aligned_cell, TANC.TOP)
        )
        # verify -----------------------
        for tc, expected_text_anchoring_type in cases:
            assert_that(tc.anchor,
                        is_(equal_to(expected_text_anchoring_type)))

    def test_assignment_to_anchor_sets_anchor_value(self):
        """Assignment to CT_TableCell.anchor sets anchor value"""
        # setup ------------------------
        cases = (
            # something => something else
            (test_table_elements.top_aligned_cell, TANC.MIDDLE),
            # something => None
            (test_table_elements.top_aligned_cell, None),
            # None => something
            (test_table_elements.cell, TANC.BOTTOM),
            # None => None
            (test_table_elements.cell, None)
        )
        # verify -----------------------
        for tc, anchor in cases:
            tc.anchor = anchor
            assert_that(tc.anchor, is_(equal_to(anchor)))

    def test_assignment_to_anchor_produces_correct_xml(self):
        """Assigning value to CT_TableCell.anchor produces correct XML"""
        # setup ------------------------
        cases = (
            # None => something
            (test_table_elements.cell, TANC.TOP,
             test_table_xml.top_aligned_cell),
            # something => None
            (test_table_elements.top_aligned_cell, None,
             test_table_xml.cell)
        )
        # verify -----------------------
        for tc, text_anchoring_type, expected_xml in cases:
            tc.anchor = text_anchoring_type
            self.assertEqualLineByLine(expected_xml, tc)

    def test_marX_property_values_are_correct(self):
        """CT_TableCell.marX property values are correct"""
        # setup ------------------------
        cases = (
            (test_table_elements.cell_with_margins, 12, 34, 56, 78),
            (test_table_elements.cell, 45720, 91440, 45720, 91440)
        )
        # verify -----------------------
        for tc, exp_marT, exp_marR, exp_marB, exp_marL in cases:
            assert_that(tc.marT, is_(equal_to(exp_marT)))
            assert_that(tc.marR, is_(equal_to(exp_marR)))
            assert_that(tc.marB, is_(equal_to(exp_marB)))
            assert_that(tc.marL, is_(equal_to(exp_marL)))

    def test_assignment_to_marX_sets_value(self):
        """Assignment to CT_TableCell.marX sets marX value"""
        # setup ------------------------
        cases = (
            # something => something else
            (
                test_table_elements.cell_with_margins,
                (98, 76, 54, 32),
                (98, 76, 54, 32)
            ),
            # something => None
            (
                test_table_elements.cell_with_margins,
                (None, None, None, None),
                (45720, 91440, 45720, 91440)
            ),
            # None => something
            (
                test_table_elements.cell,
                (98, 76, 54, 32),
                (98, 76, 54, 32)
            ),
            # None => None
            (
                test_table_elements.cell,
                (None, None, None, None),
                (45720, 91440, 45720, 91440)
            )
        )
        # verify -----------------------
        for tc, marX, expected_marX in cases:
            tc.marT, tc.marR, tc.marB, tc.marL = marX
            exp_marT, exp_marR, exp_marB, exp_marL = expected_marX
            assert_that(tc.marT, is_(equal_to(exp_marT)))
            assert_that(tc.marR, is_(equal_to(exp_marR)))
            assert_that(tc.marB, is_(equal_to(exp_marB)))
            assert_that(tc.marL, is_(equal_to(exp_marL)))

    def test_assignment_to_marX_produces_correct_xml(self):
        """Assigning value to CT_TableCell.marX produces correct XML"""
        # setup ------------------------
        cases = (
            # None => something
            (
                test_table_elements.cell,
                (12, 34, 56, 78),
                test_table_xml.cell_with_margins
            ),
            # something => None
            (
                test_table_elements.cell_with_margins,
                (None, None, None, None),
                test_table_xml.cell
            )
        )
        # verify -----------------------
        for tc, marX, expected_xml in cases:
            tc.marT, tc.marR, tc.marB, tc.marL = marX
            self.assertEqualLineByLine(expected_xml, tc)
