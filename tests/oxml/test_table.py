"""Unit-test suite for pptx.oxml.table module"""

from __future__ import annotations

import pytest

from pptx.oxml.ns import nsdecls
from pptx.oxml.table import CT_Table, TcRange

from ..unitutil.cxml import element, xml


class DescribeCT_Table(object):
    @pytest.mark.parametrize(
        ("tbl_cxml", "expected_value"),
        [
            ("a:tbl/a:tblGrid", None),
            ("a:tbl/(a:tblPr,a:tblGrid)", None),
            (
                'a:tbl/(a:tblPr/a:tableStyleId"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",a:tblGrid)',
                "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
            ),
        ],
    )
    def it_knows_its_tableStyleId(self, tbl_cxml, expected_value):
        tbl = element(tbl_cxml)
        assert tbl.tableStyleId == expected_value

    @pytest.mark.parametrize(
        ("tbl_cxml", "new_value", "expected_cxml"),
        [
            # ---assigning a GUID to a bare tbl adds tblPr and tableStyleId---
            (
                "a:tbl/a:tblGrid",
                "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
                'a:tbl/(a:tblPr/a:tableStyleId"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",a:tblGrid)',
            ),
            # ---assigning a GUID to tblPr without tableStyleId adds the child---
            (
                "a:tbl/(a:tblPr,a:tblGrid)",
                "{NEW-GUID}",
                'a:tbl/(a:tblPr/a:tableStyleId"{NEW-GUID}",a:tblGrid)',
            ),
            # ---assigning a GUID overwrites existing tableStyleId text---
            (
                'a:tbl/(a:tblPr/a:tableStyleId"{OLD-GUID}",a:tblGrid)',
                "{NEW-GUID}",
                'a:tbl/(a:tblPr/a:tableStyleId"{NEW-GUID}",a:tblGrid)',
            ),
            # ---assigning None removes the tableStyleId child (leaves tblPr)---
            (
                'a:tbl/(a:tblPr/a:tableStyleId"{SOME-ID}",a:tblGrid)',
                None,
                "a:tbl/(a:tblPr,a:tblGrid)",
            ),
            # ---assigning None to a tbl without tblPr is a no-op---
            ("a:tbl/a:tblGrid", None, "a:tbl/a:tblGrid"),
        ],
    )
    def it_can_change_its_tableStyleId(self, tbl_cxml, new_value, expected_cxml):
        tbl = element(tbl_cxml)
        tbl.tableStyleId = new_value
        assert tbl.xml == xml(expected_cxml)

    def it_can_create_a_new_tbl_element_tree(self):
        """
        Indirectly tests that column widths are a proportional split of total
        width and that row heights a proportional split of total height.
        """
        expected_xml = (
            '<a:tbl %s>\n  <a:tblPr firstRow="1" bandRow="1">\n    <a:tableSt'
            "yleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>\n "
            ' </a:tblPr>\n  <a:tblGrid>\n    <a:gridCol w="111"/>\n    <a:gri'
            'dCol w="111"/>\n    <a:gridCol w="112"/>\n  </a:tblGrid>\n  <a:t'
            'r h="222">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n '
            "       <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <"
            "a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:"
            "bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBod"
            "y>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n"
            "        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n    "
            "  </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n  <a:tr "
            'h="223">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n   '
            "     <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:"
            "tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bo"
            "dyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>"
            "\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n  "
            "      <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      "
            "</a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n</a:tbl>\n" % nsdecls("a")
        )
        tbl = CT_Table.new_tbl(2, 3, 334, 445)
        assert tbl.xml == expected_xml

    def it_accepts_float_width_and_height_when_creating_a_new_tbl(self):
        """Regression for #288: `new_tbl()` must accept non-integer dimensions."""
        tbl = CT_Table.new_tbl(2, 3, 334.5, 445.5)

        grid_widths = [int(gc.get("w")) for gc in tbl.tblGrid.iterchildren()]
        row_heights = [int(tr.get("h")) for tr in tbl.tr_lst]
        # -- every grid-col width and row-height is an integer string (no trailing ".0"), and the
        # -- pieces sum to the rounded total.
        assert all(gc.get("w").isdigit() for gc in tbl.tblGrid.iterchildren())
        assert all(tr.get("h").isdigit() for tr in tbl.tr_lst)
        assert sum(grid_widths) == round(334.5)
        assert sum(row_heights) == round(445.5)

    def it_provides_access_to_its_tc_elements(self):
        tbl_cxml = "a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))"
        tbl = element(tbl_cxml)
        tcs = tbl.xpath(".//a:tc")

        assert tbl.tc(0, 0) is tcs[0]
        assert tbl.tc(0, 1) is tcs[1]
        assert tbl.tc(1, 0) is tcs[2]
        assert tbl.tc(1, 1) is tcs[3]


class DescribeCT_TableCellProperties(object):
    """Unit-test suite for `pptx.oxml.table.CT_TableCellProperties` border descriptors."""

    @pytest.mark.parametrize(
        ("attr_name", "expected_tag"),
        [
            ("lnL", "a:lnL"),
            ("lnR", "a:lnR"),
            ("lnT", "a:lnT"),
            ("lnB", "a:lnB"),
            ("lnTlToBr", "a:lnTlToBr"),
            ("lnBlToTr", "a:lnBlToTr"),
        ],
    )
    def it_has_border_descriptors_for_each_side(self, attr_name, expected_tag):
        tcPr = element("a:tcPr/%s" % expected_tag)
        border = getattr(tcPr, attr_name)
        assert border is not None
        assert border.tag == (
            "{http://schemas.openxmlformats.org/drawingml/2006/main}" + attr_name
        )

    @pytest.mark.parametrize(
        "attr_name",
        ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"],
    )
    def and_absent_border_is_reported_as_None(self, attr_name):
        tcPr = element("a:tcPr")
        assert getattr(tcPr, attr_name) is None

    def it_inserts_borders_in_schema_order(self):
        """The six border children must appear in the order required by the schema.

        When a consumer adds borders in an arbitrary order, the descriptors must
        place each new child so the final sequence is lnL, lnR, lnT, lnB,
        lnTlToBr, lnBlToTr — otherwise PowerPoint rejects the file.
        """
        tcPr = element("a:tcPr")
        # -- add out of schema order
        tcPr.get_or_add_lnBlToTr()
        tcPr.get_or_add_lnB()
        tcPr.get_or_add_lnL()
        tcPr.get_or_add_lnT()
        tcPr.get_or_add_lnTlToBr()
        tcPr.get_or_add_lnR()
        expected_xml = xml(
            "a:tcPr/(a:lnL,a:lnR,a:lnT,a:lnB,a:lnTlToBr,a:lnBlToTr)"
        )
        assert tcPr.xml == expected_xml

    def it_preserves_fill_element_ordering_when_adding_border(self):
        """A newly-inserted border must precede existing fill/headers children."""
        tcPr = element("a:tcPr/a:solidFill")
        tcPr.get_or_add_lnL()
        # -- lnL must appear before solidFill
        expected_xml = xml("a:tcPr/(a:lnL,a:solidFill)")
        assert tcPr.xml == expected_xml


class DescribeTcRange(object):
    def it_knows_when_the_range_contains_a_merged_cell(self, contains_merge_fixture):
        tc, other_tc, expected_value = contains_merge_fixture
        tc_range = TcRange(tc, other_tc)

        contains_merged_cell = tc_range.contains_merged_cell

        assert contains_merged_cell is expected_value

    def it_knows_how_big_the_merge_range_is(self, dimensions_fixture):
        tc, other_tc, expected_value = dimensions_fixture
        tc_range = TcRange(tc, other_tc)

        dimensions = tc_range.dimensions

        assert dimensions == expected_value

    def it_knows_when_tcs_are_in_the_same_tbl(self, in_same_table_fixture):
        tc, other_tc, expected_value = in_same_table_fixture
        tc_range = TcRange(tc, other_tc)

        in_same_table = tc_range.in_same_table

        assert in_same_table is expected_value

    def it_can_iterate_tcs_not_in_left_col_of_range(self, except_left_fixture):
        tc, other_tc, expected_value = except_left_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_except_left_col_tcs())

        assert tcs == expected_value

    def it_can_iterate_tcs_not_in_top_row_of_range(self, except_top_fixture):
        tc, other_tc, expected_value = except_top_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_except_top_row_tcs())

        assert tcs == expected_value

    def it_can_iterate_left_col_of_range_tcs(self, left_col_fixture):
        tc, other_tc, expected_value = left_col_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_left_col_tcs())

        assert tcs == expected_value

    def it_can_iterate_top_row_of_range_tcs(self, top_row_fixture):
        tc, other_tc, expected_value = top_row_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_top_row_tcs())

        assert tcs == expected_value

    def it_can_migrate_range_content_to_origin_cell(self, move_fixture):
        tc, other_tc, expected_text = move_fixture
        tc_range = TcRange(tc, other_tc)

        tc_range.move_content_to_origin()

        assert tc.text == expected_text
        assert other_tc.text == ""

    @pytest.mark.parametrize(
        ("tbl_cxml", "tc_idxs", "expected_value"),
        [
            # -- 1-col table with multi-row range spans all columns --
            (
                "a:tbl/(a:tblGrid/a:gridCol{w=100},a:tr{h=50}/a:tc,a:tr{h=50}/a:tc)",
                (0, 1),
                True,
            ),
            # -- 2-col table, range 1 col wide in 2-row range is not full-width --
            (
                "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),a:tr{h=50}/(a:tc"
                ",a:tc),a:tr{h=50}/(a:tc,a:tc))",
                (0, 2),
                False,
            ),
            # -- 2-col table with 2-col x 2-row range spans all columns --
            (
                "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),a:tr{h=50}/(a:tc"
                ",a:tc),a:tr{h=50}/(a:tc,a:tc))",
                (0, 3),
                True,
            ),
            # -- single-row range is never a full-column-span (no rows to remove) --
            (
                "a:tbl/(a:tblGrid/a:gridCol{w=100},a:tr{h=50}/a:tc)",
                (0, 0),
                False,
            ),
        ],
    )
    def it_knows_when_range_spans_every_column(
        self, tbl_cxml: str, tc_idxs: tuple[int, int], expected_value: bool
    ):
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc_range = TcRange(tcs[tc_idxs[0]], tcs[tc_idxs[1]])
        assert tc_range.spans_all_columns is expected_value

    @pytest.mark.parametrize(
        ("tbl_cxml", "tc_idxs", "expected_value"),
        [
            # -- 1-row table with multi-col range spans all rows --
            (
                "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),a:tr{h=50}/(a:tc"
                ",a:tc))",
                (0, 1),
                True,
            ),
            # -- 2-row table, range 1 row tall in 2-col range is not full-height --
            (
                "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),a:tr{h=50}/(a:tc"
                ",a:tc),a:tr{h=50}/(a:tc,a:tc))",
                (0, 1),
                False,
            ),
            # -- 2-row table with 2-col x 2-row range spans all rows --
            (
                "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),a:tr{h=50}/(a:tc"
                ",a:tc),a:tr{h=50}/(a:tc,a:tc))",
                (0, 3),
                True,
            ),
            # -- single-column range is never a full-row-span (no cols to remove) --
            (
                "a:tbl/(a:tblGrid/a:gridCol{w=100},a:tr{h=50}/a:tc)",
                (0, 0),
                False,
            ),
        ],
    )
    def it_knows_when_range_spans_every_row(
        self, tbl_cxml: str, tc_idxs: tuple[int, int], expected_value: bool
    ):
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc_range = TcRange(tcs[tc_idxs[0]], tcs[tc_idxs[1]])
        assert tc_range.spans_all_rows is expected_value

    def it_can_collapse_a_full_column_span_range(self):
        """Merged rows are removed and their heights absorbed by the top row."""
        tbl = element(
            "a:tbl/(a:tblGrid/a:gridCol{w=100},a:tr{h=10}/a:tc,a:tr{h=20}/a:tc"
            ",a:tr{h=30}/a:tc,a:tr{h=40}/a:tc)"
        )
        tcs = tbl.xpath("//a:tc")
        tc_range = TcRange(tcs[1], tcs[2])  # rows 1 and 2

        tc_range.collapse_full_column_span()

        # -- original row 2 was removed; row 1 absorbed its height --
        assert [tr.h for tr in tbl.tr_lst] == [10, 50, 40]
        assert len(tbl.tr_lst) == 3

    def it_can_collapse_a_full_row_span_range(self):
        """Merged columns are removed and their widths absorbed by the leftmost column."""
        tbl = element(
            "a:tbl/(a:tblGrid/(a:gridCol{w=10},a:gridCol{w=20},a:gridCol{w=30}"
            ",a:gridCol{w=40}),a:tr{h=50}/(a:tc,a:tc,a:tc,a:tc))"
        )
        tcs = tbl.xpath("//a:tc")
        tc_range = TcRange(tcs[1], tcs[2])  # cols 1 and 2

        tc_range.collapse_full_row_span()

        # -- col 1 absorbed col 2's width; col 2 was removed --
        assert [gc.w for gc in tbl.tblGrid.gridCol_lst] == [10, 50, 40]
        assert len(tbl.tr_lst[0].tc_lst) == 3

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", False),
            ("a:tbl/a:tr/(a:tc{gridSpan=1},a:tc{hMerge=false})", False),
            ("a:tbl/a:tr/(a:tc{gridSpan=2},a:tc{hMerge=1})", True),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", False),
            ("a:tbl/(a:tr/a:tc{rowSpan=1},a:tr/a:tc{vMerge=false})", False),
            ("a:tbl/(a:tr/a:tc{rowSpan=2},a:tr/a:tc{vMerge=true})", True),
        ]
    )
    def contains_merge_fixture(self, request):
        tbl_cxml, expected_value = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        return tcs[0], tcs[1], expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", (1, 2)),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", (2, 1)),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", (2, 2)),
        ]
    )
    def dimensions_fixture(self, request):
        tbl_cxml, expected_value = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        return tcs[0], tcs[-1], expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", [0, 1], [1]),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], []),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [1, 3]),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                [0, 8],
                [1, 2, 4, 5, 7, 8],
            ),
        ]
    )
    def except_left_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", [0, 1], []),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], [1]),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [2, 3]),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                [0, 8],
                [3, 4, 5, 6, 7, 8],
            ),
        ]
    )
    def except_top_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value

    @pytest.fixture(params=[True, False])
    def in_same_table_fixture(self, request):
        expected_value = request.param
        tbl = element("a:tbl/a:tr/(a:tc,a:tc)")
        other_tbl = element("a:tbl/a:tr/(a:tc,a:tc)")
        tc = tbl.xpath("//a:tc")[0]
        other_tc = tbl.xpath("//a:tc")[1] if expected_value else other_tbl.xpath("//a:tc")[1]
        return tc, other_tc, expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", (0, 1), (0,)),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", (0, 1), (0, 1)),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", (2, 1), (0, 2)),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                (4, 8),
                (4, 7),
            ),
        ]
    )
    def left_col_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc/a:txBody/a:p,a:tc/a:txBody/a:p)", ""),
            ('a:tbl/a:tr/(a:tc/a:txBody/a:p,a:tc/a:txBody/a:p/a:r/a:t"b")', "b"),
            ('a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t"a",a:tc/a:txBody/a:p)', "a"),
            (
                'a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t"a",a:tc/a:txBody/a:p/a:r/a:t' '"b")',
                "a\nb",
            ),
            (
                'a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t"a",a:tc/a:txBody/(a:p,a:p))',
                "a\n\n",
            ),
            (
                'a:tbl/a:tr/(a:tc/a:txBody/(a:p,a:p),a:tc/a:txBody/a:p/a:r/a:t"b")',
                "\n\nb",
            ),
        ]
    )
    def move_fixture(self, request):
        tbl_cxml, expected_text = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        return tcs[0], tcs[1], expected_text

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", (0, 1), (0, 1)),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", (0, 1), (0,)),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", (2, 1), (0, 1)),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                (4, 8),
                (4, 5),
            ),
        ]
    )
    def top_row_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value
