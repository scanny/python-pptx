"""Unit-test suite for `pptx.enum.shapes`."""

from __future__ import annotations

import pytest

from pptx.enum.shapes import PROG_ID


class DescribeProgId:
    """Unit-test suite for `pptx.enum.shapes.ProgId`."""

    def it_has_members_for_the_OLE_embeddings_known_to_work_on_Windows(self):
        assert PROG_ID.DOCX
        assert PROG_ID.PPTX
        assert PROG_ID.XLSX

    def it_has_generic_members_for_common_embeddable_file_types(self):
        """PROG_ID gained ZIP/PDF/DOC/HTML convenience members for #752."""
        assert PROG_ID.ZIP.progId == "Package"
        assert PROG_ID.PDF.progId == "AcroExch.Document.DC"
        assert PROG_ID.DOC.progId == "Word.Document.8"
        assert PROG_ID.HTML.progId == "htmlfile"

    @pytest.mark.parametrize(
        ("member", "expected_extension"),
        [
            (PROG_ID.DOCX, "docx"),
            (PROG_ID.PPTX, "pptx"),
            (PROG_ID.XLSX, "xlsx"),
            (PROG_ID.ZIP, "zip"),
            (PROG_ID.PDF, "pdf"),
            (PROG_ID.HTML, "html"),
            (PROG_ID.DOC, "doc"),
        ],
    )
    def it_knows_its_part_name_extension(self, member: PROG_ID, expected_extension: str):
        assert member.extension == expected_extension

    @pytest.mark.parametrize(
        ("member", "expected_value"),
        [
            (PROG_ID.DOCX, True),
            (PROG_ID.PPTX, True),
            (PROG_ID.XLSX, True),
            (PROG_ID.ZIP, False),
            (PROG_ID.PDF, False),
            (PROG_ID.HTML, False),
            (PROG_ID.DOC, False),
        ],
    )
    def it_knows_whether_it_is_an_office_package(
        self, member: PROG_ID, expected_value: bool
    ):
        assert member.is_office_package is expected_value

    @pytest.mark.parametrize(
        ("member", "expected_value"),
        [(PROG_ID.DOCX, 609600), (PROG_ID.PPTX, 609600), (PROG_ID.XLSX, 609600)],
    )
    def it_knows_its_height(self, member: PROG_ID, expected_value: int):
        assert member.height == expected_value

    def it_knows_its_icon_filename(self):
        assert PROG_ID.DOCX.icon_filename == "docx-icon.emf"

    def it_knows_its_progId(self):
        assert PROG_ID.PPTX.progId == "PowerPoint.Show.12"

    def it_knows_its_width(self):
        assert PROG_ID.XLSX.width == 965200

    @pytest.mark.parametrize(
        ("value", "expected_value"),
        [
            # -DELETEME---------------------------------------------------------------
            (PROG_ID.DOCX, True),
            (PROG_ID.PPTX, True),
            (PROG_ID.XLSX, True),
            (17, False),
            ("XLSX", False),
        ],
    )
    def it_knows_each_of_its_members_is_an_instance(self, value: object, expected_value: bool):
        assert isinstance(value, PROG_ID) is expected_value
