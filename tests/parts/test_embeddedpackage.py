"""Unit-test suite for `pptx.parts.embeddedpackage` module."""

from __future__ import annotations

import pytest

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import OpcPackage, PackURI
from pptx.parts.chart import ChartPart, ChartWorkbook
from pptx.parts.embeddedpackage import (
    EmbeddedDocxPart,
    EmbeddedPackagePart,
    EmbeddedPptxPart,
    EmbeddedXlsxPart,
    clone_embedded_xlsx,
)

from ..unitutil.mock import ANY, FixtureRequest, class_mock, initializer_mock, instance_mock


class DescribeEmbeddedPackagePart(object):
    """Unit-test suite for `pptx.parts.embeddedpackage.EmbeddedPackagePart` objects."""

    @pytest.mark.parametrize(
        ("prog_id", "EmbeddedPartCls"),
        [
            (PROG_ID.DOCX, EmbeddedDocxPart),
            (PROG_ID.PPTX, EmbeddedPptxPart),
            (PROG_ID.XLSX, EmbeddedXlsxPart),
        ],
    )
    def it_provides_a_factory_that_creates_a_package_part_for_MS_Office_files(
        self, request: FixtureRequest, prog_id: PROG_ID, EmbeddedPartCls: type
    ):
        object_blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        embedded_object_part_ = instance_mock(request, EmbeddedPartCls)
        EmbeddedPartCls_ = class_mock(
            request, "pptx.parts.embeddedpackage.%s" % EmbeddedPartCls.__name__
        )
        EmbeddedPartCls_.new.return_value = embedded_object_part_

        ole_object_part = EmbeddedPackagePart.factory(prog_id, object_blob_, package_)

        EmbeddedPartCls_.new.assert_called_once_with(object_blob_, package_)
        assert ole_object_part is embedded_object_part_

    def but_it_creates_a_generic_object_part_for_non_MS_Office_files(self, request: FixtureRequest):
        progId = "Foo.Bar.42"
        object_blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, EmbeddedPackagePart, autospec=True)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_

        ole_object_part = EmbeddedPackagePart.factory(progId, object_blob_, package_)

        package_.next_partname.assert_called_once_with("/ppt/embeddings/oleObject%d.bin")
        _init_.assert_called_once_with(ANY, partname_, CT.OFC_OLE_OBJECT, package_, object_blob_)
        assert isinstance(ole_object_part, EmbeddedPackagePart)

    @pytest.mark.parametrize(
        ("extension", "expected_tmpl"),
        [
            ("zip", "/ppt/embeddings/oleObject%d.zip"),
            ("pdf", "/ppt/embeddings/oleObject%d.pdf"),
            ("html", "/ppt/embeddings/oleObject%d.html"),
            # -- invalid (non-ASCII-alphanum) falls back to .bin ---
            ("../../etc", "/ppt/embeddings/oleObject%d.bin"),
            ("", "/ppt/embeddings/oleObject%d.bin"),
        ],
    )
    def and_it_uses_the_caller_supplied_extension_for_a_str_progId(
        self, request: FixtureRequest, extension: str, expected_tmpl: str
    ):
        """Arbitrary progId + caller-supplied extension produces matching part-name."""
        progId = "Foo.Bar.42"
        object_blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, EmbeddedPackagePart, autospec=True)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_

        ole_object_part = EmbeddedPackagePart.factory(
            progId, object_blob_, package_, extension
        )

        package_.next_partname.assert_called_once_with(expected_tmpl)
        _init_.assert_called_once_with(ANY, partname_, CT.OFC_OLE_OBJECT, package_, object_blob_)
        assert isinstance(ole_object_part, EmbeddedPackagePart)

    def it_uses_the_enum_extension_for_a_generic_PROG_ID_member(self, request: FixtureRequest):
        """Non-Office PROG_ID members (ZIP, PDF, DOC, HTML) carry their own extension."""
        object_blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, EmbeddedPackagePart, autospec=True)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_

        ole_object_part = EmbeddedPackagePart.factory(PROG_ID.ZIP, object_blob_, package_)

        package_.next_partname.assert_called_once_with("/ppt/embeddings/oleObject%d.zip")
        _init_.assert_called_once_with(ANY, partname_, CT.OFC_OLE_OBJECT, package_, object_blob_)
        assert isinstance(ole_object_part, EmbeddedPackagePart)

    def it_provides_a_contructor_classmethod_for_subclasses(self, request: FixtureRequest):
        blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, EmbeddedXlsxPart, autospec=True)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_

        xlsx_part = EmbeddedXlsxPart.new(blob_, package_)

        package_.next_partname.assert_called_once_with(EmbeddedXlsxPart.partname_template)
        _init_.assert_called_once_with(
            xlsx_part, partname_, EmbeddedXlsxPart.content_type, package_, blob_
        )
        assert isinstance(xlsx_part, EmbeddedXlsxPart)


class Describe_clone_embedded_xlsx(object):
    """Unit-test suite for `pptx.parts.embeddedpackage.clone_embedded_xlsx`."""

    def it_duplicates_the_workbook_part_onto_the_target_chart(self, request: FixtureRequest):
        # --- source chart has an embedded workbook with known bytes ---
        source_xlsx_part_ = instance_mock(request, EmbeddedXlsxPart, blob=b"source-blob")
        source_workbook_ = instance_mock(request, ChartWorkbook, xlsx_part=source_xlsx_part_)
        source_chart_part_ = instance_mock(request, ChartPart, chart_workbook=source_workbook_)
        # --- target chart owns a distinct package ---
        target_package_ = instance_mock(request, OpcPackage)
        target_workbook_ = instance_mock(request, ChartWorkbook)
        target_chart_part_ = instance_mock(
            request, ChartPart, chart_workbook=target_workbook_, package=target_package_
        )
        new_xlsx_part_ = instance_mock(request, EmbeddedXlsxPart)
        EmbeddedXlsxPart_ = class_mock(
            request, "pptx.parts.embeddedpackage.EmbeddedXlsxPart"
        )
        EmbeddedXlsxPart_.new.return_value = new_xlsx_part_

        result = clone_embedded_xlsx(source_chart_part_, target_chart_part_)

        EmbeddedXlsxPart_.new.assert_called_once_with(b"source-blob", target_package_)
        # --- clone attaches new part to target chart's workbook ---
        assert target_workbook_.xlsx_part == new_xlsx_part_
        assert result is new_xlsx_part_

    def but_it_returns_None_when_source_has_no_embedded_workbook(
        self, request: FixtureRequest
    ):
        source_workbook_ = instance_mock(request, ChartWorkbook, xlsx_part=None)
        source_chart_part_ = instance_mock(request, ChartPart, chart_workbook=source_workbook_)
        target_chart_part_ = instance_mock(request, ChartPart)
        EmbeddedXlsxPart_ = class_mock(
            request, "pptx.parts.embeddedpackage.EmbeddedXlsxPart"
        )

        result = clone_embedded_xlsx(source_chart_part_, target_chart_part_)

        assert result is None
        EmbeddedXlsxPart_.new.assert_not_called()
