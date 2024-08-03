"""Unit-test suite for `pptx.parts.embeddedpackage` module."""

from __future__ import annotations

import pytest

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import OpcPackage, PackURI
from pptx.parts.embeddedpackage import (
    EmbeddedDocxPart,
    EmbeddedPackagePart,
    EmbeddedPptxPart,
    EmbeddedXlsxPart,
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
