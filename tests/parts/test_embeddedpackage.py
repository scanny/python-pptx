# encoding: utf-8

"""Test suite for `pptx.parts.embeddedpackage` module."""

import pytest

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import OpcPackage, PackURI
from pptx.parts.embeddedpackage import (
    EmbeddedPackagePart,
    EmbeddedDocxPart,
    EmbeddedPptxPart,
    EmbeddedXlsxPart,
)

from ..unitutil.mock import ANY, class_mock, initializer_mock, instance_mock


class DescribeEmbeddedPackagePart(object):
    """Unit-test suite for `pptx.parts.embeddedpackage.EmbeddedPackagePart` objects."""

    @pytest.mark.parametrize(
        "prog_id, EmbeddedPartCls",
        (
            (PROG_ID.DOCX, EmbeddedDocxPart),
            (PROG_ID.PPTX, EmbeddedPptxPart),
            (PROG_ID.XLSX, EmbeddedXlsxPart),
        ),
    )
    def it_provides_a_factory_that_creates_a_package_part_for_MS_Office_files(
        self, request, prog_id, EmbeddedPartCls
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

    def but_it_creates_a_generic_object_part_for_non_MS_Office_files(self, request):
        progId = "Foo.Bar.42"
        object_blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, EmbeddedPackagePart, autospec=True)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_

        ole_object_part = EmbeddedPackagePart.factory(progId, object_blob_, package_)

        package_.next_partname.assert_called_once_with(
            "/ppt/embeddings/oleObject%d.bin"
        )
        _init_.assert_called_once_with(
            ANY, partname_, CT.OFC_OLE_OBJECT, object_blob_, package_
        )
        assert isinstance(ole_object_part, EmbeddedPackagePart)

    def it_provides_a_contructor_classmethod_for_subclasses(self, request):
        blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, EmbeddedXlsxPart, autospec=True)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_

        xlsx_part = EmbeddedXlsxPart.new(blob_, package_)

        package_.next_partname.assert_called_once_with(
            EmbeddedXlsxPart.partname_template
        )
        _init_.assert_called_once_with(
            ANY, partname_, EmbeddedXlsxPart.content_type, blob_, package_
        )
        assert isinstance(xlsx_part, EmbeddedXlsxPart)
