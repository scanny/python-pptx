# encoding: utf-8

"""Test suite for `pptx.parts.embeddedpackage` module."""

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import OpcPackage, PackURI
from pptx.parts.embeddedpackage import EmbeddedXlsxPart

from ..unitutil.mock import ANY, initializer_mock, instance_mock


class DescribeEmbeddedXlsxPart(object):
    """Unit-test suite for `pptx.parts.embeddedpackage.EmbeddedXlsxPart` objects."""

    def it_can_construct_from_an_xlsx_blob(self, request):
        xlsx_blob_ = b"0123456789"
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, EmbeddedXlsxPart, autospec=True)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_

        xlsx_part = EmbeddedXlsxPart.new(xlsx_blob_, package_)

        package_.next_partname.assert_called_once_with(
            EmbeddedXlsxPart.partname_template
        )
        _init_.assert_called_once_with(
            ANY, partname_, CT.SML_SHEET, xlsx_blob_, package_
        )
        assert isinstance(xlsx_part, EmbeddedXlsxPart)
