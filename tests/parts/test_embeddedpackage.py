# encoding: utf-8

"""
Test suite for pptx.parts.embeddedpackage module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import OpcPackage, PackURI
from pptx.parts.embeddedpackage import EmbeddedXlsxPart

from ..unitutil.mock import initializer_mock, instance_mock


class DescribeEmbeddedXlsxPart(object):

    def it_can_construct_from_an_xlsx_blob(self, new_fixture):
        xlsx_blob_, package_, init_, partname_, xlsx_part_ = new_fixture

        xlsx_part = EmbeddedXlsxPart.new(xlsx_blob_, package_)

        package_.next_partname.assert_called_once_with(
            EmbeddedXlsxPart.partname_template
        )
        init_.assert_called_once_with(
            partname_, CT.SML_SHEET, xlsx_blob_, package_
        )
        assert isinstance(xlsx_part, EmbeddedXlsxPart)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self, xlsx_blob_, package_, init_, partname_, xlsx_part_):
        return (
            xlsx_blob_, package_, init_, partname_, xlsx_part_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def init_(self, request):
        return initializer_mock(request, EmbeddedXlsxPart)

    @pytest.fixture
    def package_(self, request, partname_):
        package_ = instance_mock(request, OpcPackage)
        package_.next_partname.return_value = partname_
        return package_

    @pytest.fixture
    def partname_(self, request):
        return instance_mock(request, PackURI)

    @pytest.fixture
    def xlsx_blob_(self, request):
        return instance_mock(request, bytes)

    @pytest.fixture
    def xlsx_part_(self, request):
        return instance_mock(request, EmbeddedXlsxPart)
