# encoding: utf-8

"""
Test suite for pptx.opc.packaging module
"""

from __future__ import absolute_import

try:
    from io import BytesIO  # Python 3
except ImportError:
    from StringIO import StringIO as BytesIO

import hashlib
import pytest

from zipfile import ZIP_DEFLATED, ZipFile

from pptx.exceptions import PackageNotFoundError
from pptx.opc.packuri import PACKAGE_URI, PackURI
from pptx.opc.phys_pkg import (
    _DirPkgReader,
    PhysPkgReader,
    PhysPkgWriter,
    _ZipPkgReader,
    _ZipPkgWriter,
)

from ..unitutil.file import absjoin, test_file_dir
from ..unitutil.mock import class_mock, loose_mock, Mock


test_pptx_path = absjoin(test_file_dir, "test.pptx")
dir_pkg_path = absjoin(test_file_dir, "expanded_pptx")
zip_pkg_path = test_pptx_path


class DescribeDirPkgReader(object):
    def it_is_used_by_PhysPkgReader_when_pkg_is_a_dir(self):
        phys_reader = PhysPkgReader(dir_pkg_path)
        assert isinstance(phys_reader, _DirPkgReader)

    def it_doesnt_mind_being_closed_even_though_it_doesnt_need_it(self, dir_reader):
        dir_reader.close()

    def it_can_retrieve_the_blob_for_a_pack_uri(self, dir_reader):
        pack_uri = PackURI("/ppt/presentation.xml")
        blob = dir_reader.blob_for(pack_uri)
        sha1 = hashlib.sha1(blob).hexdigest()
        assert sha1 == "51b78f4dabc0af2419d4e044ab73028c4bef53aa"

    def it_can_get_the_content_types_xml(self, dir_reader):
        sha1 = hashlib.sha1(dir_reader.content_types_xml).hexdigest()
        assert sha1 == "a68cf138be3c4eb81e47e2550166f9949423c7df"

    def it_can_retrieve_the_rels_xml_for_a_source_uri(self, dir_reader):
        rels_xml = dir_reader.rels_xml_for(PACKAGE_URI)
        sha1 = hashlib.sha1(rels_xml).hexdigest()
        assert sha1 == "64ffe86bb2bbaad53c3c1976042b907f8e10c5a3"

    def it_returns_none_when_part_has_no_rels_xml(self, dir_reader):
        partname = PackURI("/ppt/viewProps.xml")
        rels_xml = dir_reader.rels_xml_for(partname)
        assert rels_xml is None

    # fixtures ---------------------------------------------

    @pytest.fixture
    def pkg_file_(self, request):
        return loose_mock(request)

    @pytest.fixture(scope="class")
    def dir_reader(self):
        return _DirPkgReader(dir_pkg_path)


class DescribePhysPkgReader(object):
    def it_raises_when_pkg_path_is_not_a_package(self):
        with pytest.raises(PackageNotFoundError):
            PhysPkgReader("foobar")


class DescribeZipPkgReader(object):
    def it_is_used_by_PhysPkgReader_when_pkg_is_a_zip(self):
        phys_reader = PhysPkgReader(zip_pkg_path)
        assert isinstance(phys_reader, _ZipPkgReader)

    def it_is_used_by_PhysPkgReader_when_pkg_is_a_stream(self):
        with open(zip_pkg_path, "rb") as stream:
            phys_reader = PhysPkgReader(stream)
        assert isinstance(phys_reader, _ZipPkgReader)

    def it_opens_pkg_file_zip_on_construction(self, ZipFile_, pkg_file_):
        _ZipPkgReader(pkg_file_)
        ZipFile_.assert_called_once_with(pkg_file_, "r")

    def it_can_be_closed(self, ZipFile_):
        # mockery ----------------------
        zipf = ZipFile_.return_value
        zip_pkg_reader = _ZipPkgReader(None)
        # exercise ---------------------
        zip_pkg_reader.close()
        # verify -----------------------
        zipf.close.assert_called_once_with()

    def it_can_retrieve_the_blob_for_a_pack_uri(self, phys_reader):
        pack_uri = PackURI("/ppt/presentation.xml")
        blob = phys_reader.blob_for(pack_uri)
        sha1 = hashlib.sha1(blob).hexdigest()
        assert sha1 == "efa7bee0ac72464903a67a6744c1169035d52a54"

    def it_has_the_content_types_xml(self, phys_reader):
        sha1 = hashlib.sha1(phys_reader.content_types_xml).hexdigest()
        assert sha1 == "ab762ac84414fce18893e18c3f53700c01db56c3"

    def it_can_retrieve_rels_xml_for_source_uri(self, phys_reader):
        rels_xml = phys_reader.rels_xml_for(PACKAGE_URI)
        sha1 = hashlib.sha1(rels_xml).hexdigest()
        assert sha1 == "e31451d4bbe7d24adbe21454b8e9fdae92f50de5"

    def it_returns_none_when_part_has_no_rels_xml(self, phys_reader):
        partname = PackURI("/ppt/viewProps.xml")
        rels_xml = phys_reader.rels_xml_for(partname)
        assert rels_xml is None

    # fixtures ---------------------------------------------

    @pytest.fixture(scope="class")
    def phys_reader(self, request):
        phys_reader = _ZipPkgReader(zip_pkg_path)
        request.addfinalizer(phys_reader.close)
        return phys_reader

    @pytest.fixture
    def pkg_file_(self, request):
        return loose_mock(request)


class DescribeZipPkgWriter(object):
    def it_is_used_by_PhysPkgWriter_unconditionally(self, tmp_pptx_path):
        phys_writer = PhysPkgWriter(tmp_pptx_path)
        assert isinstance(phys_writer, _ZipPkgWriter)

    def it_opens_pkg_file_zip_on_construction(self, ZipFile_):
        pkg_file = Mock(name="pkg_file")
        _ZipPkgWriter(pkg_file)
        ZipFile_.assert_called_once_with(pkg_file, "w", compression=ZIP_DEFLATED)

    def it_can_be_closed(self, ZipFile_):
        # mockery ----------------------
        zipf = ZipFile_.return_value
        zip_pkg_writer = _ZipPkgWriter(None)
        # exercise ---------------------
        zip_pkg_writer.close()
        # verify -----------------------
        zipf.close.assert_called_once_with()

    def it_can_write_a_blob(self, pkg_file):
        # setup ------------------------
        pack_uri = PackURI("/part/name.xml")
        blob = "<BlobbityFooBlob/>".encode("utf-8")
        # exercise ---------------------
        pkg_writer = PhysPkgWriter(pkg_file)
        pkg_writer.write(pack_uri, blob)
        pkg_writer.close()
        # verify -----------------------
        written_blob_sha1 = hashlib.sha1(blob).hexdigest()
        zipf = ZipFile(pkg_file, "r")
        retrieved_blob = zipf.read(pack_uri.membername)
        zipf.close()
        retrieved_blob_sha1 = hashlib.sha1(retrieved_blob).hexdigest()
        assert retrieved_blob_sha1 == written_blob_sha1

    # fixtures ---------------------------------------------

    @pytest.fixture
    def pkg_file(self, request):
        pkg_file = BytesIO()
        request.addfinalizer(pkg_file.close)
        return pkg_file


# fixtures -------------------------------------------------


@pytest.fixture
def tmp_pptx_path(tmpdir):
    return str(tmpdir.join("test_python-pptx.pptx"))


@pytest.fixture
def ZipFile_(request):
    return class_mock(request, "pptx.opc.phys_pkg.ZipFile")
