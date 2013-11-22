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

from lxml import etree
from mock import Mock
from zipfile import BadZipfile, ZIP_DEFLATED, ZipFile

from pptx.exceptions import NotXMLError, PackageNotFoundError
from pptx.opc.packuri import PACKAGE_URI, PackURI
from pptx.opc.phys_pkg import (
    DirectoryFileSystem, _DirPkgReader, FileSystem, PhysPkgReader,
    PhysPkgWriter, ZipFileSystem, _ZipPkgReader, _ZipPkgWriter
)

from ..unitutil import absjoin, class_mock, loose_mock, test_file_dir


test_pptx_path = absjoin(test_file_dir, 'test.pptx')
dir_pkg_path = absjoin(test_file_dir, 'expanded_pptx')
zip_pkg_path = test_pptx_path


class DescribeBaseFileSystem(object):

    def test___contains__(self):
        """'in' operator returns True if URI is in filesystem"""
        expected_URIs = (
            '/[Content_Types].xml',
            '/docProps/app.xml',
            '/ppt/presentation.xml',
            '/ppt/slideMasters/slideMaster1.xml',
            '/ppt/slideLayouts/_rels/slideLayout1.xml.rels')
        fs = FileSystem(zip_pkg_path)
        for uri in expected_URIs:
            assert uri in fs

    def test_getblob_correct_length(self):
        """BaseFileSystem.getblob() returns object of expected length"""
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = FileSystem(zip_pkg_path)
        # exercise --------------------
        blob = fs.getblob(partname)
        # verify ----------------------
        assert len(blob) == 8147

    def test_getblob_raises_on_bad_itemuri(self):
        """BaseFileSystem.getblob(itemURI) raises on bad itemURI"""
        # setup -----------------------
        bad_itemURI = '/spam/eggs/egg1.xml'
        fs = FileSystem(zip_pkg_path)
        # verify ----------------------
        with pytest.raises(LookupError):
            fs.getblob(bad_itemURI)

    def test_getelement_return_count(self):
        """ElementTree element for specified package item is returned"""
        dir_fs = FileSystem(dir_pkg_path)
        zip_fs = FileSystem(zip_pkg_path)
        for fs in (dir_fs, zip_fs):
            elm = fs.getelement('/[Content_Types].xml')
            assert len(elm) == 24

    def test_getelement_raises_on_bad_itemuri(self):
        """BaseFileSystem.getelement(itemURI) raises on bad itemURI"""
        # setup -----------------------
        bad_itemURI = '/spam/eggs/egg1.xml'
        fs = FileSystem(zip_pkg_path)
        # verify ----------------------
        with pytest.raises(LookupError):
            fs.getelement(bad_itemURI)

    def test_getelement_raises_on_binary(self):
        """Calling getelement() for binary item raises exception"""
        # call getelement for thumbnail
        fs = FileSystem(zip_pkg_path)
        with pytest.raises(NotXMLError):
            fs.getelement('/docProps/thumbnail.jpeg')


class DescribeDirectoryFileSystem(object):

    def test_constructor_raises_on_non_dir_path(self):
        """DirectoryFileSystem(path) raises on non-dir *path*"""
        with pytest.raises(ValueError):
            DirectoryFileSystem(zip_pkg_path)

    def test_getstream_correct_length(self):
        """BytesIO instance for specified package item is returned"""
        fs = DirectoryFileSystem(dir_pkg_path)
        stream = fs.getstream('/[Content_Types].xml')
        elm = etree.parse(stream).getroot()
        assert len(elm) == 24

    def test_getstream_raises_on_bad_URI(self):
        """DirectoryFileSystem.getstream() raises on bad URI"""
        fs = DirectoryFileSystem(dir_pkg_path)
        with pytest.raises(LookupError):
            fs.getstream('!blat/rhumba.xml')

    def test_itemURIs_count(self):
        """DirectoryFileSystem.itemURIs has expected count"""
        # verify ----------------------
        fs = DirectoryFileSystem(dir_pkg_path)
        assert len(fs.itemURIs) == 38

    def test_itemURIs_plausible(self):
        """All URIs in DirectoryFileSystem.itemURIs are plausible"""
        # setup -----------------------
        fs = DirectoryFileSystem(dir_pkg_path)
        # verify ----------------------
        for itemURI in fs.itemURIs:
            # plausible segment count
            expected_min = 1
            expected_max = 4
            # leading slash produces empty string in split list
            segment_count = len(itemURI.split('/'))-1
            msg = ("item URI has implausible number of segments:\n"
                   "itemURI ==> '%s'" % (itemURI))
            assert segment_count >= expected_min, msg
            assert segment_count <= expected_max, msg
            # check for forward slash
            msg = ("item URI '%s' does not start with forward slash ('/')"
                   % (itemURI))
            assert itemURI.startswith('/'), msg


class DescribeDirPkgReader(object):

    def it_is_used_by_PhysPkgReader_when_pkg_is_a_dir(self):
        phys_reader = PhysPkgReader(dir_pkg_path)
        assert isinstance(phys_reader, _DirPkgReader)

    def it_doesnt_mind_being_closed_even_though_it_doesnt_need_it(
            self, dir_reader):
        dir_reader.close()

    def it_can_retrieve_the_blob_for_a_pack_uri(self, dir_reader):
        pack_uri = PackURI('/ppt/presentation.xml')
        blob = dir_reader.blob_for(pack_uri)
        sha1 = hashlib.sha1(blob).hexdigest()
        assert sha1 == '51b78f4dabc0af2419d4e044ab73028c4bef53aa'

    def it_can_get_the_content_types_xml(self, dir_reader):
        sha1 = hashlib.sha1(dir_reader.content_types_xml).hexdigest()
        assert sha1 == 'a68cf138be3c4eb81e47e2550166f9949423c7df'

    def it_can_retrieve_the_rels_xml_for_a_source_uri(self, dir_reader):
        rels_xml = dir_reader.rels_xml_for(PACKAGE_URI)
        sha1 = hashlib.sha1(rels_xml).hexdigest()
        assert sha1 == '64ffe86bb2bbaad53c3c1976042b907f8e10c5a3'

    def it_returns_none_when_part_has_no_rels_xml(self, dir_reader):
        partname = PackURI('/ppt/viewProps.xml')
        rels_xml = dir_reader.rels_xml_for(partname)
        assert rels_xml is None

    # fixtures ---------------------------------------------

    @pytest.fixture
    def pkg_file_(self, request):
        return loose_mock(request)

    @pytest.fixture(scope='class')
    def dir_reader(self):
        return _DirPkgReader(dir_pkg_path)


class DescribeFileSystem(object):
    """Test FileSystem"""
    def test_constructor_returns_dirfs_for_dirpath(self):
        """FileSystem(dirpath) returns instance of DirectoryFileSystem"""
        fs = FileSystem(dir_pkg_path)
        assert isinstance(fs, DirectoryFileSystem)

    def test_constructor_returns_zipfs_for_zipfile_path(self):
        """FileSystem(zipfile_path) returns instance of ZipFileSystem"""
        fs = FileSystem(zip_pkg_path)
        assert isinstance(fs, ZipFileSystem)

    def test_constructor_returns_zipfs_for_zip_stream(self):
        """FileSystem(zipfile_stream) returns instance of ZipFileSystem"""
        with open(zip_pkg_path, 'rb') as stream:
            fs = FileSystem(stream)
        assert isinstance(fs, ZipFileSystem)

    def test_constructor_raises_on_bad_path(self):
        """FileSystem(path) constructor raises on bad path"""
        # setup -----------------------
        bad_path = 'blat/rhumba.1x&'
        # verify ----------------------
        with pytest.raises(PackageNotFoundError):
            FileSystem(bad_path)

    def test_constructor_raises_on_non_zip_stream(self):
        """FileSystem(path) constructor raises on non-zip stream"""
        # setup -----------------------
        non_zip_stream = BytesIO('not a zip file')
        # verify ----------------------
        with pytest.raises(BadZipfile):
            FileSystem(non_zip_stream)


class DescribePhysPkgReader(object):

    def it_raises_when_pkg_path_is_not_a_package(self):
        with pytest.raises(PackageNotFoundError):
            PhysPkgReader('foobar')


class DescribePhysPkgWriter(object):

    def it_constructs_a_pkg_writer_instance(self, _ZipPkgWriter_, pkg_file_):
        phys_pkg_writer = PhysPkgWriter(pkg_file_)
        _ZipPkgWriter_.assert_called_once_with(pkg_file_)
        assert phys_pkg_writer == _ZipPkgWriter_.return_value

    # fixtures ---------------------------------------------

    @pytest.fixture
    def pkg_file_(self, request):
        return loose_mock(request)

    @pytest.fixture
    def _ZipPkgWriter_(self, request):
        return class_mock(request, 'pptx.opc.phys_pkg._ZipPkgWriter')


class DescribeZipFileSystem(object):
    """
    Test ZipFileSystem (writing aspect)
    """
    def test_constructor_accepts_stream(self):
        """ZipFileSystem() constructor accepts zip archive as stream"""
        with open(zip_pkg_path, 'rb') as stream:
            fs = ZipFileSystem(stream)
        assert isinstance(fs, ZipFileSystem)

    def test_getstream_correct_length(self):
        """
        [Content_Types].xml retrieved as stream has correct element count
        """
        fs = ZipFileSystem(zip_pkg_path)
        stream = fs.getstream('/[Content_Types].xml')
        content_types_elm = etree.parse(stream).getroot()
        assert len(content_types_elm) == 24

    def test_getstream_raises_on_bad_URI(self):
        """ZipFileSystem.getstream() raises on bad URI"""
        # setup -----------------------
        fs = FileSystem(zip_pkg_path)
        # verify ----------------------
        with pytest.raises(LookupError):
            fs.getstream('!blat/rhumba.xml')

    def test_itemURIs_count(self):
        """ZipFileSystem.itemURIs has expected count"""
        # verify ----------------------
        fs = ZipFileSystem(zip_pkg_path)
        assert len(fs.itemURIs) == 38

    def test_itemURIs_plausible(self):
        """All URIs in ZipFileSystem.itemURIs are plausible"""
        # setup -----------------------
        fs = ZipFileSystem(zip_pkg_path)
        # verify ----------------------
        for itemURI in fs.itemURIs:
            # plausible segment count
            expected_min = 1
            expected_max = 4
            # leading slash produces empty string in split list
            segment_count = len(itemURI.split('/'))-1
            msg = ("item URI has implausible number of segments:\n"
                   "itemURI ==> '%s'" % (itemURI))
            assert segment_count >= expected_min, msg
            assert segment_count <= expected_max, msg
            # check for forward slash
            msg = ("item URI '%s' does not start with forward slash ('/')"
                   % (itemURI))
            assert itemURI.startswith('/'), msg

    # fixtures ---------------------------------------------

    @pytest.fixture
    def xml_in(self):
        return (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<p:presentationPr xmlns:a="http://main" xmlns:r="http://relatio'
            'nships" xmlns:p="http://presentationml">\n'
            '  <p:extLst>\n'
            '    <p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}">\n'
            '      <r:discardImageEditData val="0"/>\n'
            '    </p:ext>\n'
            '  </p:extLst>\n'
            '</p:presentationPr>\n'
        )

    @pytest.fixture
    def xml_out(self):
        return (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<p:presentationPr xmlns:a="http://main" xmlns:r="http://relatio'
            'nships" xmlns:p="http://presentationml"><p:extLst><p:ext uri="{'
            'E76CE94A-603C-4142-B9EB-6D1370010A27}"><r:discardImageEditData '
            'val="0"/></p:ext></p:extLst></p:presentationPr>'
        )


class DescribeZipPkgReader(object):

    def it_is_used_by_PhysPkgReader_when_pkg_is_a_zip(self):
        phys_reader = PhysPkgReader(zip_pkg_path)
        assert isinstance(phys_reader, _ZipPkgReader)

    def it_is_used_by_PhysPkgReader_when_pkg_is_a_stream(self):
        with open(zip_pkg_path, 'rb') as stream:
            phys_reader = PhysPkgReader(stream)
        assert isinstance(phys_reader, _ZipPkgReader)

    def it_opens_pkg_file_zip_on_construction(self, ZipFile_, pkg_file_):
        _ZipPkgReader(pkg_file_)
        ZipFile_.assert_called_once_with(pkg_file_, 'r')

    def it_can_be_closed(self, ZipFile_):
        # mockery ----------------------
        zipf = ZipFile_.return_value
        zip_pkg_reader = _ZipPkgReader(None)
        # exercise ---------------------
        zip_pkg_reader.close()
        # verify -----------------------
        zipf.close.assert_called_once_with()

    def it_can_retrieve_the_blob_for_a_pack_uri(self, phys_reader):
        pack_uri = PackURI('/ppt/presentation.xml')
        blob = phys_reader.blob_for(pack_uri)
        sha1 = hashlib.sha1(blob).hexdigest()
        assert sha1 == 'efa7bee0ac72464903a67a6744c1169035d52a54'

    def it_has_the_content_types_xml(self, phys_reader):
        sha1 = hashlib.sha1(phys_reader.content_types_xml).hexdigest()
        assert sha1 == 'ab762ac84414fce18893e18c3f53700c01db56c3'

    def it_can_retrieve_rels_xml_for_source_uri(self, phys_reader):
        rels_xml = phys_reader.rels_xml_for(PACKAGE_URI)
        sha1 = hashlib.sha1(rels_xml).hexdigest()
        assert sha1 == 'e31451d4bbe7d24adbe21454b8e9fdae92f50de5'

    def it_returns_none_when_part_has_no_rels_xml(self, phys_reader):
        partname = PackURI('/ppt/viewProps.xml')
        rels_xml = phys_reader.rels_xml_for(partname)
        assert rels_xml is None

    # fixtures ---------------------------------------------

    @pytest.fixture(scope='class')
    def phys_reader(self, request):
        phys_reader = _ZipPkgReader(zip_pkg_path)
        request.addfinalizer(phys_reader.close)
        return phys_reader

    @pytest.fixture
    def pkg_file_(self, request):
        return loose_mock(request)


class DescribeZipPkgWriter(object):

    @pytest.fixture
    def pkg_file(self, request):
        pkg_file = BytesIO()
        request.addfinalizer(pkg_file.close)
        return pkg_file

    def it_opens_pkg_file_zip_on_construction(self, ZipFile_):
        pkg_file = Mock(name='pkg_file')
        _ZipPkgWriter(pkg_file)
        ZipFile_.assert_called_once_with(
            pkg_file, 'w', compression=ZIP_DEFLATED
        )

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
        pack_uri = PackURI('/part/name.xml')
        blob = '<BlobbityFooBlob/>'.encode('utf-8')
        # exercise ---------------------
        pkg_writer = PhysPkgWriter(pkg_file)
        pkg_writer.write(pack_uri, blob)
        pkg_writer.close()
        # verify -----------------------
        written_blob_sha1 = hashlib.sha1(blob).hexdigest()
        zipf = ZipFile(pkg_file, 'r')
        retrieved_blob = zipf.read(pack_uri.membername)
        zipf.close()
        retrieved_blob_sha1 = hashlib.sha1(retrieved_blob).hexdigest()
        assert retrieved_blob_sha1 == written_blob_sha1


# fixtures -------------------------------------------------

@pytest.fixture
def tmp_pptx_path(tmpdir):
    return str(tmpdir.join('test_python-pptx.pptx'))


@pytest.fixture
def ZipFile_(request):
    return class_mock(request, 'pptx.opc.phys_pkg.ZipFile')
