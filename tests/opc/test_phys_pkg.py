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

from pptx.exceptions import (
    DuplicateKeyError, NotXMLError, PackageNotFoundError
)
from pptx.opc.phys_pkg import (
    DirectoryFileSystem, FileSystem, PhysPkgWriter, ZipFileSystem,
    ZipPkgWriter
)
from pptx.util import Partname as PackURI

from ..unitutil import absjoin, class_mock, test_file_dir


test_pptx_path = absjoin(test_file_dir, 'test.pptx')
dir_pkg_path = absjoin(test_file_dir, 'expanded_pptx')
zip_pkg_path = test_pptx_path


@pytest.fixture
def tmp_pptx_path(tmpdir):
    return str(tmpdir.join('test_python-pptx.pptx'))


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


class DescribePhysPkgWriter(object):

    @pytest.fixture
    def ZipPkgWriter_(self, request):
        return class_mock(request, 'pptx.opc.phys_pkg.ZipPkgWriter')

    def it_constructs_a_pkg_writer_instance(self, ZipPkgWriter_):
        # mockery ----------------------
        pkg_file = Mock(name='pkg_file')
        # exercise ---------------------
        phys_pkg_writer = PhysPkgWriter(pkg_file)
        # verify -----------------------
        ZipPkgWriter_.assert_called_once_with(pkg_file)
        assert phys_pkg_writer == ZipPkgWriter_.return_value


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

    def test_write_blob_round_trips(self, tmp_pptx_path):
        """ZipFileSystem.write_blob() round-trips intact"""
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = FileSystem(zip_pkg_path)
        in_blob = fs.getblob(partname)
        test_fs = ZipFileSystem(tmp_pptx_path, 'w')
        # exercise --------------------
        test_fs.write_blob(in_blob, partname)
        # verify ----------------------
        out_blob = test_fs.getblob(partname)
        expected = in_blob
        actual = out_blob
        msg = ("retrived blob (len %d) differs from original (len %d)" %
               (len(actual), len(expected)))
        assert actual == expected, msg

    def test_write_blob_raises_on_dup_itemuri(self, tmp_pptx_path):
        """ZipFileSystem.write_blob() raises on duplicate itemURI"""
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = FileSystem(zip_pkg_path)
        blob = fs.getblob(partname)
        test_fs = ZipFileSystem(tmp_pptx_path, 'w')
        test_fs.write_blob(blob, partname)
        # verify ----------------------
        with pytest.raises(DuplicateKeyError):
            test_fs.write_blob(blob, partname)

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


class DescribeZipPkgWriter(object):

    @pytest.fixture
    def pkg_file(self, request):
        pkg_file = BytesIO()
        request.addfinalizer(pkg_file.close)
        return pkg_file

    def it_opens_pkg_file_zip_on_construction(self, ZipFile_):
        pkg_file = Mock(name='pkg_file')
        ZipPkgWriter(pkg_file)
        ZipFile_.assert_called_once_with(pkg_file, 'w',
                                         compression=ZIP_DEFLATED)

    def it_can_be_closed(self, ZipFile_):
        # mockery ----------------------
        zipf = ZipFile_.return_value
        zip_pkg_writer = ZipPkgWriter(None)
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
def ZipFile_(request):
    return class_mock(request, 'pptx.opc.phys_pkg.ZipFile')
