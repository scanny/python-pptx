# encoding: utf-8

"""
Test suite for pptx.opc.package module
"""

from __future__ import absolute_import

import pytest

from collections import namedtuple
from lxml import etree
from mock import call, Mock
from StringIO import StringIO
from zipfile import ZipFile, is_zipfile

import pptx.presentation

from pptx.opc.package import (
    _ContentTypesItem, OldPart, Package, PartTypeSpec, Unmarshaller
)
from pptx.opc.phys_pkg import PhysPkgReader
from pptx.spec import PTS_CARDINALITY_TUPLE, PTS_HASRELS_ALWAYS

from ..unitutil import absjoin, method_mock, test_file_dir


test_pptx_path = absjoin(test_file_dir, 'test.pptx')
dir_pkg_path = absjoin(test_file_dir, 'expanded_pptx')
zip_pkg_path = test_pptx_path


@pytest.fixture
def tmp_pptx_path(tmpdir):
    return str(tmpdir.join('test_python-pptx.pptx'))


class Describe_ContentTypesItem__getitem__(object):
    """Test dictionary-style access of content type using partname as key"""
    def test_it_finds_default_case_insensitive(self, cti):
        """_ContentTypesItem[partname] finds default case insensitive"""
        # setup ------------------------
        partname = '/ppt/media/image1.JPG'
        content_type = 'image/jpeg'
        cti._defaults = {'jpg': content_type}
        # exercise ---------------------
        val = cti[partname]
        # verify -----------------------
        assert val == content_type

    def test_it_finds_override_case_insensitive(self, cti):
        """_ContentTypesItem[partname] finds override case insensitive"""
        # setup ------------------------
        partname = '/foo/bar.xml'
        case_mangled_partname = '/FoO/bAr.XML'
        content_type = 'application/vnd.content_type'
        cti._overrides = {
            partname: content_type
        }
        # exercise ---------------------
        val = cti[case_mangled_partname]
        # verify -----------------------
        assert val == content_type

    def test_getitem_raises_on_bad_partname(self, cti):
        """_ContentTypesItem[partname] raises on bad partname"""
        # verify -----------------------
        with pytest.raises(LookupError):
            cti['!blat/rhumba.1x&']

    # fixtures ---------------------------------------------

    @pytest.fixture
    def cti(self):
        return _ContentTypesItem()


class Describe_ContentTypesItem_compose(object):

    def test_compose_returns_self(self, cti):
        """_ContentTypesItem.compose() returns self-reference"""
        # setup -----------------------
        pkg = Package().open(zip_pkg_path)
        # exercise --------------------
        retval = cti.compose(pkg.parts)
        # verify ----------------------
        expected = cti
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    def test_compose_correct_count(self, cti):
        """_ContentTypesItem.compose() produces expected element count"""
        # setup -----------------------
        pkg = Package().open(zip_pkg_path)
        # exercise --------------------
        cti.compose(pkg.parts)
        # verify ----------------------
        assert len(cti) == 24

    def test_compose_raises_on_bad_partname_ext(self, cti):
        """_ContentTypesItem.compose() raises on bad partname extension"""
        # setup -----------------------
        MockPart = namedtuple('MockPart', 'partname')
        parts = [MockPart('/ppt/!blat/rhumba.1x&')]
        # verify ----------------------
        with pytest.raises(LookupError):
            cti.compose(parts)

    def test_element_correct_length(self, cti):
        """_ContentTypesItem.element() has expected element count"""
        # setup -----------------------
        pkg = Package().open(zip_pkg_path)
        # exercise --------------------
        cti.compose(pkg.parts)
        # verify ----------------------
        assert len(cti.element) == 24

    def test_load_spotcheck(self):
        """_ContentTypesItem can load itself from a filesystem instance"""
        # setup ------------------------
        dir_fs = PhysPkgReader(dir_pkg_path)
        zip_fs = PhysPkgReader(zip_pkg_path)
        for fs in (dir_fs, zip_fs):
            # exercise ---------------------
            cti = _ContentTypesItem().load(fs)
            # test -------------------------
            expected = 'application/vnd.openxmlformats-officedocument'\
                       '.presentationml.slideLayout+xml'
            actual = cti['/ppt/slideLayouts/slideLayout1.xml']
            msg = "expected content type '%s', got '%s'" % (expected, actual)
            assert actual == expected, msg

    # fixtures ---------------------------------------------

    @pytest.fixture
    def cti(self):
        return _ContentTypesItem()


class DescribePackage(object):

    def test_marshal_returns_self(self, pkg):
        """Package.marshal() returns self-reference"""
        # setup -----------------------
        model_pkg = pptx.presentation.Package.open(test_pptx_path)
        # exercise --------------------
        retval = pkg.marshal(model_pkg)
        # verify ----------------------
        expected = pkg
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    def test_open_returns_self(self, pkg):
        """Package.open() returns self-reference"""
        for file in (dir_pkg_path, zip_pkg_path, open(zip_pkg_path, 'rb')):
            # exercise ----------------
            retval = pkg.open(file)
            # verify ------------------
            expected = pkg
            actual = retval
            msg = "expected '%s', got '%s'" % (expected, actual)
            assert actual == expected, msg

    def test_open_part_count(self, pkg):
        """Package.open() produces expected part count"""
        # exercise --------------------
        pkg.open(zip_pkg_path)
        # verify ----------------------
        assert len(pkg.parts) == 22

    def test_open_populates_target_part(self, pkg):
        """Part.open() populates Relationship.target"""
        # exercise --------------------
        pkg.open(zip_pkg_path)
        # verify ----------------------
        for rel in pkg.relationships:
            assert isinstance(rel.target, OldPart)
        for part in pkg.parts:
            for rel in part.relationships:
                assert isinstance(rel.target, OldPart)

    def test_parts_property_empty_on_construction(self, pkg):
        assert len(pkg.parts) == 0

    def test_relationships_property_empty_on_construction(self, pkg):
        assert len(pkg.relationships) == 0

    def test_relationships_correct_length_after_open(self, pkg):
        pkg.open(zip_pkg_path)
        assert len(pkg.relationships) == 4

    def test_relationships_discarded_before_open(self, pkg):
        pkg.open(zip_pkg_path)
        pkg.open(dir_pkg_path)
        assert len(pkg.relationships) == 4

    def test_save_accepts_stream(self, tmp_pptx_path):
        pkg = Package().open(dir_pkg_path)
        stream = StringIO()
        # exercise --------------------
        pkg.save(stream)
        # verify ----------------------
        # can't use is_zipfile() directly on stream in Python 2.6
        stream.seek(0)
        with open(tmp_pptx_path, 'wb') as f:
            f.write(stream.read())
        msg = "Package.save(stream) did not create zipfile"
        assert is_zipfile(tmp_pptx_path), msg

    def test_save_writes_pptx_zipfile(self, pkg, tmp_pptx_path):
        pkg.open(dir_pkg_path)
        save_path = tmp_pptx_path
        # exercise --------------------
        pkg.save(save_path)
        # verify ----------------------
        msg = "no zipfile at %s" % (save_path)
        assert is_zipfile(save_path), msg

    def test_save_member_count(self, pkg, tmp_pptx_path):
        """Package.save() produces expected zip member count"""
        # setup -----------------------
        pkg.open(dir_pkg_path)
        save_path = tmp_pptx_path
        # exercise --------------------
        pkg.save(save_path)
        # verify ----------------------
        zip = ZipFile(save_path)
        names = zip.namelist()
        zip.close()
        partnames = [name for name in names if not name.endswith('/')]
        assert len(partnames) == 38

    # fixtures ---------------------------------------------

    @pytest.fixture
    def pkg(self):
        return Package()


class DescribePart(object):

    def test_blob_none_on_construction(self, part):
        """Part.blob is None on construction"""
        expected = None
        actual = part.blob
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    def test_content_type_correct(self, part, content_type):
        """Part.content_type returns correct value."""
        # setup -----------------------
        part.typespec = Mock()
        part.typespec.content_type = content_type
        # exercise --------------------
        retval = part.content_type
        # verify ----------------------
        expected = content_type
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    def test_partname_property_none_on_construction(self, part):
        assert part.partname is None

    def test_relationships_property_empty_on_construction(self, part):
        assert len(part.relationships) == 0

    def test_typespec_none_on_construction(self, part):
        expected = None
        actual = part.typespec
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    # fixtures ---------------------------------------------

    @pytest.fixture
    def content_type(self):
        return (
            'application/vnd.openxmlformats-officedocument.presentationml.sl'
            'ideMaster+xml'
        )

    @pytest.fixture
    def part(self):
        return OldPart()


class DescribePartTypeSpec(object):
    """Test PartTypeSpec"""
    def test_construction_returns_correct_parttypespec(self):
        """PartTypeSpec(content_type) returns correct PartTypeSpec"""
        # setup -----------------------
        content_type = 'application/vnd.openxmlformats-officedocument.'\
                       'presentationml.notesMaster+xml'
        # exercise --------------------
        pts = PartTypeSpec(content_type)
        # verify ----------------------
        actual = (pts.basename == 'notesMaster' and
                  pts.ext == '.xml' and
                  pts.cardinality == PTS_CARDINALITY_TUPLE and
                  not pts.required and
                  pts.baseURI == '/ppt/notesMasters' and
                  pts.has_rels == PTS_HASRELS_ALWAYS and
                  pts.reltype == ('http://schemas.openxmlformats.org/office'
                                  'Document/2006/relationships/notesMaster')
                  )
        msg = "PartTypeSpec('%s') returns unexpected values" % (content_type)
        assert actual, msg

    def test_construction_raises_on_bad_contenttype(self):
        """PartTypeSpec(ct) raises exception on unrecognized content type"""
        # setup -----------------------
        content_type = 'application/vnd.baloneyMaster+xml'
        # verify ----------------------
        with pytest.raises(KeyError):
            PartTypeSpec(content_type)

    def test_format_correct(self):
        """PartTypeSpec.format returns correct value"""
        # setup -----------------------
        ct_slideMaster = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.slideLayout+xml'
        ct_jpeg = 'image/jpeg'
        ct_printerSettings = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.printerSettings'
        ct_slide = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.slide+xml'
        test_cases = ((ct_slideMaster, 'xml'), (ct_jpeg, 'binary'),
                      (ct_printerSettings, 'binary'), (ct_slide, 'xml'))
        # exercise --------------------
        for ct, exp_format in test_cases:
            pts = PartTypeSpec(ct)
            # verify ----------------------
            expected = exp_format
            actual = pts.format
            msg = "expected '%s', got '%s'" % (expected, actual)
            assert actual == expected, msg


class DescribeRelationship(object):
    """Test Relationship"""
    def setUp(self):
        """Create a new relationship from a string"""
        self.baseURI = '/'
        self.rId = 'rId1'
        self.reltype = 'http://schemas.openxmlformats.org/officeDocument/'\
                       '2006/relationships/officeDocument'
        self.target = 'ppt/presentation.xml'
        tmpl = '<Relationship Id="%s" Type="%s" Target="%s"/>'
        self.rel_xml = tmpl % (self.rId, self.reltype, self.target)
        self.rel_elm = etree.fromstring(self.rel_xml)

    # def test_construction_correct_attr_values(self):
    #     """Relationship attributes loaded from ElementTree.Element"""
    #     # exercise --------------------
    #     rel = self.rel._load(self.baseURI, self.rel_elm)
    #     # verify ----------------------
    #     expected = (self.rId, self.reltype, self.target)
    #     actual = (rel.rId, rel.reltype)
    #     msg = "expected '%s', got '%s'" % (expected, actual)
    #     assert actual == expected, msg


class DescribeUnmarshaller(object):

    def it_can_unmarshal_from_a_pkg_reader(
            self, _unmarshal_parts, _unmarshal_relationships):
        # mockery ----------------------
        pkg = Mock(name='pkg')
        pkg_reader = Mock(name='pkg_reader')
        part_factory = Mock(name='part_factory')
        parts = {1: Mock(name='part_1'), 2: Mock(name='part_2')}
        _unmarshal_parts.return_value = parts
        # exercise ---------------------
        Unmarshaller.unmarshal(pkg_reader, pkg, part_factory)
        # verify -----------------------
        _unmarshal_parts.assert_called_once_with(pkg_reader, part_factory)
        _unmarshal_relationships.assert_called_once_with(pkg_reader, pkg,
                                                         parts)
        for part in parts.values():
            part.after_unmarshal.assert_called_once_with()

    def it_can_unmarshal_parts(self):
        # test data --------------------
        part_properties = (
            ('/part/name1.xml', 'app/vnd.contentType_A', '<Part_1/>'),
            ('/part/name2.xml', 'app/vnd.contentType_B', '<Part_2/>'),
            ('/part/name3.xml', 'app/vnd.contentType_C', '<Part_3/>'),
        )
        # mockery ----------------------
        pkg_reader = Mock(name='pkg_reader')
        pkg_reader.iter_sparts.return_value = part_properties
        part_factory = Mock(name='part_factory')
        parts = [Mock(name='part1'), Mock(name='part2'), Mock(name='part3')]
        part_factory.side_effect = parts
        # exercise ---------------------
        retval = Unmarshaller._unmarshal_parts(pkg_reader, part_factory)
        # verify -----------------------
        expected_calls = [call(*p) for p in part_properties]
        expected_parts = dict((p[0], parts[idx]) for (idx, p) in
                              enumerate(part_properties))
        assert part_factory.call_args_list == expected_calls
        assert retval == expected_parts

    def it_can_unmarshal_relationships(self):
        # test data --------------------
        reltype = 'http://reltype'
        # mockery ----------------------
        pkg_reader = Mock(name='pkg_reader')
        pkg_reader.iter_srels.return_value = (
            ('/',         Mock(name='srel1', rId='rId1', reltype=reltype,
             target_partname='partname1', is_external=False)),
            ('/',         Mock(name='srel2', rId='rId2', reltype=reltype,
             target_ref='target_ref_1',   is_external=True)),
            ('partname1', Mock(name='srel3', rId='rId3', reltype=reltype,
             target_partname='partname2', is_external=False)),
            ('partname2', Mock(name='srel4', rId='rId4', reltype=reltype,
             target_ref='target_ref_2',   is_external=True)),
        )
        pkg = Mock(name='pkg')
        parts = {}
        for num in range(1, 3):
            name = 'part%d' % num
            part = Mock(name=name)
            parts['partname%d' % num] = part
            pkg.attach_mock(part, name)
        # exercise ---------------------
        Unmarshaller._unmarshal_relationships(pkg_reader, pkg, parts)
        # verify -----------------------
        expected_pkg_calls = [
            call._add_relationship(
                reltype, parts['partname1'], 'rId1', False),
            call._add_relationship(
                reltype, 'target_ref_1', 'rId2', True),
            call.part1._add_relationship(
                reltype, parts['partname2'], 'rId3', False),
            call.part2._add_relationship(
                reltype, 'target_ref_2', 'rId4', True),
        ]
        assert pkg.mock_calls == expected_pkg_calls

    # fixtures ---------------------------------------------

    @pytest.fixture
    def _unmarshal_parts(self, request):
        return method_mock(request, Unmarshaller, '_unmarshal_parts')

    @pytest.fixture
    def _unmarshal_relationships(self, request):
        return method_mock(request, Unmarshaller, '_unmarshal_relationships')
