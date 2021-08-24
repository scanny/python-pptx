# encoding: utf-8

"""Unit-test suite for `pptx.opc.serialized` module."""

try:
    from io import BytesIO  # Python 3
except ImportError:
    from StringIO import StringIO as BytesIO

import hashlib
import pytest

from zipfile import ZIP_DEFLATED, ZipFile

from pptx.exceptions import PackageNotFoundError
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.oxml import CT_Relationship
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.opc.serialized import (
    PackageReader,
    PackageWriter,
    _ContentTypeMap,
    _ContentTypesItem,
    _DirPkgReader,
    _PhysPkgReader,
    _PhysPkgWriter,
    _SerializedPart,
    _SerializedRelationship,
    _SerializedRelationshipCollection,
    _ZipPkgReader,
    _ZipPkgWriter,
)

from .unitdata.types import a_Default, a_Types, an_Override
from ..unitutil.file import absjoin, test_file_dir
from ..unitutil.mock import (
    MagicMock,
    Mock,
    call,
    class_mock,
    function_mock,
    instance_mock,
    method_mock,
    patch,
    property_mock,
)


test_pptx_path = absjoin(test_file_dir, "test.pptx")
dir_pkg_path = absjoin(test_file_dir, "expanded_pptx")
zip_pkg_path = test_pptx_path


class DescribePackageReader(object):
    """Unit-test suite for `pptx.opc.serialized.PackageReader` objects."""

    def it_knows_whether_it_contains_a_partname(self, _blob_reader_prop_):
        _blob_reader_prop_.return_value = set(("/ppt", "/docProps"))
        package_reader = PackageReader(None)

        assert "/ppt" in package_reader
        assert "/xyz" not in package_reader

    def it_can_get_a_blob_by_partname(self, _blob_reader_prop_):
        _blob_reader_prop_.return_value = {"/ppt/slides/slide1.xml": b"blob"}
        package_reader = PackageReader(None)

        assert package_reader["/ppt/slides/slide1.xml"] == b"blob"

    def it_can_get_the_rels_xml_for_a_partname(self, _blob_reader_prop_):
        _blob_reader_prop_.return_value = {"/ppt/_rels/presentation.xml.rels": b"blob"}
        package_reader = PackageReader(None)

        assert package_reader.rels_xml_for(PackURI("/ppt/presentation.xml")) == b"blob"

    # fixture components -----------------------------------

    @pytest.fixture
    def _blob_reader_prop_(self, request):
        return property_mock(request, PackageReader, "_blob_reader")


class Describe_ContentTypeMap(object):
    def it_can_construct_from_ct_item_xml(self, from_xml_fixture):
        content_types_xml, expected_defaults, expected_overrides = from_xml_fixture
        ct_map = _ContentTypeMap.from_xml(content_types_xml)
        assert ct_map._defaults == expected_defaults
        assert ct_map._overrides == expected_overrides

    def it_matches_an_override_on_case_insensitive_partname(
        self, match_override_fixture
    ):
        ct_map, partname, content_type = match_override_fixture
        assert ct_map[partname] == content_type

    def it_falls_back_to_case_insensitive_extension_default_match(
        self, match_default_fixture
    ):
        ct_map, partname, content_type = match_default_fixture
        assert ct_map[partname] == content_type

    def it_should_raise_on_partname_not_found(self):
        ct_map = _ContentTypeMap()
        with pytest.raises(KeyError):
            ct_map[PackURI("/!blat/rhumba.1x&")]

    def it_should_raise_on_key_not_instance_of_PackURI(self):
        ct_map = _ContentTypeMap()
        ct_map._add_override(PackURI("/part/name1.xml"), "app/vnd.type1")
        with pytest.raises(KeyError):
            ct_map["/part/name1.xml"]

    # fixtures ---------------------------------------------

    @pytest.fixture
    def from_xml_fixture(self):
        entries = (
            ("Default", "xml", CT.XML),
            ("Default", "PNG", CT.PNG),
            ("Override", "/ppt/presentation.xml", CT.PML_PRESENTATION_MAIN),
        )
        content_types_xml = self._xml_from(entries)
        expected_defaults = {}
        expected_overrides = {}
        for entry in entries:
            if entry[0] == "Default":
                ext = entry[1].lower()
                content_type = entry[2]
                expected_defaults[ext] = content_type
            elif entry[0] == "Override":
                partname, content_type = entry[1:]
                expected_overrides[partname] = content_type
        return content_types_xml, expected_defaults, expected_overrides

    @pytest.fixture(
        params=[
            ("/foo/bar.xml", "xml", "application/xml"),
            ("/foo/bar.PNG", "png", "image/png"),
            ("/foo/bar.jpg", "JPG", "image/jpeg"),
        ]
    )
    def match_default_fixture(self, request):
        partname_str, ext, content_type = request.param
        partname = PackURI(partname_str)
        ct_map = _ContentTypeMap()
        ct_map._add_override(PackURI("/bar/foo.xyz"), "application/xyz")
        ct_map._add_default(ext, content_type)
        return ct_map, partname, content_type

    @pytest.fixture(
        params=[
            ("/foo/bar.xml", "/foo/bar.xml"),
            ("/foo/bar.xml", "/FOO/Bar.XML"),
            ("/FoO/bAr.XmL", "/foo/bar.xml"),
        ]
    )
    def match_override_fixture(self, request):
        partname_str, should_match_partname_str = request.param
        partname = PackURI(partname_str)
        should_match_partname = PackURI(should_match_partname_str)
        content_type = "appl/vnd-foobar"
        ct_map = _ContentTypeMap()
        ct_map._add_override(partname, content_type)
        return ct_map, should_match_partname, content_type

    def _xml_from(self, entries):
        """
        Return XML for a [Content_Types].xml based on items in *entries*.
        """
        types_bldr = a_Types().with_nsdecls()
        for entry in entries:
            if entry[0] == "Default":
                ext, content_type = entry[1:]
                default_bldr = a_Default()
                default_bldr.with_Extension(ext)
                default_bldr.with_ContentType(content_type)
                types_bldr.with_child(default_bldr)
            elif entry[0] == "Override":
                partname, content_type = entry[1:]
                override_bldr = an_Override()
                override_bldr.with_PartName(partname)
                override_bldr.with_ContentType(content_type)
                types_bldr.with_child(override_bldr)
        return types_bldr.xml()


class DescribePackageWriter(object):
    def it_can_write_a_package(self, _PhysPkgWriter_, _write_methods):
        # mockery ----------------------
        pkg_file = Mock(name="pkg_file")
        pkg_rels = Mock(name="pkg_rels")
        parts = Mock(name="parts")
        phys_writer = _PhysPkgWriter_.return_value
        # exercise ---------------------
        PackageWriter.write(pkg_file, pkg_rels, parts)
        # verify -----------------------
        expected_calls = [
            call._write_content_types_stream(phys_writer, parts),
            call._write_pkg_rels(phys_writer, pkg_rels),
            call._write_parts(phys_writer, parts),
        ]
        _PhysPkgWriter_.assert_called_once_with(pkg_file)
        assert _write_methods.mock_calls == expected_calls
        phys_writer.close.assert_called_once_with()

    def it_can_write_a_content_types_stream(self, xml_for, serialize_part_xml_):
        # mockery ----------------------
        phys_writer = Mock(name="phys_writer")
        parts = Mock(name="parts")
        # exercise ---------------------
        PackageWriter._write_content_types_stream(phys_writer, parts)
        # verify -----------------------
        xml_for.assert_called_once_with(parts)
        serialize_part_xml_.assert_called_once_with(xml_for.return_value)
        phys_writer.write.assert_called_once_with(
            "/[Content_Types].xml", serialize_part_xml_.return_value
        )

    def it_can_write_a_pkg_rels_item(self):
        # mockery ----------------------
        phys_writer = Mock(name="phys_writer")
        pkg_rels = Mock(name="pkg_rels")
        # exercise ---------------------
        PackageWriter._write_pkg_rels(phys_writer, pkg_rels)
        # verify -----------------------
        phys_writer.write.assert_called_once_with("/_rels/.rels", pkg_rels.xml)

    def it_can_write_a_list_of_parts(self):
        # mockery ----------------------
        phys_writer = Mock(name="phys_writer")
        rels = MagicMock(name="rels")
        rels.__len__.return_value = 1
        part1 = Mock(name="part1", _rels=rels)
        part2 = Mock(name="part2", _rels=[])
        # exercise ---------------------
        PackageWriter._write_parts(phys_writer, [part1, part2])
        # verify -----------------------
        expected_calls = [
            call(part1.partname, part1.blob),
            call(part1.partname.rels_uri, part1._rels.xml),
            call(part2.partname, part2.blob),
        ]
        assert phys_writer.write.mock_calls == expected_calls

    # fixtures ---------------------------------------------

    @pytest.fixture
    def _PhysPkgWriter_(self, request):
        _patch = patch("pptx.opc.serialized._PhysPkgWriter")
        request.addfinalizer(_patch.stop)
        return _patch.start()

    @pytest.fixture
    def serialize_part_xml_(self, request):
        return function_mock(request, "pptx.opc.serialized.serialize_part_xml")

    @pytest.fixture
    def _write_methods(self, request):
        """Mock that patches all the _write_* methods of PackageWriter"""
        root_mock = Mock(name="PackageWriter")
        patch1 = patch.object(PackageWriter, "_write_content_types_stream")
        patch2 = patch.object(PackageWriter, "_write_pkg_rels")
        patch3 = patch.object(PackageWriter, "_write_parts")
        root_mock.attach_mock(patch1.start(), "_write_content_types_stream")
        root_mock.attach_mock(patch2.start(), "_write_pkg_rels")
        root_mock.attach_mock(patch3.start(), "_write_parts")

        def fin():
            patch1.stop()
            patch2.stop()
            patch3.stop()

        request.addfinalizer(fin)
        return root_mock

    @pytest.fixture
    def xml_for(self, request):
        return method_mock(request, _ContentTypesItem, "xml_for")


class Describe_PhysPkgReader(object):
    """Unit-test suite for `pptx.opc.serialized._PhysPkgReader` objects."""

    def it_constructs_ZipPkgReader_when_pkg_is_file_like(
        self, _ZipPkgReader_, zip_pkg_reader_
    ):
        _ZipPkgReader_.return_value = zip_pkg_reader_
        file_like_pkg = BytesIO(b"pkg-bytes")

        phys_reader = _PhysPkgReader.factory(file_like_pkg)

        _ZipPkgReader_.assert_called_once_with(file_like_pkg)
        assert phys_reader is zip_pkg_reader_

    def and_it_constructs_DirPkgReader_when_pkg_is_a_dir(self, request):
        dir_pkg_reader_ = instance_mock(request, _DirPkgReader)
        _DirPkgReader_ = class_mock(
            request, "pptx.opc.serialized._DirPkgReader", return_value=dir_pkg_reader_
        )

        phys_reader = _PhysPkgReader.factory(dir_pkg_path)

        _DirPkgReader_.assert_called_once_with(dir_pkg_path)
        assert phys_reader is dir_pkg_reader_

    def and_it_constructs_ZipPkgReader_when_pkg_is_a_zip_file_path(
        self, _ZipPkgReader_, zip_pkg_reader_
    ):
        _ZipPkgReader_.return_value = zip_pkg_reader_
        pkg_file_path = test_pptx_path

        phys_reader = _PhysPkgReader.factory(pkg_file_path)

        _ZipPkgReader_.assert_called_once_with(pkg_file_path)
        assert phys_reader is zip_pkg_reader_

    def but_it_raises_when_pkg_path_is_not_a_package(self):
        with pytest.raises(PackageNotFoundError) as e:
            _PhysPkgReader.factory("foobar")
        assert str(e.value) == "Package not found at 'foobar'"

    # --- fixture components -------------------------------

    @pytest.fixture
    def zip_pkg_reader_(self, request):
        return instance_mock(request, _ZipPkgReader)

    @pytest.fixture
    def _ZipPkgReader_(self, request):
        return class_mock(request, "pptx.opc.serialized._ZipPkgReader")


class Describe_DirPkgReader(object):
    """Unit-test suite for `pptx.opc.serialized._DirPkgReader` objects."""

    def it_can_retrieve_the_blob_for_a_pack_uri(self):
        blob = _DirPkgReader(dir_pkg_path)[PackURI("/ppt/presentation.xml")]

        sha1 = hashlib.sha1(blob).hexdigest()
        assert sha1 == "51b78f4dabc0af2419d4e044ab73028c4bef53aa"

    def but_it_raises_KeyError_when_requested_member_is_not_present(self):
        with pytest.raises(KeyError) as e:
            _DirPkgReader(dir_pkg_path)[PackURI("/ppt/foobar.xml")]
        assert str(e.value) == "\"no member '/ppt/foobar.xml' in package\""


class Describe_ZipPkgReader(object):
    """Unit-test suite for `pptx.opc.serialized._ZipPkgReader` objects."""

    def it_knows_whether_it_contains_a_partname(self, zip_pkg_reader):
        assert PackURI("/ppt/presentation.xml") in zip_pkg_reader
        assert PackURI("/ppt/foobar.xml") not in zip_pkg_reader

    def it_can_get_a_blob_by_partname(self, zip_pkg_reader):
        blob = zip_pkg_reader[PackURI("/ppt/presentation.xml")]
        assert hashlib.sha1(blob).hexdigest() == (
            "efa7bee0ac72464903a67a6744c1169035d52a54"
        )

    def but_it_raises_KeyError_when_requested_member_is_not_present(
        self, zip_pkg_reader
    ):
        with pytest.raises(KeyError) as e:
            zip_pkg_reader[PackURI("/ppt/foobar.xml")]
        assert str(e.value) == "\"no member '/ppt/foobar.xml' in package\""

    def it_loads_the_package_blobs_on_first_access_to_help(self, zip_pkg_reader):
        blobs = zip_pkg_reader._blobs
        assert len(blobs) == 38
        assert "/ppt/presentation.xml" in blobs
        assert "/ppt/_rels/presentation.xml.rels" in blobs

    # --- fixture components -------------------------------

    @pytest.fixture(scope="class")
    def zip_pkg_reader(self, request):
        return _ZipPkgReader(zip_pkg_path)


class Describe_ZipPkgWriter(object):
    def it_is_used_by_PhysPkgWriter_unconditionally(self, tmp_pptx_path):
        phys_writer = _PhysPkgWriter(tmp_pptx_path)
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
        pkg_writer = _PhysPkgWriter(pkg_file)
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


class Describe_SerializedPart(object):
    def it_remembers_construction_values(self):
        # test data --------------------
        partname = "/part/name.xml"
        content_type = "app/vnd.type"
        blob = "<Part/>"
        srels = "srels proxy"
        # exercise ---------------------
        spart = _SerializedPart(partname, content_type, blob, srels)
        # verify -----------------------
        assert spart.partname == partname
        assert spart.content_type == content_type
        assert spart.blob == blob
        assert spart.srels == srels


class Describe_SerializedRelationship(object):
    def it_remembers_construction_values(self):
        # test data --------------------
        rel_elm = CT_Relationship.new(
            "rId9", "ReLtYpE", "docProps/core.xml", RTM.INTERNAL
        )
        # exercise ---------------------
        srel = _SerializedRelationship("/", rel_elm)
        # verify -----------------------
        assert srel.rId == "rId9"
        assert srel.reltype == "ReLtYpE"
        assert srel.target_ref == "docProps/core.xml"
        assert srel.target_mode == RTM.INTERNAL

    def it_knows_when_it_is_external(self):
        cases = (RTM.INTERNAL, RTM.EXTERNAL)
        expected_values = (False, True)
        for target_mode, expected_value in zip(cases, expected_values):
            rel_elm = CT_Relationship.new(
                "rId9", "ReLtYpE", "docProps/core.xml", target_mode
            )
            srel = _SerializedRelationship(None, rel_elm)
            assert srel.is_external is expected_value

    def it_can_calculate_its_target_partname(self):
        # test data --------------------
        cases = (
            ("/", "docProps/core.xml", "/docProps/core.xml"),
            ("/ppt", "viewProps.xml", "/ppt/viewProps.xml"),
            (
                "/ppt/slides",
                "../slideLayouts/slideLayout1.xml",
                "/ppt/slideLayouts/slideLayout1.xml",
            ),
        )
        for baseURI, target_ref, expected_partname in cases:
            # setup --------------------
            rel_elm = Mock(
                name="rel_elm",
                rId=None,
                reltype=None,
                target_ref=target_ref,
                target_mode=RTM.INTERNAL,
            )
            # exercise -----------------
            srel = _SerializedRelationship(baseURI, rel_elm)
            # verify -------------------
            assert srel.target_partname == expected_partname

    def it_raises_on_target_partname_when_external(self):
        rel_elm = CT_Relationship.new(
            "rId9", "ReLtYpE", "docProps/core.xml", RTM.EXTERNAL
        )
        srel = _SerializedRelationship("/", rel_elm)
        with pytest.raises(ValueError):
            srel.target_partname


class Describe_SerializedRelationshipCollection(object):
    def it_can_load_from_xml(self, parse_xml, _SerializedRelationship_):
        # mockery ----------------------
        baseURI, rels_item_xml, rel_elm_1, rel_elm_2 = (
            Mock(name="baseURI"),
            Mock(name="rels_item_xml"),
            Mock(name="rel_elm_1"),
            Mock(name="rel_elm_2"),
        )
        rels_elm = Mock(name="rels_elm", relationship_lst=[rel_elm_1, rel_elm_2])
        parse_xml.return_value = rels_elm
        # exercise ---------------------
        srels = _SerializedRelationshipCollection.load_from_xml(baseURI, rels_item_xml)
        # verify -----------------------
        expected_calls = [call(baseURI, rel_elm_1), call(baseURI, rel_elm_2)]
        parse_xml.assert_called_once_with(rels_item_xml)
        assert _SerializedRelationship_.call_args_list == expected_calls
        assert isinstance(srels, _SerializedRelationshipCollection)

    def it_should_be_iterable(self):
        srels = _SerializedRelationshipCollection()
        try:
            for x in srels:
                pass
        except TypeError:
            msg = "_SerializedRelationshipCollection object is not iterable"
            pytest.fail(msg)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def parse_xml(self, request):
        return function_mock(request, "pptx.opc.serialized.parse_xml")

    @pytest.fixture
    def _SerializedRelationship_(self, request):
        return class_mock(request, "pptx.opc.serialized._SerializedRelationship")


class Describe_ContentTypesItem(object):
    def it_can_compose_content_types_xml(self, xml_for_fixture):
        parts, expected_xml = xml_for_fixture
        types_elm = _ContentTypesItem.xml_for(parts)
        assert types_elm.xml == expected_xml

    # fixtures ---------------------------------------------

    def _mock_part(self, request, name, partname_str, content_type):
        partname = PackURI(partname_str)
        return instance_mock(
            request, Part, name=name, partname=partname, content_type=content_type
        )

    @pytest.fixture(
        params=[
            ("Default", "/ppt/MEDIA/image.PNG", CT.PNG),
            ("Default", "/ppt/media/image.xml", CT.XML),
            ("Default", "/ppt/media/image.rels", CT.OPC_RELATIONSHIPS),
            ("Default", "/ppt/media/image.jpeg", CT.JPEG),
            ("Override", "/docProps/core.xml", "app/vnd.core"),
            ("Override", "/ppt/slides/slide1.xml", "app/vnd.ct_sld"),
            ("Override", "/zebra/foo.bar", "app/vnd.foobar"),
        ]
    )
    def xml_for_fixture(self, request):
        elm_type, partname_str, content_type = request.param
        part_ = self._mock_part(request, "part_", partname_str, content_type)
        # expected_xml -----------------
        types_bldr = a_Types().with_nsdecls()
        ext = partname_str.split(".")[-1].lower()
        if elm_type == "Default" and ext not in ("rels", "xml"):
            default_bldr = a_Default()
            default_bldr.with_Extension(ext)
            default_bldr.with_ContentType(content_type)
            types_bldr.with_child(default_bldr)

        types_bldr.with_child(
            a_Default().with_Extension("rels").with_ContentType(CT.OPC_RELATIONSHIPS)
        )
        types_bldr.with_child(
            a_Default().with_Extension("xml").with_ContentType(CT.XML)
        )

        if elm_type == "Override":
            override_bldr = an_Override()
            override_bldr.with_PartName(partname_str)
            override_bldr.with_ContentType(content_type)
            types_bldr.with_child(override_bldr)

        expected_xml = types_bldr.xml()
        return [part_], expected_xml


# fixtures -------------------------------------------------


@pytest.fixture
def tmp_pptx_path(tmpdir):
    return str(tmpdir.join("test_python-pptx.pptx"))


@pytest.fixture
def ZipFile_(request):
    return class_mock(request, "pptx.opc.serialized.zipfile.ZipFile")
