# encoding: utf-8

"""Unit-test suite for `pptx.opc.serialized` module."""

import hashlib
import pytest
import zipfile

from pptx.compat import BytesIO
from pptx.exceptions import PackageNotFoundError
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part, _Relationships
from pptx.opc.packuri import CONTENT_TYPES_URI, PackURI
from pptx.opc.serialized import (
    PackageReader,
    PackageWriter,
    _ContentTypesItem,
    _DirPkgReader,
    _PhysPkgReader,
    _PhysPkgWriter,
    _ZipPkgReader,
    _ZipPkgWriter,
)

from ..unitutil.file import absjoin, snippet_text, test_file_dir
from ..unitutil.mock import (
    ANY,
    call,
    class_mock,
    function_mock,
    initializer_mock,
    instance_mock,
    method_mock,
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

    def but_it_returns_None_when_the_part_has_no_rels(self, _blob_reader_prop_):
        _blob_reader_prop_.return_value = {"/ppt/_rels/presentation.xml.rels": b"blob"}
        package_reader = PackageReader(None)

        assert package_reader.rels_xml_for(PackURI("/ppt/slides.slide1.xml")) is None

    def it_constructs_its_blob_reader_to_help(self, request):
        phys_pkg_reader_ = instance_mock(request, _PhysPkgReader)
        _PhysPkgReader_ = class_mock(request, "pptx.opc.serialized._PhysPkgReader")
        _PhysPkgReader_.factory.return_value = phys_pkg_reader_
        package_reader = PackageReader("prs.pptx")

        blob_reader = package_reader._blob_reader

        _PhysPkgReader_.factory.assert_called_once_with("prs.pptx")
        assert blob_reader is phys_pkg_reader_

    # fixture components -----------------------------------

    @pytest.fixture
    def _blob_reader_prop_(self, request):
        return property_mock(request, PackageReader, "_blob_reader")


class DescribePackageWriter(object):
    """Unit-test suite for `pptx.opc.serialized.PackageWriter` objects."""

    def it_provides_a_write_interface_classmethod(self, request, relationships_):
        _init_ = initializer_mock(request, PackageWriter)
        _write_ = method_mock(request, PackageWriter, "_write")

        PackageWriter.write("prs.pptx", relationships_, ("part_1", "part_2"))

        _init_.assert_called_once_with(
            ANY, "prs.pptx", relationships_, ("part_1", "part_2")
        )
        _write_.assert_called_once_with(ANY)

    def it_can_write_a_package(self, request, phys_writer_):
        _PhysPkgWriter_ = class_mock(request, "pptx.opc.serialized._PhysPkgWriter")
        phys_writer_.__enter__.return_value = phys_writer_
        _PhysPkgWriter_.factory.return_value = phys_writer_
        _write_content_types_stream_ = method_mock(
            request, PackageWriter, "_write_content_types_stream"
        )
        _write_pkg_rels_ = method_mock(request, PackageWriter, "_write_pkg_rels")
        _write_parts_ = method_mock(request, PackageWriter, "_write_parts")
        package_writer = PackageWriter("prs.pptx", None, None)

        package_writer._write()

        _PhysPkgWriter_.factory.assert_called_once_with("prs.pptx")
        _write_content_types_stream_.assert_called_once_with(
            package_writer, phys_writer_
        )
        _write_pkg_rels_.assert_called_once_with(package_writer, phys_writer_)
        _write_parts_.assert_called_once_with(package_writer, phys_writer_)

    def it_can_write_a_content_types_stream(self, request, phys_writer_):
        _ContentTypesItem_ = class_mock(
            request, "pptx.opc.serialized._ContentTypesItem"
        )
        _ContentTypesItem_.xml_for.return_value = "part_xml"
        serialize_part_xml_ = function_mock(
            request, "pptx.opc.serialized.serialize_part_xml", return_value=b"xml"
        )
        package_writer = PackageWriter(None, None, ("part_1", "part_2"))

        package_writer._write_content_types_stream(phys_writer_)

        _ContentTypesItem_.xml_for.assert_called_once_with(("part_1", "part_2"))
        serialize_part_xml_.assert_called_once_with("part_xml")
        phys_writer_.write.assert_called_once_with(CONTENT_TYPES_URI, b"xml")

    def it_can_write_a_sequence_of_parts(self, request, phys_writer_):
        parts_ = (
            instance_mock(
                request,
                Part,
                partname=PackURI("/ppt/%s.xml" % x),
                blob="blob_%s" % x,
                rels=instance_mock(request, _Relationships, xml="rels_xml_%s" % x),
            )
            for x in ("a", "b", "c")
        )
        package_writer = PackageWriter(None, None, parts_)

        package_writer._write_parts(phys_writer_)

        assert phys_writer_.write.call_args_list == [
            call("/ppt/a.xml", "blob_a"),
            call("/ppt/_rels/a.xml.rels", "rels_xml_a"),
            call("/ppt/b.xml", "blob_b"),
            call("/ppt/_rels/b.xml.rels", "rels_xml_b"),
            call("/ppt/c.xml", "blob_c"),
            call("/ppt/_rels/c.xml.rels", "rels_xml_c"),
        ]

    def it_can_write_a_pkg_rels_item(self, request, phys_writer_, relationships_):
        relationships_.xml = b"pkg-rels-xml"
        package_writer = PackageWriter(None, relationships_, None)

        package_writer._write_pkg_rels(phys_writer_)

        phys_writer_.write.assert_called_once_with("/_rels/.rels", b"pkg-rels-xml")

    # fixture components -----------------------------------

    @pytest.fixture
    def phys_writer_(self, request):
        return instance_mock(request, _ZipPkgWriter)

    @pytest.fixture
    def relationships_(self, request):
        return instance_mock(request, _Relationships)


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

    def it_knows_whether_it_contains_a_partname(self, dir_pkg_reader):
        assert PackURI("/ppt/presentation.xml") in dir_pkg_reader
        assert PackURI("/ppt/foobar.xml") not in dir_pkg_reader

    def it_can_retrieve_the_blob_for_a_pack_uri(self, dir_pkg_reader):
        blob = dir_pkg_reader[PackURI("/ppt/presentation.xml")]
        assert (
            hashlib.sha1(blob).hexdigest() == "51b78f4dabc0af2419d4e044ab73028c4bef53aa"
        )

    def but_it_raises_KeyError_when_requested_member_is_not_present(
        self, dir_pkg_reader
    ):
        with pytest.raises(KeyError) as e:
            dir_pkg_reader[PackURI("/ppt/foobar.xml")]
        assert str(e.value) == "\"no member '/ppt/foobar.xml' in package\""

    # --- fixture components -------------------------------

    @pytest.fixture(scope="class")
    def dir_pkg_reader(self, request):
        return _DirPkgReader(dir_pkg_path)


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


class Describe_PhysPkgWriter(object):
    """Unit-test suite for `pptx.opc.serialized._PhysPkgWriter` objects."""

    def it_constructs_ZipPkgWriter_unconditionally(self, request):
        zip_pkg_writer_ = instance_mock(request, _ZipPkgWriter)
        _ZipPkgWriter_ = class_mock(
            request, "pptx.opc.serialized._ZipPkgWriter", return_value=zip_pkg_writer_
        )

        phys_writer = _PhysPkgWriter.factory("prs.pptx")

        _ZipPkgWriter_.assert_called_once_with("prs.pptx")
        assert phys_writer is zip_pkg_writer_


class Describe_ZipPkgWriter(object):
    """Unit-test suite for `pptx.opc.serialized._ZipPkgWriter` objects."""

    def it_has_an__enter__method_for_context_management(self):
        pkg_writer = _ZipPkgWriter(None)
        assert pkg_writer.__enter__() is pkg_writer

    def and_it_closes_the_zip_archive_on_context__exit__(self, _zipf_prop_):
        _ZipPkgWriter(None).__exit__(None, None, None)
        _zipf_prop_.return_value.close.assert_called_once_with()

    def it_can_write_a_blob(self, _zipf_prop_):
        """Integrates with zipfile.ZipFile."""
        pack_uri = PackURI("/part/name.xml")
        _zipf_prop_.return_value = zipf = zipfile.ZipFile(BytesIO(), "w")
        pkg_writer = _ZipPkgWriter(None)

        pkg_writer.write(pack_uri, b"blob")

        members = {PackURI("/%s" % name): zipf.read(name) for name in zipf.namelist()}
        assert len(members) == 1
        assert members[pack_uri] == b"blob"

    def it_provides_access_to_the_open_zip_file_to_help(self, request):
        ZipFile_ = class_mock(request, "pptx.opc.serialized.zipfile.ZipFile")
        pkg_writer = _ZipPkgWriter("prs.pptx")

        zipf = pkg_writer._zipf

        ZipFile_.assert_called_once_with(
            "prs.pptx", "w", compression=zipfile.ZIP_DEFLATED
        )
        assert zipf is ZipFile_.return_value

    # fixtures ---------------------------------------------

    @pytest.fixture
    def _zipf_prop_(self, request):
        return property_mock(request, _ZipPkgWriter, "_zipf")


class Describe_ContentTypesItem(object):
    """Unit-test suite for `pptx.opc.serialized._ContentTypesItem` objects."""

    def it_provides_an_interface_classmethod(self, request):
        _init_ = initializer_mock(request, _ContentTypesItem)
        property_mock(request, _ContentTypesItem, "_xml", return_value=b"xml")

        xml = _ContentTypesItem.xml_for(("part", "zuh"))

        _init_.assert_called_once_with(ANY, ("part", "zuh"))
        assert xml == b"xml"

    def it_can_compose_content_types_xml(self, request):
        defaults = {"png": CT.PNG, "xml": CT.XML, "rels": CT.OPC_RELATIONSHIPS}
        overrides = {
            "/docProps/core.xml": "app/vnd.core",
            "/ppt/slides/slide1.xml": "app/vnd.ct_sld",
            "/zebra/foo.bar": "app/vnd.foobar",
        }
        property_mock(
            request,
            _ContentTypesItem,
            "_defaults_and_overrides",
            return_value=(defaults, overrides),
        )

        content_types = _ContentTypesItem(None)._xml

        assert content_types.xml == snippet_text("content-types-xml").strip()

    def it_computes_defaults_and_overrides_to_help(self, request):
        parts = (
            instance_mock(
                request, Part, partname=PackURI(partname), content_type=content_type
            )
            for partname, content_type in (
                ("/media/image1.png", CT.PNG),
                ("/ppt/slides/slide1.xml", CT.PML_SLIDE),
                ("/foo/bar.xml", CT.XML),
                ("/docProps/core.xml", CT.OPC_CORE_PROPERTIES),
            )
        )
        content_types = _ContentTypesItem(parts)

        defaults, overrides = content_types._defaults_and_overrides

        assert defaults == {"png": CT.PNG, "rels": CT.OPC_RELATIONSHIPS, "xml": CT.XML}
        assert overrides == {
            "/ppt/slides/slide1.xml": CT.PML_SLIDE,
            "/docProps/core.xml": CT.OPC_CORE_PROPERTIES,
        }
