# encoding: utf-8

"""Unit-test suite for `pptx.opc.package` module."""

import collections
import io

import pytest

from pptx.opc.constants import (
    CONTENT_TYPE as CT,
    RELATIONSHIP_TARGET_MODE as RTM,
    RELATIONSHIP_TYPE as RT,
)
from pptx.opc.oxml import CT_Relationship
from pptx.opc.packuri import PACKAGE_URI, PackURI
from pptx.opc.package import (
    OpcPackage,
    Part,
    PartFactory,
    _Relationship,
    _Relationships,
    Unmarshaller,
    XmlPart,
)
from pptx.opc.serialized import PackageReader
from pptx.oxml import parse_xml
from pptx.oxml.xmlchemy import BaseOxmlElement
from pptx.package import Package

from ..unitutil.cxml import element
from ..unitutil.file import absjoin, snippet_bytes, test_file_dir
from ..unitutil.mock import (
    ANY,
    call,
    class_mock,
    cls_attr_mock,
    function_mock,
    initializer_mock,
    instance_mock,
    loose_mock,
    method_mock,
    Mock,
    patch,
    property_mock,
)


class DescribeOpcPackage(object):
    """Unit-test suite for `pptx.opc.package.OpcPackage` objects."""

    def it_can_open_a_pkg_file(self, request):
        package_ = instance_mock(request, OpcPackage)
        _init_ = initializer_mock(request, OpcPackage)
        _load_ = method_mock(request, OpcPackage, "_load", return_value=package_)

        package = OpcPackage.open("package.pptx")

        _init_.assert_called_once_with(ANY, "package.pptx")
        _load_.assert_called_once_with(ANY)
        assert package is package_

    def it_can_iterate_over_its_parts(self, request):
        part_, part_2_ = [
            instance_mock(request, Part, name="part_%d" % i) for i in range(2)
        ]
        rels_iter = (
            instance_mock(
                request, _Relationship, is_external=is_external, target_part=target
            )
            for is_external, target in (
                (True, "http://some/url/"),
                (False, part_),
                (False, part_),
                (False, part_2_),
                (False, part_),
                (False, part_2_),
            )
        )
        method_mock(request, OpcPackage, "iter_rels", return_value=rels_iter)
        package = OpcPackage(None)

        assert list(package.iter_parts()) == [part_, part_2_]

    def it_can_iterate_over_its_relationships(self, request, _rels_prop_):
        """
        +----------+          +--------+
        | pkg_rels |-- r0 --> | part_0 |
        +----------+          +--------+
             |     |            |    ^
          r2 |     | r1      r3 |    | r4
             |     |            |    |
             v     |            v    |
         external  |          +--------+
                   +--------> | part_1 |
                              +--------+
        """
        part_0_, part_1_ = [
            instance_mock(request, Part, name="part_%d" % i) for i in range(2)
        ]
        rels = tuple(
            instance_mock(
                request,
                _Relationship,
                name="r%d" % i,
                is_external=ext,
                target_part=part,
            )
            for i, (ext, part) in enumerate(
                (
                    (False, part_0_),
                    (False, part_1_),
                    (True, None),
                    (False, part_1_),
                    (False, part_0_),
                )
            )
        )
        _rels_prop_.return_value = rels[:3]
        part_0_.rels = rels[3:4]
        part_1_.rels = rels[4:]
        package = OpcPackage(None)

        assert tuple(package.iter_rels()) == (
            rels[0],
            rels[3],
            rels[4],
            rels[1],
            rels[2],
        )

    def it_can_add_a_relationship_to_a_part(self, request, _rels_prop_, relationships_):
        _rels_prop_.return_value = relationships_
        relationship_ = instance_mock(request, _Relationship)
        relationships_.add_relationship.return_value = relationship_
        target_ = instance_mock(request, Part, name="target_part")
        package = OpcPackage(None)

        relationship = package.load_rel(RT.SLIDE, target_, "rId99")

        relationships_.add_relationship.assert_called_once_with(
            RT.SLIDE, target_, "rId99", False
        )
        assert relationship is relationship_

    def it_can_establish_a_relationship_to_another_part(
        self, request, _rels_prop_, relationships_
    ):
        relationships_.get_or_add.return_value = "rId99"
        _rels_prop_.return_value = relationships_
        part_ = instance_mock(request, Part)
        package = OpcPackage(None)

        rId = package.relate_to(part_, "http://rel/type")

        relationships_.get_or_add.assert_called_once_with("http://rel/type", part_)
        assert rId == "rId99"

    def it_can_provide_a_list_of_the_parts_it_contains(self):
        # mockery ----------------------
        parts = [Mock(name="part1"), Mock(name="part2")]
        pkg = OpcPackage(None)
        # verify -----------------------
        with patch.object(OpcPackage, "iter_parts", return_value=parts):
            assert pkg.parts == [parts[0], parts[1]]

    def it_can_find_a_part_related_by_reltype(
        self, request, _rels_prop_, relationships_
    ):
        related_part_ = instance_mock(request, Part, name="related_part_")
        relationships_.part_with_reltype.return_value = related_part_
        _rels_prop_.return_value = relationships_
        package = OpcPackage(None)

        related_part = package.part_related_by(RT.SLIDE)

        relationships_.part_with_reltype.assert_called_once_with(RT.SLIDE)
        assert related_part is related_part_

    def it_can_find_the_next_available_vector_partname(self, next_partname_fixture):
        package, partname_template, expected_partname = next_partname_fixture
        partname = package.next_partname(partname_template)
        assert isinstance(partname, PackURI)
        assert partname == expected_partname

    def it_can_save_to_a_pkg_file(self, request, _rels_prop_, relationships_):
        PackageWriter_ = class_mock(request, "pptx.opc.package.PackageWriter")
        _rels_prop_.return_value = relationships_
        property_mock(request, OpcPackage, "parts", return_value=["parts"])
        package = OpcPackage(None)

        package.save("prs.pptx")

        PackageWriter_.write.assert_called_once_with(
            "prs.pptx", relationships_, ["parts"]
        )

    def it_loads_the_pkg_file_to_help(self, request):
        package_reader_ = instance_mock(request, PackageReader)
        PackageReader_ = class_mock(request, "pptx.opc.package.PackageReader")
        PackageReader_.from_file.return_value = package_reader_
        PartFactory_ = class_mock(request, "pptx.opc.package.PartFactory")
        Unmarshaller_ = class_mock(request, "pptx.opc.package.Unmarshaller")
        package = OpcPackage("prs.pptx")

        return_value = package._load()

        PackageReader_.from_file.assert_called_once_with("prs.pptx")
        Unmarshaller_.unmarshal.assert_called_once_with(
            package_reader_, package, PartFactory_
        )
        assert return_value is package

    def it_constructs_its_relationships_object_to_help(self, request, relationships_):
        _Relationships_ = class_mock(
            request, "pptx.opc.package._Relationships", return_value=relationships_
        )
        package = OpcPackage(None)

        rels = package._rels

        _Relationships_.assert_called_once_with(PACKAGE_URI.baseURI)
        assert rels is relationships_

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[((), 1), ((1,), 2), ((1, 2), 3), ((2, 3), 1), ((1, 3), 2)])
    def next_partname_fixture(self, request, iter_parts_):
        existing_partname_numbers, next_partname_number = request.param
        package = OpcPackage(None)
        parts = [
            instance_mock(
                request, Part, name="part[%d]" % idx, partname="/foo/bar/baz%d.xml" % n
            )
            for idx, n in enumerate(existing_partname_numbers)
        ]
        iter_parts_.return_value = iter(parts)
        partname_template = "/foo/bar/baz%d.xml"
        expected_partname = PackURI("/foo/bar/baz%d.xml" % next_partname_number)
        return package, partname_template, expected_partname

    # fixture components -----------------------------------

    @pytest.fixture
    def iter_parts_(self, request):
        return method_mock(request, OpcPackage, "iter_parts")

    @pytest.fixture
    def relationships_(self, request):
        return instance_mock(request, _Relationships)

    @pytest.fixture
    def _rels_prop_(self, request):
        return property_mock(request, OpcPackage, "_rels")


class DescribePart(object):
    """Unit-test suite for `pptx.opc.package.Part` objects."""

    def it_can_be_constructed_by_PartFactory(self, request, package_):
        partname_ = PackURI("/ppt/slides/slide1.xml")
        _init_ = initializer_mock(request, Part)

        part = Part.load(partname_, CT.PML_SLIDE, b"blob", package_)

        _init_.assert_called_once_with(part, partname_, CT.PML_SLIDE, b"blob", package_)
        assert isinstance(part, Part)

    def it_uses_the_load_blob_as_its_blob(self, blob_fixture):
        part, load_blob = blob_fixture
        assert part.blob is load_blob

    def it_can_change_its_blob(self):
        part, new_blob = Part(None, None, "xyz", None), "foobar"
        part.blob = new_blob
        assert part.blob == new_blob

    def it_knows_its_content_type(self, content_type_fixture):
        part, expected_content_type = content_type_fixture
        assert part.content_type == expected_content_type

    @pytest.mark.parametrize("ref_count, calls", ((2, []), (1, [call("rId42")])))
    def it_can_drop_a_relationship(
        self, request, _rels_prop_, relationships_, ref_count, calls
    ):
        _rel_ref_count_ = method_mock(
            request, Part, "_rel_ref_count", return_value=ref_count
        )
        _rels_prop_.return_value = relationships_
        part = Part(None, None, None)

        part.drop_rel("rId42")

        _rel_ref_count_.assert_called_once_with(part, "rId42")
        assert relationships_.pop.call_args_list == calls

    def it_can_load_a_relationship(self, load_rel_fixture):
        part, rels_, reltype_, target_, rId_ = load_rel_fixture

        part.load_rel(reltype_, target_, rId_)

        rels_.add_relationship.assert_called_once_with(reltype_, target_, rId_, False)

    def it_knows_the_package_it_belongs_to(self, package_get_fixture):
        part, expected_package = package_get_fixture
        assert part.package == expected_package

    def it_can_find_a_related_part_by_reltype(self, related_part_fixture):
        part, reltype_, related_part_ = related_part_fixture

        related_part = part.part_related_by(reltype_)

        part.rels.part_with_reltype.assert_called_once_with(reltype_)
        assert related_part is related_part_

    def it_knows_its_partname(self, partname_get_fixture):
        part, expected_partname = partname_get_fixture
        assert part.partname == expected_partname

    def it_can_change_its_partname(self, partname_set_fixture):
        part, new_partname = partname_set_fixture
        part.partname = new_partname
        assert part.partname == new_partname

    def it_can_establish_a_relationship_to_another_part(
        self, _rels_prop_, relationships_, part_
    ):
        relationships_.get_or_add.return_value = "rId42"
        _rels_prop_.return_value = relationships_
        part = Part(None, None, None)

        rId = part.relate_to(part_, RT.SLIDE)

        relationships_.get_or_add.assert_called_once_with(RT.SLIDE, part_)
        assert rId == "rId42"

    def it_can_establish_an_external_relationship(self, relate_to_url_fixture):
        part, url_, reltype_, rId_ = relate_to_url_fixture

        rId = part.relate_to(url_, reltype_, is_external=True)

        part.rels.get_or_add_ext_rel.assert_called_once_with(reltype_, url_)
        assert rId is rId_

    def it_can_find_a_related_part_by_rId(
        self, request, _rels_prop_, relationships_, relationship_, part_
    ):
        relationship_.target_part = part_
        relationships_.__getitem__.return_value = relationship_
        _rels_prop_.return_value = relationships_
        part = Part(None, None, None)

        related_part = part.related_part("rId17")

        relationships_.__getitem__.assert_called_once_with("rId17")
        assert related_part is part_

    def it_provides_access_to_its_relationships(self, rels_fixture):
        part, Relationships_, partname_, rels_ = rels_fixture

        rels = part.rels

        Relationships_.assert_called_once_with(partname_.baseURI)
        assert rels is rels_

    def it_can_find_the_uri_of_an_external_relationship(self, target_ref_fixture):
        part, rId_, url_ = target_ref_fixture

        url = part.target_ref(rId_)

        assert url == url_

    def it_can_load_a_blob_from_a_file_path_to_help(self):
        path = absjoin(test_file_dir, "minimal.pptx")
        with open(path, "rb") as f:
            file_bytes = f.read()
        part = Part(None, None, None, None)

        assert part._blob_from_file(path) == file_bytes

    def it_can_load_a_blob_from_a_file_like_object_to_help(self):
        part = Part(None, None, None, None)
        assert part._blob_from_file(io.BytesIO(b"012345")) == b"012345"

    # fixtures ---------------------------------------------

    @pytest.fixture
    def blob_fixture(self, blob_):
        part = Part(None, None, blob_, None)
        return part, blob_

    @pytest.fixture
    def content_type_fixture(self):
        content_type = "content/type"
        part = Part(None, content_type, None, None)
        return part, content_type

    @pytest.fixture
    def load_rel_fixture(self, part, _rels_prop_, rels_, reltype_, part_, rId_):
        _rels_prop_.return_value = rels_
        return part, rels_, reltype_, part_, rId_

    @pytest.fixture
    def package_get_fixture(self, package_):
        part = Part(None, None, None, package_)
        return part, package_

    @pytest.fixture
    def partname_get_fixture(self):
        partname = PackURI("/part/name")
        part = Part(partname, None, None, None)
        return part, partname

    @pytest.fixture
    def partname_set_fixture(self):
        old_partname = PackURI("/old/part/name")
        new_partname = PackURI("/new/part/name")
        part = Part(old_partname, None, None, None)
        return part, new_partname

    @pytest.fixture
    def relate_to_url_fixture(self, part, _rels_prop_, rels_, url_, reltype_, rId_):
        _rels_prop_.return_value = rels_
        return part, url_, reltype_, rId_

    @pytest.fixture
    def related_part_fixture(self, part, _rels_prop_, rels_, reltype_, part_):
        _rels_prop_.return_value = rels_
        return part, reltype_, part_

    @pytest.fixture
    def rels_fixture(self, Relationships_, partname_, rels_):
        part = Part(partname_, None)
        return part, Relationships_, partname_, rels_

    @pytest.fixture
    def target_ref_fixture(self, part, _rels_prop_, rId_, rel_, url_):
        _rels_prop_.return_value = {rId_: rel_}
        return part, rId_, url_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def blob_(self, request):
        return instance_mock(request, bytes)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, OpcPackage)

    @pytest.fixture
    def part(self):
        return Part(None, None)

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, Part)

    @pytest.fixture
    def partname_(self, request):
        return instance_mock(request, PackURI)

    @pytest.fixture
    def Relationships_(self, request, rels_):
        return class_mock(
            request, "pptx.opc.package._Relationships", return_value=rels_
        )

    @pytest.fixture
    def rel_(self, request, rId_, url_):
        return instance_mock(request, _Relationship, rId=rId_, target_ref=url_)

    @pytest.fixture
    def relationship_(self, request):
        return instance_mock(request, _Relationship)

    @pytest.fixture
    def relationships_(self, request):
        return instance_mock(request, _Relationships)

    @pytest.fixture
    def rels_(self, request, part_, rel_, rId_):
        rels_ = instance_mock(request, _Relationships)
        rels_.part_with_reltype.return_value = part_
        rels_.get_or_add.return_value = rel_
        rels_.get_or_add_ext_rel.return_value = rId_
        return rels_

    @pytest.fixture
    def _rels_prop_(self, request):
        return property_mock(request, Part, "_rels")

    @pytest.fixture
    def reltype_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def rId_(self, request):
        return instance_mock(request, str)

    @pytest.fixture
    def url_(self, request):
        return instance_mock(request, str)


class DescribeXmlPart(object):
    """Unit-test suite for `pptx.opc.package.XmlPart` objects."""

    def it_can_be_constructed_by_PartFactory(self, request):
        partname = PackURI("/ppt/slides/slide1.xml")
        element_ = element("p:sld")
        package_ = instance_mock(request, OpcPackage)
        parse_xml_ = function_mock(
            request, "pptx.opc.package.parse_xml", return_value=element_
        )
        _init_ = initializer_mock(request, XmlPart)

        part = XmlPart.load(partname, CT.PML_SLIDE, b"blob", package_)

        parse_xml_.assert_called_once_with(b"blob")
        _init_.assert_called_once_with(part, partname, CT.PML_SLIDE, element_, package_)
        assert isinstance(part, XmlPart)

    def it_can_serialize_to_xml(self, blob_fixture):
        xml_part, element_, serialize_part_xml_ = blob_fixture
        blob = xml_part.blob
        serialize_part_xml_.assert_called_once_with(element_)
        assert blob is serialize_part_xml_.return_value

    def it_knows_its_the_part_for_its_child_objects(self, part_fixture):
        xml_part = part_fixture
        assert xml_part.part is xml_part

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def blob_fixture(self, request, element_, serialize_part_xml_):
        xml_part = XmlPart(None, None, element_, None)
        return xml_part, element_, serialize_part_xml_

    @pytest.fixture
    def part_fixture(self):
        return XmlPart(None, None, None, None)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def element_(self, request):
        return instance_mock(request, BaseOxmlElement)

    @pytest.fixture
    def serialize_part_xml_(self, request):
        return function_mock(request, "pptx.opc.package.serialize_part_xml")


class DescribePartFactory(object):
    """Unit-test suite for `pptx.opc.package.PartFactory` objects."""

    def it_constructs_custom_part_type_for_registered_content_types(
        self, part_args_, CustomPartClass_, part_of_custom_type_
    ):
        # fixture ----------------------
        partname, content_type, pkg, blob = part_args_
        # exercise ---------------------
        PartFactory.part_type_for[content_type] = CustomPartClass_
        part = PartFactory(partname, content_type, pkg, blob)
        # verify -----------------------
        CustomPartClass_.load.assert_called_once_with(partname, content_type, pkg, blob)
        assert part is part_of_custom_type_

    def it_constructs_part_using_default_class_when_no_custom_registered(
        self, part_args_2_, DefaultPartClass_, part_of_default_type_
    ):
        partname, content_type, pkg, blob = part_args_2_
        part = PartFactory(partname, content_type, pkg, blob)
        DefaultPartClass_.load.assert_called_once_with(
            partname, content_type, pkg, blob
        )
        assert part is part_of_default_type_

    # fixtures ---------------------------------------------

    @pytest.fixture
    def part_of_custom_type_(self, request):
        return instance_mock(request, Part)

    @pytest.fixture
    def CustomPartClass_(self, request, part_of_custom_type_):
        CustomPartClass_ = Mock(name="CustomPartClass", spec=Part)
        CustomPartClass_.load.return_value = part_of_custom_type_
        return CustomPartClass_

    @pytest.fixture
    def part_of_default_type_(self, request):
        return instance_mock(request, Part)

    @pytest.fixture
    def DefaultPartClass_(self, request, part_of_default_type_):
        DefaultPartClass_ = cls_attr_mock(request, PartFactory, "default_part_type")
        DefaultPartClass_.load.return_value = part_of_default_type_
        return DefaultPartClass_

    @pytest.fixture
    def part_args_(self, request):
        partname_ = PackURI("/foo/bar.xml")
        content_type_ = "content/type"
        pkg_ = instance_mock(request, Package, name="pkg_")
        blob_ = b"blob_"
        return partname_, content_type_, pkg_, blob_

    @pytest.fixture
    def part_args_2_(self, request):
        partname_2_ = PackURI("/bar/foo.xml")
        content_type_2_ = "foobar/type"
        pkg_2_ = instance_mock(request, Package, name="pkg_2_")
        blob_2_ = b"blob_2_"
        return partname_2_, content_type_2_, pkg_2_, blob_2_


class Describe_Relationships(object):
    """Unit-test suite for `pptx.opc.package._Relationships` objects."""

    def it_has_dict_style_lookup_of_rel_by_rId(self, _rels_prop_, relationship_):
        _rels_prop_.return_value = {"rId17": relationship_}
        assert _Relationships(None)["rId17"] is relationship_

    def but_it_raises_KeyError_when_no_relationship_has_rId(self, _rels_prop_):
        _rels_prop_.return_value = {}
        with pytest.raises(KeyError) as e:
            _Relationships(None)["rId6"]
        assert str(e.value) == "\"no relationship with key 'rId6'\""

    def it_can_iterate_the_relationships_it_contains(self, request, _rels_prop_):
        rels_ = set(instance_mock(request, _Relationship) for n in range(5))
        _rels_prop_.return_value = {"rId%d" % (i + 1): r for i, r in enumerate(rels_)}
        relationships = _Relationships(None)

        for r in relationships:
            rels_.remove(r)

        assert len(rels_) == 0

    def it_has_a_len(self, _rels_prop_):
        _rels_prop_.return_value = {"a": 0, "b": 1}
        assert len(_Relationships(None)) == 2

    def it_can_add_a_relationship_to_a_target_part(
        self, part_, _get_matching_, _add_relationship_
    ):
        _get_matching_.return_value = None
        _add_relationship_.return_value = "rId7"
        relationships = _Relationships(None)

        rId = relationships.get_or_add(RT.IMAGE, part_)

        _get_matching_.assert_called_once_with(relationships, RT.IMAGE, part_)
        _add_relationship_.assert_called_once_with(relationships, RT.IMAGE, part_)
        assert rId == "rId7"

    def but_it_returns_an_existing_relationship_if_it_matches(
        self, part_, _get_matching_
    ):
        _get_matching_.return_value = "rId3"
        relationships = _Relationships(None)

        rId = relationships.get_or_add(RT.IMAGE, part_)

        _get_matching_.assert_called_once_with(relationships, RT.IMAGE, part_)
        assert rId == "rId3"

    def it_can_add_an_external_relationship_to_a_URI(
        self, _get_matching_, _add_relationship_
    ):
        _get_matching_.return_value = None
        _add_relationship_.return_value = "rId2"
        relationships = _Relationships(None)

        rId = relationships.get_or_add_ext_rel(RT.HYPERLINK, "http://url")

        _get_matching_.assert_called_once_with(
            relationships, RT.HYPERLINK, "http://url", is_external=True
        )
        _add_relationship_.assert_called_once_with(
            relationships, RT.HYPERLINK, "http://url", is_external=True
        )
        assert rId == "rId2"

    def but_it_returns_an_existing_external_relationship_if_it_matches(
        self, part_, _get_matching_
    ):
        _get_matching_.return_value = "rId10"
        relationships = _Relationships(None)

        rId = relationships.get_or_add_ext_rel(RT.HYPERLINK, "http://url")

        _get_matching_.assert_called_once_with(
            relationships, RT.HYPERLINK, "http://url", is_external=True
        )
        assert rId == "rId10"

    def it_can_load_from_the_xml_in_a_rels_part(self, request, _Relationship_, part_):
        rels_ = tuple(
            instance_mock(request, _Relationship, rId="rId%d" % (i + 1))
            for i in range(5)
        )
        _Relationship_.from_xml.side_effect = iter(rels_)
        parts = {"/ppt/slideLayouts/slideLayout1.xml": part_}
        xml_rels = parse_xml(snippet_bytes("rels-load-from-xml"))
        relationships = _Relationships(None)

        relationships.load_from_xml("/ppt/slides", xml_rels, parts)

        assert _Relationship_.from_xml.call_args_list == [
            call("/ppt/slides", xml_rels[0], parts),
            call("/ppt/slides", xml_rels[1], parts),
        ]
        assert relationships._rels == {"rId1": rels_[0], "rId2": rels_[1]}

    def it_can_find_a_part_with_reltype(
        self, _rels_by_reltype_prop_, relationship_, part_
    ):
        relationship_.target_part = part_
        _rels_by_reltype_prop_.return_value = collections.defaultdict(
            list, ((RT.SLIDE_LAYOUT, [relationship_]),)
        )
        relationships = _Relationships(None)

        assert relationships.part_with_reltype(RT.SLIDE_LAYOUT) is part_

    def but_it_raises_KeyError_when_there_is_no_such_part(self, _rels_by_reltype_prop_):
        _rels_by_reltype_prop_.return_value = collections.defaultdict(list)
        relationships = _Relationships(None)

        with pytest.raises(KeyError) as e:
            relationships.part_with_reltype(RT.SLIDE_LAYOUT)
        assert str(e.value) == (
            "\"no relationship of type 'http://schemas.openxmlformats.org/"
            "officeDocument/2006/relationships/slideLayout' in collection\""
        )

    def and_it_raises_ValueError_when_there_is_more_than_one_part_with_reltype(
        self, _rels_by_reltype_prop_, relationship_, part_
    ):
        relationship_.target_part = part_
        _rels_by_reltype_prop_.return_value = collections.defaultdict(
            list, ((RT.SLIDE_LAYOUT, [relationship_, relationship_]),)
        )
        relationships = _Relationships(None)

        with pytest.raises(ValueError) as e:
            relationships.part_with_reltype(RT.SLIDE_LAYOUT)
        assert str(e.value) == (
            "multiple relationships of type 'http://schemas.openxmlformats.org/"
            "officeDocument/2006/relationships/slideLayout' in collection"
        )

    def it_can_pop_a_relationship_to_remove_it_from_the_collection(
        self, _rels_prop_, relationship_
    ):
        _rels_prop_.return_value = {"rId22": relationship_}
        relationships = _Relationships(None)

        relationships.pop("rId22")

        assert relationships._rels == {}

    def it_can_serialize_itself_to_XML(self, request, _rels_prop_):
        _rels_prop_.return_value = {
            "rId1": instance_mock(
                request,
                _Relationship,
                rId="rId1",
                reltype=RT.SLIDE,
                target_ref="../slides/slide1.xml",
                is_external=False,
            ),
            "rId2": instance_mock(
                request,
                _Relationship,
                rId="rId2",
                reltype=RT.HYPERLINK,
                target_ref="http://url",
                is_external=True,
            ),
        }
        relationships = _Relationships(None)

        assert relationships.xml == snippet_bytes("relationships")

    def it_can_add_a_relationship_to_a_part_to_help(
        self,
        request,
        _next_rId_prop_,
        _Relationship_,
        relationship_,
        _rels_prop_,
        part_,
    ):
        _next_rId_prop_.return_value = "rId8"
        _Relationship_.return_value = relationship_
        _rels_prop_.return_value = {}
        relationships = _Relationships("/ppt")

        rId = relationships._add_relationship(RT.SLIDE, part_)

        _Relationship_.assert_called_once_with(
            "/ppt", "rId8", RT.SLIDE, target_mode=RTM.INTERNAL, target=part_
        )
        assert relationships._rels == {"rId8": relationship_}
        assert rId == "rId8"

    def and_it_can_add_an_external_relationship_to_help(
        self, request, _next_rId_prop_, _rels_prop_, _Relationship_, relationship_
    ):
        _next_rId_prop_.return_value = "rId9"
        _Relationship_.return_value = relationship_
        _rels_prop_.return_value = {}
        relationships = _Relationships("/ppt")

        rId = relationships._add_relationship(
            RT.HYPERLINK, "http://url", is_external=True
        )

        _Relationship_.assert_called_once_with(
            "/ppt", "rId9", RT.HYPERLINK, target_mode=RTM.EXTERNAL, target="http://url"
        )
        assert relationships._rels == {"rId9": relationship_}
        assert rId == "rId9"

    def it_can_get_a_matching_relationship_to_help(
        self, _rels_by_reltype_prop_, relationship_, part_
    ):
        relationship_.is_external = False
        relationship_.target_part = part_
        relationship_.rId = "rId10"
        _rels_by_reltype_prop_.return_value = {RT.SLIDE: [relationship_]}
        relationships = _Relationships(None)

        assert relationships._get_matching(RT.SLIDE, part_) == "rId10"

    def but_it_returns_None_when_there_is_no_matching_relationship(
        self, _rels_by_reltype_prop_
    ):
        _rels_by_reltype_prop_.return_value = collections.defaultdict(list)
        relationships = _Relationships(None)

        assert relationships._get_matching(RT.HYPERLINK, "http://url", True) is None

    @pytest.mark.parametrize(
        "rIds, expected_value",
        (
            ((), "rId1"),
            (("rId1",), "rId2"),
            (("rId1", "rId2"), "rId3"),
            (("rId1", "rId4"), "rId3"),
            (("rId1", "rId4", "rId6"), "rId3"),
            (("rId1", "rId2", "rId6"), "rId4"),
        ),
    )
    def it_finds_the_next_rId_to_help(self, _rels_prop_, rIds, expected_value):
        _rels_prop_.return_value = {rId: None for rId in rIds}
        relationships = _Relationships(None)

        assert relationships._next_rId == expected_value

    def it_collects_relationships_by_reltype_to_help(self, request, _rels_prop_):
        rels = {
            "rId%d" % (i + 1): instance_mock(request, _Relationship, reltype=reltype)
            for i, reltype in enumerate((RT.SLIDE, RT.IMAGE, RT.SLIDE, RT.HYPERLINK))
        }
        _rels_prop_.return_value = rels
        relationships = _Relationships(None)

        rels_by_reltype = relationships._rels_by_reltype

        assert rels["rId1"] in rels_by_reltype[RT.SLIDE]
        assert rels["rId2"] in rels_by_reltype[RT.IMAGE]
        assert rels["rId3"] in rels_by_reltype[RT.SLIDE]
        assert rels["rId4"] in rels_by_reltype[RT.HYPERLINK]
        assert rels_by_reltype[RT.CHART] == []

    # fixture components -----------------------------------

    @pytest.fixture
    def _add_relationship_(self, request):
        return method_mock(request, _Relationships, "_add_relationship")

    @pytest.fixture
    def _get_matching_(self, request):
        return method_mock(request, _Relationships, "_get_matching")

    @pytest.fixture
    def _next_rId_prop_(self, request):
        return property_mock(request, _Relationships, "_next_rId")

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, Part)

    @pytest.fixture
    def _Relationship_(self, request):
        return class_mock(request, "pptx.opc.package._Relationship")

    @pytest.fixture
    def relationship_(self, request):
        return instance_mock(request, _Relationship)

    @pytest.fixture
    def _rels_by_reltype_prop_(self, request):
        return property_mock(request, _Relationships, "_rels_by_reltype")

    @pytest.fixture
    def _rels_prop_(self, request):
        return property_mock(request, _Relationships, "_rels")


class Describe_Relationship(object):
    """Unit-test suite for `pptx.opc.package._Relationship` objects."""

    def it_can_construct_from_xml(self, request, part_):
        _init_ = initializer_mock(request, _Relationship)
        rel_elm = instance_mock(
            request,
            CT_Relationship,
            rId="rId42",
            reltype=RT.SLIDE,
            targetMode=RTM.INTERNAL,
            target_ref="slides/slide7.xml",
        )
        parts = {"/ppt/slides/slide7.xml": part_}

        relationship = _Relationship.from_xml("/ppt", rel_elm, parts)

        _init_.assert_called_once_with(
            relationship,
            "/ppt",
            "rId42",
            RT.SLIDE,
            RTM.INTERNAL,
            part_,
        )
        assert isinstance(relationship, _Relationship)

    @pytest.mark.parametrize(
        "target_mode, expected_value",
        ((RTM.INTERNAL, False), (RTM.EXTERNAL, True), (None, False)),
    )
    def it_knows_whether_it_is_external(self, target_mode, expected_value):
        relationship = _Relationship(None, None, None, target_mode, None)
        assert relationship.is_external == expected_value

    def it_knows_its_relationship_type(self):
        relationship = _Relationship(None, None, RT.SLIDE, None, None)
        assert relationship.reltype == RT.SLIDE

    def it_knows_its_rId(self):
        relationship = _Relationship(None, "rId42", None, None, None)
        assert relationship.rId == "rId42"

    def it_provides_access_to_its_target_part(self, part_):
        relationship = _Relationship(None, None, None, RTM.INTERNAL, part_)
        assert relationship.target_part is part_

    def but_it_raises_ValueError_on_target_part_for_external_rel(self):
        relationship = _Relationship(None, None, None, RTM.EXTERNAL, None)
        with pytest.raises(ValueError) as e:
            relationship.target_part
        assert str(e.value) == (
            "`.target_part` property on _Relationship is undefined when "
            "target-mode is external"
        )

    def it_knows_its_target_partname(self, part_):
        part_.partname = PackURI("/ppt/slideLayouts/slideLayout4.xml")
        relationship = _Relationship(None, None, None, RTM.INTERNAL, part_)

        assert relationship.target_partname == "/ppt/slideLayouts/slideLayout4.xml"

    def but_it_raises_ValueError_on_target_partname_for_external_rel(self):
        relationship = _Relationship(None, None, None, RTM.EXTERNAL, None)

        with pytest.raises(ValueError) as e:
            relationship.target_partname

        assert str(e.value) == (
            "`.target_partname` property on _Relationship is undefined when "
            "target-mode is external"
        )

    def it_knows_the_target_uri_for_an_external_rel(self):
        relationship = _Relationship(None, None, None, RTM.EXTERNAL, "http://url")
        assert relationship.target_ref == "http://url"

    def and_it_knows_the_relative_partname_for_an_internal_rel(self, request):
        """Internal relationships have a relative reference for `.target_ref`.

        A relative reference looks like "../slideLayouts/slideLayout1.xml". This form
        is suitable for writing to a .rels file.
        """
        property_mock(
            request,
            _Relationship,
            "target_partname",
            return_value=PackURI("/ppt/media/image1.png"),
        )
        relationship = _Relationship("/ppt/slides", None, None, None, None)

        assert relationship.target_ref == "../media/image1.png"

    # --- fixture components -------------------------------

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, Part)


class DescribeUnmarshaller(object):
    def it_can_unmarshal_from_a_pkg_reader(
        self,
        pkg_reader_,
        pkg_,
        part_factory_,
        _unmarshal_parts,
        _unmarshal_relationships,
        parts_dict_,
    ):
        Unmarshaller.unmarshal(pkg_reader_, pkg_, part_factory_)

        _unmarshal_parts.assert_called_once_with(pkg_reader_, pkg_, part_factory_)
        _unmarshal_relationships.assert_called_once_with(pkg_reader_, pkg_, parts_dict_)

    def it_can_unmarshal_parts(
        self,
        pkg_reader_,
        pkg_,
        part_factory_,
        parts_dict_,
        partnames_,
        content_types_,
        blobs_,
    ):
        # fixture ----------------------
        partname_, partname_2_ = partnames_
        content_type_, content_type_2_ = content_types_
        blob_, blob_2_ = blobs_
        # exercise ---------------------
        parts = Unmarshaller._unmarshal_parts(pkg_reader_, pkg_, part_factory_)
        # verify -----------------------
        assert part_factory_.call_args_list == [
            call(partname_, content_type_, blob_, pkg_),
            call(partname_2_, content_type_2_, blob_2_, pkg_),
        ]
        assert parts == parts_dict_

    def it_can_unmarshal_relationships(self):
        # test data --------------------
        reltype = "http://reltype"
        # mockery ----------------------
        pkg_reader = Mock(name="pkg_reader")
        pkg_reader.iter_srels.return_value = (
            (
                "/",
                Mock(
                    name="srel1",
                    rId="rId1",
                    reltype=reltype,
                    target_partname="partname1",
                    is_external=False,
                ),
            ),
            (
                "/",
                Mock(
                    name="srel2",
                    rId="rId2",
                    reltype=reltype,
                    target_ref="target_ref_1",
                    is_external=True,
                ),
            ),
            (
                "partname1",
                Mock(
                    name="srel3",
                    rId="rId3",
                    reltype=reltype,
                    target_partname="partname2",
                    is_external=False,
                ),
            ),
            (
                "partname2",
                Mock(
                    name="srel4",
                    rId="rId4",
                    reltype=reltype,
                    target_ref="target_ref_2",
                    is_external=True,
                ),
            ),
        )
        pkg = Mock(name="pkg")
        parts = {}
        for num in range(1, 3):
            name = "part%d" % num
            part = Mock(name=name)
            parts["partname%d" % num] = part
            pkg.attach_mock(part, name)
        # exercise ---------------------
        Unmarshaller._unmarshal_relationships(pkg_reader, pkg, parts)
        # verify -----------------------
        expected_pkg_calls = [
            call.load_rel(reltype, parts["partname1"], "rId1", False),
            call.load_rel(reltype, "target_ref_1", "rId2", True),
            call.part1.load_rel(reltype, parts["partname2"], "rId3", False),
            call.part2.load_rel(reltype, "target_ref_2", "rId4", True),
        ]
        assert pkg.mock_calls == expected_pkg_calls

    # fixtures ---------------------------------------------

    @pytest.fixture
    def blobs_(self, request):
        blob_ = loose_mock(request, spec=str, name="blob_")
        blob_2_ = loose_mock(request, spec=str, name="blob_2_")
        return blob_, blob_2_

    @pytest.fixture
    def content_types_(self, request):
        content_type_ = loose_mock(request, spec=str, name="content_type_")
        content_type_2_ = loose_mock(request, spec=str, name="content_type_2_")
        return content_type_, content_type_2_

    @pytest.fixture
    def part_factory_(self, request, parts_):
        part_factory_ = loose_mock(request, spec=Part)
        part_factory_.side_effect = parts_
        return part_factory_

    @pytest.fixture
    def partnames_(self, request):
        partname_ = loose_mock(request, spec=str, name="partname_")
        partname_2_ = loose_mock(request, spec=str, name="partname_2_")
        return partname_, partname_2_

    @pytest.fixture
    def parts_(self, request):
        part_ = instance_mock(request, Part, name="part_")
        part_2_ = instance_mock(request, Part, name="part_2")
        return part_, part_2_

    @pytest.fixture
    def parts_dict_(self, request, partnames_, parts_):
        partname_, partname_2_ = partnames_
        part_, part_2_ = parts_
        return {partname_: part_, partname_2_: part_2_}

    @pytest.fixture
    def pkg_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def pkg_reader_(self, request, partnames_, content_types_, blobs_):
        partname_, partname_2_ = partnames_
        content_type_, content_type_2_ = content_types_
        blob_, blob_2_ = blobs_
        spart_return_values = (
            (partname_, content_type_, blob_),
            (partname_2_, content_type_2_, blob_2_),
        )
        pkg_reader_ = instance_mock(request, PackageReader)
        pkg_reader_.iter_sparts.return_value = spart_return_values
        return pkg_reader_

    @pytest.fixture
    def _unmarshal_parts(self, request, parts_dict_):
        return method_mock(
            request, Unmarshaller, "_unmarshal_parts", return_value=parts_dict_
        )

    @pytest.fixture
    def _unmarshal_relationships(self, request):
        return method_mock(request, Unmarshaller, "_unmarshal_relationships")
