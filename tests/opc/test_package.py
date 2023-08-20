# encoding: utf-8

"""Unit-test suite for `pptx.opc.package` module."""

import collections
import io
import itertools

import pytest

from pptx.opc.constants import (
    CONTENT_TYPE as CT,
    RELATIONSHIP_TARGET_MODE as RTM,
    RELATIONSHIP_TYPE as RT,
)
from pptx.opc.oxml import CT_Relationship, CT_Relationships
from pptx.opc.package import (
    OpcPackage,
    Part,
    PartFactory,
    XmlPart,
    _ContentTypeMap,
    _PackageLoader,
    _RelatableMixin,
    _Relationship,
    _Relationships,
)
from pptx.opc.packuri import PACKAGE_URI, PackURI
from pptx.oxml import parse_xml
from pptx.parts.presentation import PresentationPart

from ..unitutil.cxml import element
from ..unitutil.file import absjoin, snippet_bytes, testfile_bytes, test_file_dir
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


class Describe_RelatableMixin(object):
    """Unit-test suite for `pptx.opc.package._RelatableMixin`.

    This mixin is used for both OpcPackage and Part because both a package and a part
    can have relationships to target parts.
    """

    def it_can_find_a_part_related_by_reltype(self, _rels_prop_, relationships_, part_):
        relationships_.part_with_reltype.return_value = part_
        _rels_prop_.return_value = relationships_
        mixin = _RelatableMixin()

        related_part = mixin.part_related_by(RT.CHART)

        relationships_.part_with_reltype.assert_called_once_with(RT.CHART)
        assert related_part is part_

    def it_can_establish_a_relationship_to_another_part(
        self, _rels_prop_, relationships_, part_
    ):
        relationships_.get_or_add.return_value = "rId42"
        _rels_prop_.return_value = relationships_
        mixin = _RelatableMixin()

        rId = mixin.relate_to(part_, RT.SLIDE)

        relationships_.get_or_add.assert_called_once_with(RT.SLIDE, part_)
        assert rId == "rId42"

    def and_it_can_establish_a_relationship_to_an_external_link(
        self, request, _rels_prop_, relationships_
    ):
        relationships_.get_or_add_ext_rel.return_value = "rId24"
        _rels_prop_.return_value = relationships_
        mixin = _RelatableMixin()

        rId = mixin.relate_to("http://url", RT.HYPERLINK, is_external=True)

        relationships_.get_or_add_ext_rel.assert_called_once_with(
            RT.HYPERLINK, "http://url"
        )
        assert rId == "rId24"

    def it_can_find_a_related_part_by_rId(
        self, request, _rels_prop_, relationships_, relationship_, part_
    ):
        _rels_prop_.return_value = relationships_
        relationships_.__getitem__.return_value = relationship_
        relationship_.target_part = part_
        mixin = _RelatableMixin()

        related_part = mixin.related_part("rId17")

        relationships_.__getitem__.assert_called_once_with("rId17")
        assert related_part is part_

    def it_can_find_a_target_ref_URI_by_rId(
        self, request, _rels_prop_, relationships_, relationship_
    ):
        _rels_prop_.return_value = relationships_
        relationships_.__getitem__.return_value = relationship_
        relationship_.target_ref = "http://url"
        mixin = _RelatableMixin()

        target_ref = mixin.target_ref("rId9")

        relationships_.__getitem__.assert_called_once_with("rId9")
        assert target_ref == "http://url"

    # fixture components -----------------------------------

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, Part)

    @pytest.fixture
    def relationship_(self, request):
        return instance_mock(request, _Relationship)

    @pytest.fixture
    def relationships_(self, request):
        return instance_mock(request, _Relationships)

    @pytest.fixture
    def _rels_prop_(self, request):
        return property_mock(request, _RelatableMixin, "_rels")


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

    def it_can_drop_a_relationship(self, _rels_prop_, relationships_):
        _rels_prop_.return_value = relationships_

        OpcPackage(None).drop_rel("rId42")

        relationships_.pop.assert_called_once_with("rId42")

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
        all_rels = tuple(
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
        _rels_prop_.return_value = {r.rId: r for r in all_rels[:3]}
        part_0_.rels = {r.rId: r for r in all_rels[3:4]}
        part_1_.rels = {r.rId: r for r in all_rels[4:]}
        package = OpcPackage(None)

        rels = set(package.iter_rels())

        # -- sequence is not guaranteed, but count (len) and uniqueness are --
        assert rels == set(all_rels)

    def it_provides_access_to_the_main_document_part(self, request):
        presentation_part_ = instance_mock(request, PresentationPart)
        part_related_by_ = method_mock(
            request, OpcPackage, "part_related_by", return_value=presentation_part_
        )
        package = OpcPackage(None)

        presentation_part = package.main_document_part

        part_related_by_.assert_called_once_with(package, RT.OFFICE_DOCUMENT)
        assert presentation_part is presentation_part_

    @pytest.mark.parametrize(
        "ns, expected_n", (((), 1), ((1,), 2), ((1, 2), 3), ((2, 4), 3), ((1, 4), 3))
    )
    def it_can_find_the_next_available_partname(self, request, ns, expected_n):
        tmpl = "/x%d.xml"
        method_mock(
            request,
            OpcPackage,
            "iter_parts",
            return_value=(instance_mock(request, Part, partname=tmpl % n) for n in ns),
        )
        next_partname = tmpl % expected_n
        PackURI_ = class_mock(
            request, "pptx.opc.package.PackURI", return_value=PackURI(next_partname)
        )
        package = OpcPackage(None)

        partname = package.next_partname(tmpl)

        PackURI_.assert_called_once_with(next_partname)
        assert partname == next_partname

    def it_can_save_to_a_pkg_file(self, request, _rels_prop_, relationships_):
        _rels_prop_.return_value = relationships_
        parts_ = tuple(instance_mock(request, Part) for _ in range(3))
        method_mock(request, OpcPackage, "iter_parts", return_value=iter(parts_))
        PackageWriter_ = class_mock(request, "pptx.opc.package.PackageWriter")
        package = OpcPackage(None)

        package.save("prs.pptx")

        PackageWriter_.write.assert_called_once_with("prs.pptx", relationships_, parts_)

    def it_loads_the_pkg_file_to_help(self, request, _rels_prop_, relationships_):
        _PackageLoader_ = class_mock(request, "pptx.opc.package._PackageLoader")
        _PackageLoader_.load.return_value = "pkg-rels-xml", {"partname": "part"}
        _rels_prop_.return_value = relationships_
        package = OpcPackage("prs.pptx")

        return_value = package._load()

        _PackageLoader_.load.assert_called_once_with("prs.pptx", package)
        relationships_.load_from_xml.assert_called_once_with(
            PACKAGE_URI, "pkg-rels-xml", {"partname": "part"}
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

    # fixture components -----------------------------------

    @pytest.fixture
    def relationships_(self, request):
        return instance_mock(request, _Relationships)

    @pytest.fixture
    def _rels_prop_(self, request):
        return property_mock(request, OpcPackage, "_rels")


class Describe_PackageLoader(object):
    """Unit-test suite for `pptx.opc.package._PackageLoader` objects."""

    def it_provides_a_load_interface_classmethod(self, request, package_):
        _init_ = initializer_mock(request, _PackageLoader)
        pkg_xml_rels_ = element("r:Relationships")
        _load_ = method_mock(
            request,
            _PackageLoader,
            "_load",
            return_value=(pkg_xml_rels_, {"partname": "part"}),
        )

        pkg_xml_rels, parts = _PackageLoader.load("prs.pptx", package_)

        _init_.assert_called_once_with(ANY, "prs.pptx", package_)
        _load_.assert_called_once_with(ANY)
        assert pkg_xml_rels is pkg_xml_rels_
        assert parts == {"partname": "part"}

    def it_loads_the_package_to_help(self, request, _xml_rels_prop_):
        parts_ = {
            "partname_%d" % n: instance_mock(request, Part, partname="partname_%d" % n)
            for n in range(1, 4)
        }
        property_mock(request, _PackageLoader, "_parts", return_value=parts_)
        rels_ = dict(
            itertools.chain(
                (("/", instance_mock(request, _Relationships)),),
                (
                    ("partname_%d" % n, instance_mock(request, _Relationships))
                    for n in range(1, 4)
                ),
            )
        )
        _xml_rels_prop_.return_value = rels_
        package_loader = _PackageLoader(None, None)

        pkg_xml_rels, parts = package_loader._load()

        for part_ in parts_.values():
            part_.load_rels_from_xml.assert_called_once_with(
                rels_[part_.partname], parts_
            )
        assert pkg_xml_rels is rels_["/"]
        assert parts is parts_

    def it_loads_the_xml_relationships_from_the_package_to_help(self, request):
        pkg_xml_rels = parse_xml(snippet_bytes("package-rels-xml"))
        prs_xml_rels = parse_xml(snippet_bytes("presentation-rels-xml"))
        slide_xml_rels = CT_Relationships.new()
        thumbnail_xml_rels = CT_Relationships.new()
        core_xml_rels = CT_Relationships.new()
        _xml_rels_for_ = method_mock(
            request,
            _PackageLoader,
            "_xml_rels_for",
            side_effect=iter(
                (
                    pkg_xml_rels,
                    prs_xml_rels,
                    slide_xml_rels,
                    thumbnail_xml_rels,
                    core_xml_rels,
                )
            ),
        )
        package_loader = _PackageLoader(None, None)

        xml_rels = package_loader._xml_rels

        # print(f"{_xml_rels_for_.call_args_list=}")
        assert _xml_rels_for_.call_args_list == [
            call(package_loader, "/"),
            call(package_loader, "/ppt/presentation.xml"),
            call(package_loader, "/ppt/slides/slide1.xml"),
            call(package_loader, "/docProps/thumbnail.jpeg"),
            call(package_loader, "/docProps/core.xml"),
        ]
        assert xml_rels == {
            "/": pkg_xml_rels,
            "/ppt/presentation.xml": prs_xml_rels,
            "/ppt/slides/slide1.xml": slide_xml_rels,
            "/docProps/thumbnail.jpeg": thumbnail_xml_rels,
            "/docProps/core.xml": core_xml_rels,
        }

    # fixture components -----------------------------------

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, OpcPackage)

    @pytest.fixture
    def _xml_rels_prop_(self, request):
        return property_mock(request, _PackageLoader, "_xml_rels")


class DescribePart(object):
    """Unit-test suite for `pptx.opc.package.Part` objects."""

    def it_can_be_constructed_by_PartFactory(self, request, package_):
        partname_ = instance_mock(request, PackURI)
        _init_ = initializer_mock(request, Part)

        part = Part.load(partname_, CT.PML_SLIDE, package_, b"blob")

        _init_.assert_called_once_with(part, partname_, CT.PML_SLIDE, package_, b"blob")
        assert isinstance(part, Part)

    def it_uses_the_load_blob_as_its_blob(self):
        assert Part(None, None, None, b"blob").blob == b"blob"

    def it_can_change_its_blob(self):
        part = Part(None, None, None, b"old-blob")
        part.blob = b"new-blob"
        assert part.blob == b"new-blob"

    def it_knows_its_content_type(self):
        assert Part(None, CT.PML_SLIDE, None).content_type == CT.PML_SLIDE

    @pytest.mark.parametrize("ref_count, calls", ((2, []), (1, [call("rId42")])))
    def it_can_drop_a_relationship(self, request, relationships_, ref_count, calls):
        _rel_ref_count_ = method_mock(
            request, Part, "_rel_ref_count", return_value=ref_count
        )
        property_mock(request, Part, "_rels", return_value=relationships_)
        part = Part(None, None, None)

        part.drop_rel("rId42")

        _rel_ref_count_.assert_called_once_with(part, "rId42")
        assert relationships_.pop.call_args_list == calls

    def it_knows_the_package_it_belongs_to(self, package_):
        assert Part(None, None, package_).package is package_

    def it_knows_its_partname(self):
        assert Part(PackURI("/part/name"), None, None).partname == PackURI("/part/name")

    def it_can_change_its_partname(self):
        part = Part(PackURI("/old/part/name"), None, None)
        part.partname = PackURI("/new/part/name")
        assert part.partname == PackURI("/new/part/name")

    def it_provides_access_to_its_relationships_for_traversal(
        self, request, relationships_
    ):
        property_mock(request, Part, "_rels", return_value=relationships_)
        assert Part(None, None, None).rels is relationships_

    def it_can_load_a_blob_from_a_file_path_to_help(self):
        path = absjoin(test_file_dir, "minimal.pptx")
        with open(path, "rb") as f:
            file_bytes = f.read()
        part = Part(None, None, None, None)

        assert part._blob_from_file(path) == file_bytes

    def it_can_load_a_blob_from_a_file_like_object_to_help(self):
        part = Part(None, None, None, None)
        assert part._blob_from_file(io.BytesIO(b"012345")) == b"012345"

    def it_constructs_its_relationships_object_to_help(self, request, relationships_):
        _Relationships_ = class_mock(
            request, "pptx.opc.package._Relationships", return_value=relationships_
        )
        part = Part(PackURI("/ppt/slides/slide1.xml"), None, None)

        rels = part._rels

        _Relationships_.assert_called_once_with("/ppt/slides")
        assert rels is relationships_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, OpcPackage)

    @pytest.fixture
    def relationships_(self, request):
        return instance_mock(request, _Relationships)


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

        part = XmlPart.load(partname, CT.PML_SLIDE, package_, b"blob")

        parse_xml_.assert_called_once_with(b"blob")
        _init_.assert_called_once_with(part, partname, CT.PML_SLIDE, package_, element_)
        assert isinstance(part, XmlPart)

    def it_can_serialize_to_xml(self, request):
        element_ = element("p:sld")
        serialize_part_xml_ = function_mock(
            request, "pptx.opc.package.serialize_part_xml"
        )
        xml_part = XmlPart(None, None, None, element_)

        blob = xml_part.blob

        serialize_part_xml_.assert_called_once_with(element_)
        assert blob is serialize_part_xml_.return_value

    def it_knows_it_is_the_part_for_its_child_objects(self):
        xml_part = XmlPart(None, None, None, None)
        assert xml_part.part is xml_part


class DescribePartFactory(object):
    """Unit-test suite for `pptx.opc.package.PartFactory` objects."""

    def it_constructs_custom_part_type_for_registered_content_types(
        self, request, package_, part_
    ):
        SlidePart_ = class_mock(request, "pptx.opc.package.XmlPart")
        SlidePart_.load.return_value = part_
        partname = PackURI("/ppt/slides/slide7.xml")
        PartFactory.part_type_for[CT.PML_SLIDE] = SlidePart_

        part = PartFactory(partname, CT.PML_SLIDE, package_, b"blob")

        SlidePart_.load.assert_called_once_with(
            partname, CT.PML_SLIDE, package_, b"blob"
        )
        assert part is part_

    def it_constructs_part_using_default_class_when_no_custom_registered(
        self, request, package_, part_
    ):
        Part_ = class_mock(request, "pptx.opc.package.Part")
        Part_.load.return_value = part_
        partname = PackURI("/bar/foo.xml")

        part = PartFactory(partname, CT.OFC_VML_DRAWING, package_, b"blob")

        Part_.load.assert_called_once_with(
            partname, CT.OFC_VML_DRAWING, package_, b"blob"
        )
        assert part is part_

    # fixtures components ----------------------------------

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, OpcPackage)

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, Part)


class Describe_ContentTypeMap(object):
    """Unit-test suite for `pptx.opc.package._ContentTypeMap` objects."""

    def it_can_construct_from_content_types_xml(self, request):
        _init_ = initializer_mock(request, _ContentTypeMap)
        content_types_xml = (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-ty'
            'pes">\n'
            '  <Default Extension="xml" ContentType="application/xml"/>\n'
            '  <Default Extension="PNG" ContentType="image/png"/>\n'
            '  <Override PartName="/ppt/presentation.xml" ContentType="application/v'
            'nd.openxmlformats-officedocument.presentationml.presentation.main+xml"/'
            ">\n"
            "</Types>\n"
        )

        ct_map = _ContentTypeMap.from_xml(content_types_xml)

        _init_.assert_called_once_with(
            ct_map,
            {"/ppt/presentation.xml": CT.PML_PRESENTATION_MAIN},
            {"png": CT.PNG, "xml": CT.XML},
        )

    @pytest.mark.parametrize(
        "partname, expected_value",
        (
            ("/docProps/core.xml", CT.OPC_CORE_PROPERTIES),
            ("/ppt/presentation.xml", CT.PML_PRESENTATION_MAIN),
            ("/PPT/Presentation.XML", CT.PML_PRESENTATION_MAIN),
            ("/ppt/viewprops.xml", CT.PML_VIEW_PROPS),
        ),
    )
    def it_matches_an_override_on_case_insensitive_partname(
        self, content_type_map, partname, expected_value
    ):
        assert content_type_map[PackURI(partname)] == expected_value

    @pytest.mark.parametrize(
        "partname, expected_value",
        (
            ("/foo/bar.xml", CT.XML),
            ("/FOO/BAR.Rels", CT.OPC_RELATIONSHIPS),
            ("/foo/bar.jpeg", CT.JPEG),
        ),
    )
    def it_falls_back_to_case_insensitive_extension_default_match(
        self, content_type_map, partname, expected_value
    ):
        assert content_type_map[PackURI(partname)] == expected_value

    def it_raises_KeyError_on_partname_not_found(self, content_type_map):
        with pytest.raises(KeyError) as e:
            content_type_map[PackURI("/!blat/rhumba.1x&")]
        assert str(e.value) == (
            "\"no content-type for partname '/!blat/rhumba.1x&' "
            'in [Content_Types].xml"'
        )

    def it_raises_TypeError_on_key_not_instance_of_PackURI(self, content_type_map):
        with pytest.raises(TypeError) as e:
            content_type_map["/part/name1.xml"]
        assert str(e.value) == "_ContentTypeMap key must be <type 'PackURI'>, got str"

    # fixtures ---------------------------------------------

    @pytest.fixture(scope="class")
    def content_type_map(self):
        return _ContentTypeMap.from_xml(
            testfile_bytes("expanded_pptx", "[Content_Types].xml")
        )


class Describe_Relationships(object):
    """Unit-test suite for `pptx.opc.package._Relationships` objects."""

    @pytest.mark.parametrize("rId, expected_value", (("rId1", True), ("rId2", False)))
    def it_knows_whether_it_contains_a_relationship_with_rId(
        self, _rels_prop_, rId, expected_value
    ):
        _rels_prop_.return_value = {"rId1": None}
        assert (rId in _Relationships(None)) is expected_value

    def it_has_dict_style_lookup_of_rel_by_rId(self, _rels_prop_, relationship_):
        _rels_prop_.return_value = {"rId17": relationship_}
        assert _Relationships(None)["rId17"] is relationship_

    def but_it_raises_KeyError_when_no_relationship_has_rId(self, _rels_prop_):
        _rels_prop_.return_value = {}
        with pytest.raises(KeyError) as e:
            _Relationships(None)["rId6"]
        assert str(e.value) == "\"no relationship with key 'rId6'\""

    def it_can_iterate_the_rIds_of_the_relationships_it_contains(
        self, request, _rels_prop_
    ):
        rels_ = set(instance_mock(request, _Relationship) for n in range(5))
        _rels_prop_.return_value = {"rId%d" % (i + 1): r for i, r in enumerate(rels_)}
        relationships = _Relationships(None)

        for rId in relationships:
            rels_.remove(relationships[rId])

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
            "rId11": instance_mock(
                request,
                _Relationship,
                rId="rId11",
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
            "foo7W": instance_mock(
                request,
                _Relationship,
                rId="foo7W",
                reltype=RT.IMAGE,
                target_ref="../media/image1.png",
                is_external=False,
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

    @pytest.mark.parametrize(
        "target_ref, is_external, expected_value",
        (
            ("http://url", True, "rId1"),
            ("part_1", False, "rId2"),
            ("http://foo", True, "rId3"),
            ("part_2", False, "rId4"),
            ("http://bar", True, None),
        ),
    )
    def it_can_get_a_matching_relationship_to_help(
        self, request, _rels_by_reltype_prop_, target_ref, is_external, expected_value
    ):
        part_1, part_2 = (instance_mock(request, Part) for _ in range(2))
        _rels_by_reltype_prop_.return_value = {
            RT.SLIDE: [
                instance_mock(
                    request,
                    _Relationship,
                    rId=rId,
                    target_part=target_part,
                    target_ref=ref,
                    is_external=external,
                )
                for rId, target_part, ref, external in (
                    ("rId1", None, "http://url", True),
                    ("rId2", part_1, "/ppt/foo.bar", False),
                    ("rId3", None, "http://foo", True),
                    ("rId4", part_2, "/ppt/bar.foo", False),
                )
            ]
        }
        target = (
            target_ref if is_external else part_1 if target_ref == "part_1" else part_2
        )
        relationships = _Relationships(None)

        matching = relationships._get_matching(RT.SLIDE, target, is_external)

        assert matching == expected_value

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
            relationship, "/ppt", "rId42", RT.SLIDE, RTM.INTERNAL, part_
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
