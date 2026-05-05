# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.extprops` module."""

from __future__ import annotations

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.extprops import CT_ExtendedProperties
from pptx.parts.extprops import ExtendedPropertiesPart


class DescribeExtendedPropertiesPart(object):
    """Unit-test suite for `pptx.parts.extprops.ExtendedPropertiesPart` objects."""

    def it_can_construct_a_default_extended_properties_part(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]

        assert isinstance(ext_props, ExtendedPropertiesPart)
        assert ext_props.content_type is CT.OFC_EXTENDED_PROPERTIES
        assert ext_props.partname == "/docProps/app.xml"
        assert isinstance(ext_props._element, CT_ExtendedProperties)
        assert ext_props.slide_count == 0

    def it_exposes_the_slide_count_read_write(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]

        ext_props.slide_count = 4

        assert ext_props.slide_count == 4

    def it_updates_the_slide_count_at_save_time(self, request: pytest.FixtureRequest):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]
        # -- set the in-memory count to a stale value --
        ext_props.slide_count = 0
        # -- inject a fake package whose presentation part reports 5 sldId children --
        ext_props._package = _FakePackage(n_slides=5)  # type: ignore[assignment]

        blob = ext_props.blob

        assert b"<Slides>5</Slides>" in blob
        assert ext_props.slide_count == 5

    def it_leaves_the_slide_count_alone_when_presentation_is_unreachable(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]
        ext_props.slide_count = 2

        # -- package attribute is None / bogus, so part_related_by will raise AttributeError --
        ext_props._package = None  # type: ignore[assignment]

        # -- calling blob should not raise; the existing count is preserved --
        blob = ext_props.blob

        assert b"<Slides>2</Slides>" in blob
        assert ext_props.slide_count == 2

    # -- string-valued public properties (issue #105) --------------------

    @pytest.mark.parametrize(
        "prop_name",
        [
            "application",
            "app_version",
            "company",
            "hyperlink_base",
            "manager",
            "presentation_format",
            "template",
        ],
    )
    def it_returns_empty_string_for_missing_string_fields(self, prop_name: str):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]

        assert getattr(ext_props, prop_name) == ""

    @pytest.mark.parametrize(
        ("prop_name", "value"),
        [
            ("application", "Microsoft Office PowerPoint"),
            ("app_version", "16.0000"),
            ("company", "Acme Corp"),
            ("hyperlink_base", "https://example.com/docs/"),
            ("manager", "Alice Jones"),
            ("presentation_format", "Widescreen"),
            ("template", "Default Theme.potx"),
        ],
    )
    def it_round_trips_string_fields(self, prop_name: str, value: str):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]

        setattr(ext_props, prop_name, value)

        assert getattr(ext_props, prop_name) == value

    def it_persists_string_fields_in_the_serialized_blob(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]
        ext_props.company = "Acme Corp"
        ext_props.manager = "Alice Jones"
        # -- avoid the save-time slide-count refresh by making the package unreachable --
        ext_props._package = None  # type: ignore[assignment]

        blob = ext_props.blob

        assert b"<Company>Acme Corp</Company>" in blob
        assert b"<Manager>Alice Jones</Manager>" in blob

    def it_overwrites_existing_string_field_values(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]
        ext_props.company = "Old Corp"

        ext_props.company = "New Corp"

        assert ext_props.company == "New Corp"

    # -- `<DocSecurity>` (ECMA-376 §22.2.2.6) ----------------------------

    def it_returns_None_when_doc_security_is_absent(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]

        assert ext_props.doc_security is None

    def it_round_trips_doc_security_through_the_blob(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]
        # -- avoid the save-time slide-count refresh; it's orthogonal here --
        ext_props._package = None  # type: ignore[assignment]

        ext_props.doc_security = 2

        assert ext_props.doc_security == 2
        assert b"<DocSecurity>2</DocSecurity>" in ext_props.blob

    def it_removes_doc_security_on_assignment_to_None(self):
        ext_props = ExtendedPropertiesPart.default(None)  # type: ignore[arg-type]
        ext_props._package = None  # type: ignore[assignment]
        ext_props.doc_security = 4
        assert ext_props.doc_security == 4

        ext_props.doc_security = None

        assert ext_props.doc_security is None
        assert b"DocSecurity" not in ext_props.blob


class _FakePackage(object):
    """Minimal test double that stands in for `pptx.package.Package`."""

    def __init__(self, n_slides: int):
        self._n_slides = n_slides

    def part_related_by(self, reltype: str):
        assert reltype == RT.OFFICE_DOCUMENT
        return _FakePresentationPart(self._n_slides)


class _FakePresentationPart(object):
    def __init__(self, n_slides: int):
        self._element = _FakePresentationElement(n_slides)


class _FakePresentationElement(object):
    def __init__(self, n_slides: int):
        # -- a list is the simplest thing that supports len() --
        self.sldIdLst = [object() for _ in range(n_slides)]
