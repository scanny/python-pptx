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
