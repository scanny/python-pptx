# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.customprops` module."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.custprops import CT_CustomProperties
from pptx.parts.customprops import CustomProperties, CustomPropertiesPart


class DescribeCustomPropertiesPart(object):
    """Unit-test suite for `pptx.parts.customprops.CustomPropertiesPart`."""

    def it_can_construct_a_default_custom_properties_part(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]

        assert isinstance(part, CustomPropertiesPart)
        assert part.content_type is CT.OFC_CUSTOM_PROPERTIES
        assert part.partname == "/docProps/custom.xml"
        assert isinstance(part._element, CT_CustomProperties)
        assert len(part) == 0

    def it_supports_getitem_setitem_and_delitem(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]

        part["x"] = "one"
        part["y"] = 42
        assert part["x"] == "one"
        assert part["y"] == 42

        del part["x"]
        assert "x" not in part
        with pytest.raises(KeyError):
            part["x"]
        with pytest.raises(KeyError):
            del part["missing"]

    def it_supports_contains_for_non_strings(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part["x"] = "one"
        # -- non-str keys should not blow up; simply return False --
        assert 42 not in part
        assert None not in part  # type: ignore[operator]
        assert "x" in part

    def it_supports_len_and_iter_in_document_order(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part["c"] = 1
        part["a"] = 2
        part["b"] = 3
        assert len(part) == 3
        assert list(part) == ["c", "a", "b"]

    def it_supports_keys_values_items_and_get(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part["a"] = "x"
        part["b"] = 2

        assert part.keys() == ["a", "b"]
        assert part.values() == ["x", 2]
        assert part.items() == [("a", "x"), ("b", 2)]
        assert part.get("a") == "x"
        assert part.get("missing") is None
        assert part.get("missing", "default") == "default"


class DescribeCustomProperties(object):
    """Unit-test suite for `pptx.parts.customprops.CustomProperties` facade."""

    def it_acts_empty_when_no_part_is_related(self, request: pytest.FixtureRequest):
        package = _FakePackage(existing_part=None)
        cp = CustomProperties(package)  # type: ignore[arg-type]

        assert len(cp) == 0
        assert list(cp) == []
        assert "anything" not in cp
        assert cp.keys() == []
        assert cp.values() == []
        assert cp.items() == []
        assert cp.get("nope") is None
        assert cp.get("nope", "d") == "d"

        with pytest.raises(KeyError):
            cp["anything"]
        with pytest.raises(KeyError):
            del cp["anything"]

    def it_creates_the_part_lazily_on_write(self):
        package = _FakePackage(existing_part=None)
        cp = CustomProperties(package)  # type: ignore[arg-type]

        cp["author"] = "Alice"

        # -- a part was created and cached on the fake package --
        assert package.created_part is not None
        assert cp["author"] == "Alice"
        assert "author" in cp
        assert len(cp) == 1

    def it_reads_through_to_an_existing_part(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part["k"] = "v"
        package = _FakePackage(existing_part=part)
        cp = CustomProperties(package)  # type: ignore[arg-type]

        assert cp["k"] == "v"
        assert "k" in cp
        assert cp.keys() == ["k"]
        assert cp.items() == [("k", "v")]

    def it_can_round_trip_each_value_type(self):
        package = _FakePackage(existing_part=None)
        cp = CustomProperties(package)  # type: ignore[arg-type]

        cp["s"] = "str"
        cp["i"] = 7
        cp["f"] = 1.5
        cp["b"] = True
        cp["d"] = dt.datetime(2030, 6, 15, 12, 0, 0)

        assert cp["s"] == "str"
        assert cp["i"] == 7
        assert cp["f"] == 1.5
        assert cp["b"] is True
        assert cp["d"] == dt.datetime(2030, 6, 15, 12, 0, 0)

    def it_supports_update_clear_pop_and_setdefault(self):
        package = _FakePackage(existing_part=None)
        cp = CustomProperties(package)  # type: ignore[arg-type]

        cp.update({"a": 1, "b": "two"})
        assert cp.items() == [("a", 1), ("b", "two")]

        assert cp.pop("a") == 1
        assert "a" not in cp
        assert cp.pop("missing", "default") == "default"
        with pytest.raises(KeyError):
            cp.pop("missing")

        assert cp.setdefault("b", "ignored") == "two"
        assert cp.setdefault("c", "new") == "new"
        assert cp["c"] == "new"

        cp.clear()
        assert len(cp) == 0

    def it_does_not_touch_the_package_when_clear_has_no_part(self):
        package = _FakePackage(existing_part=None)
        cp = CustomProperties(package)  # type: ignore[arg-type]

        cp.clear()

        assert package.created_part is None  # no lazy creation

    def it_does_not_materialize_part_on_read_only_operations(self):
        package = _FakePackage(existing_part=None)
        cp = CustomProperties(package)  # type: ignore[arg-type]

        _ = len(cp)
        _ = list(cp)
        _ = cp.get("x")
        _ = cp.keys()
        _ = cp.items()
        _ = cp.values()

        assert package.created_part is None


# -- test doubles ---------------------------------------------------------


class _FakePackage:
    """Minimal stand-in for `pptx.package.Package`.

    Supports the two entry points the CustomProperties facade uses:
    ``part_related_by(RT.CUSTOM_PROPERTIES)`` and ``custom_properties_part`` (the lazy
    creator). When `existing_part` is not None, reads find it; writes call
    ``custom_properties_part`` which auto-creates a default part if one doesn't already
    exist.
    """

    def __init__(self, existing_part: CustomPropertiesPart | None):
        self._part = existing_part
        # -- tracks a part created by the lazy creator (distinct from a part that was
        # -- provided up-front so tests can assert on "did the facade create one?").
        self.created_part: CustomPropertiesPart | None = None

    def part_related_by(self, reltype: str):
        assert reltype == RT.CUSTOM_PROPERTIES
        if self._part is None:
            raise KeyError(reltype)
        return self._part

    @property
    def custom_properties_part(self) -> CustomPropertiesPart:
        if self._part is None:
            self._part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
            self.created_part = self._part
        return self._part
