"""Custom-properties part, corresponds to ``/docProps/custom.xml`` in the package.

This part is distinct from both the core-properties part (Dublin Core metadata in
``/docProps/core.xml``) and the extended-properties part (application-level metadata in
``/docProps/app.xml``). It holds user-defined "Custom" document properties — the ones
PowerPoint shows under *File → Info → Properties → Advanced → Custom*. These are the
properties a Word- or PowerPoint-level field code such as ``{ DOCPROPERTY MyProp }``
resolves against, which is what makes this the right wiring point for issue #259.

See :class:`pptx.oxml.custprops.CT_CustomProperties` for the XML-level model; this
module provides the part wrapper (OPC plumbing) and a ``dict``-like :class:`.CustomProperties`
facade returned by :attr:`pptx.presentation.Presentation.custom_properties`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterator

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.custprops import CT_CustomProperties, CustomPropValue

if TYPE_CHECKING:
    from pptx.package import Package


# -- sentinel for pop()'s optional `default` argument. None is a plausible value for
# -- some custom-property retrieval scenarios so we use a dedicated private object to
# -- mean "raise KeyError if missing and no default was provided".
_MISSING: Any = object()


class CustomPropertiesPart(XmlPart):
    """Corresponds to part named `/docProps/custom.xml`.

    Holds user-defined custom document properties for this package. The part is created
    lazily on first access via :attr:`Package.custom_properties_part`.
    """

    _element: CT_CustomProperties

    @classmethod
    def default(cls, package: Package) -> CustomPropertiesPart:
        """Return a default new |CustomPropertiesPart| instance.

        Suitable as a starting point for a package that doesn't yet have one. The returned
        part contains a bare ``<Properties/>`` root with no ``<property>`` children.
        """
        return cls(
            PackURI("/docProps/custom.xml"),
            CT.OFC_CUSTOM_PROPERTIES,
            package,
            CT_CustomProperties.new_customProperties(),
        )

    # -- dict-like proxy over the underlying XML ---------------------------

    def __len__(self) -> int:
        return len(self._element)

    def __iter__(self) -> Iterator[str]:
        return iter(self._element)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        return self._element.has_name(name)

    def __getitem__(self, name: str) -> CustomPropValue:
        value = self._element.get_value(name)
        if value is None:
            raise KeyError(name)
        return value

    def __setitem__(self, name: str, value: CustomPropValue) -> None:
        self._element.set_value(name, value)

    def __delitem__(self, name: str) -> None:
        if not self._element.delete_value(name):
            raise KeyError(name)

    def get(
        self, name: str, default: CustomPropValue | None = None
    ) -> CustomPropValue | None:
        """Return the value for `name` if present, else `default`."""
        value = self._element.get_value(name)
        return default if value is None else value

    def keys(self) -> list[str]:
        """Return the list of property names in document order."""
        return self._element.names()

    def values(self) -> list[CustomPropValue]:
        """Return the list of property values in document order."""
        return [
            v
            for v in (self._element.get_value(n) for n in self._element.names())
            if v is not None
        ]

    def items(self) -> list[tuple[str, CustomPropValue]]:
        """Return the list of `(name, value)` pairs in document order."""
        pairs: list[tuple[str, CustomPropValue]] = []
        for name in self._element.names():
            value = self._element.get_value(name)
            if value is not None:
                pairs.append((name, value))
        return pairs


class CustomProperties:
    """Dict-like facade over a :class:`CustomPropertiesPart`.

    Surfaced via :attr:`pptx.presentation.Presentation.custom_properties`. Keys are the
    user-visible property names (strings); values are the python-level representations of
    the typed-vt children: ``str`` for ``vt:lpwstr``, ``int`` for ``vt:i4``, ``float`` for
    ``vt:r8``, ``bool`` for ``vt:bool``, and ``datetime.datetime`` for ``vt:filetime``.

    The facade delegates get/set/delete and iteration to :class:`CustomPropertiesPart` but
    keeps part-creation lazy: a default-empty part is materialized the first time a write
    operation needs it. Reads never create the part, so a presentation that never touches
    custom properties round-trips without gaining a stray ``/docProps/custom.xml`` entry.
    """

    def __init__(self, package: Package):
        super().__init__()
        self._package = package

    # -- membership / lookup ---------------------------------------------

    def __len__(self) -> int:
        part = self._existing_part()
        return 0 if part is None else len(part)

    def __iter__(self) -> Iterator[str]:
        part = self._existing_part()
        if part is None:
            return iter(())
        return iter(part)

    def __contains__(self, name: object) -> bool:
        part = self._existing_part()
        if part is None:
            return False
        return name in part

    def __getitem__(self, name: str) -> CustomPropValue:
        part = self._existing_part()
        if part is None:
            raise KeyError(name)
        return part[name]

    def __setitem__(self, name: str, value: CustomPropValue) -> None:
        self._get_or_add_part()[name] = value

    def __delitem__(self, name: str) -> None:
        part = self._existing_part()
        if part is None:
            raise KeyError(name)
        del part[name]

    # -- dict-style convenience ------------------------------------------

    def get(
        self, name: str, default: CustomPropValue | None = None
    ) -> CustomPropValue | None:
        part = self._existing_part()
        if part is None:
            return default
        return part.get(name, default)

    def keys(self) -> list[str]:
        part = self._existing_part()
        return [] if part is None else part.keys()

    def values(self) -> list[CustomPropValue]:
        part = self._existing_part()
        return [] if part is None else part.values()

    def items(self) -> list[tuple[str, CustomPropValue]]:
        part = self._existing_part()
        return [] if part is None else part.items()

    def update(self, other: dict[str, CustomPropValue]) -> None:
        """Bulk-add / overwrite properties from the mapping `other`."""
        if not other:
            return
        part = self._get_or_add_part()
        for name, value in other.items():
            part[name] = value

    def clear(self) -> None:
        """Remove every custom property (but not the part itself)."""
        part = self._existing_part()
        if part is None:
            return
        for name in list(part.keys()):
            del part[name]

    def pop(self, name: str, default: Any = _MISSING) -> CustomPropValue | None:
        """Remove `name` and return its value, or `default` if not present.

        Raises ``KeyError`` if `name` is absent and no `default` is supplied — matching
        ``dict.pop()`` semantics.
        """
        part = self._existing_part()
        if part is None or name not in part:
            if default is _MISSING:
                raise KeyError(name)
            return default
        value = part[name]
        del part[name]
        return value

    def setdefault(
        self, name: str, default: CustomPropValue
    ) -> CustomPropValue:
        """Return the current value for `name`, setting it to `default` if absent."""
        part = self._existing_part()
        if part is not None and name in part:
            return part[name]
        self._get_or_add_part()[name] = default
        return default

    # -- part-lookup helpers ---------------------------------------------

    def _existing_part(self) -> CustomPropertiesPart | None:
        """Return the package's custom-properties part, or None if one isn't related.

        Read-only operations use this to avoid forcing a default part into a package that
        legitimately has no custom properties (PowerPoint files without any).
        """
        try:
            return self._package.part_related_by(RT.CUSTOM_PROPERTIES)
        except KeyError:
            return None

    def _get_or_add_part(self) -> CustomPropertiesPart:
        """Return the package's custom-properties part, creating one if needed."""
        return self._package.custom_properties_part


# Re-export the value type alias so downstream callers can type-annotate against it
# without reaching into the oxml layer.
__all__ = ["CustomProperties", "CustomPropertiesPart", "CustomPropValue"]
