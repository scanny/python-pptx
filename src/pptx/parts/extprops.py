"""Extended-properties part, corresponds to ``/docProps/app.xml`` part in package.

This part contains application-level document properties such as the slide count, the
authoring application's name, and the slide-titles list. python-pptx does not surface most of
these as a user API; the part exists so that python-pptx can keep the `<Slides>` count in sync
with the actual number of slides at save-time. See GitHub issue #131 — Gmail (and certain other
downstream tools) refuse to preview a presentation when the count is stale.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.extprops import CT_ExtendedProperties

if TYPE_CHECKING:
    from pptx.package import Package


class ExtendedPropertiesPart(XmlPart):
    """Corresponds to part named `/docProps/app.xml`.

    Holds the extended (application-level) document properties for this document package.
    """

    _element: CT_ExtendedProperties

    @classmethod
    def default(cls, package: Package) -> ExtendedPropertiesPart:
        """Return a default new |ExtendedPropertiesPart| instance.

        Suitable as a starting point for a package that doesn't yet have one. Contains a single
        `<Slides>0</Slides>` child element; the count is updated at save-time to reflect the
        current number of slides.
        """
        return cls(
            PackURI("/docProps/app.xml"),
            CT.OFC_EXTENDED_PROPERTIES,
            package,
            CT_ExtendedProperties.new_extendedProperties(),
        )

    @property
    def slide_count(self) -> int:
        """Value of the `<Slides>` element, or 0 if absent."""
        return self._element.slide_count

    @slide_count.setter
    def slide_count(self, value: int) -> None:
        self._element.slide_count = value

    @property
    def blob(self) -> bytes:  # pyright: ignore[reportIncompatibleMethodOverride]
        """bytes XML serialization of this part.

        Before serializing, refresh the `<Slides>` count to match the current `sldIdLst` in the
        presentation part. This is the fix for issue #131 — previously the count was whatever
        value was present when the package was opened (or zero for a newly-created package).
        """
        self._sync_slide_count()
        return super().blob

    def _sync_slide_count(self) -> None:
        """Write the current slide count from the presentation's `sldIdLst` into `<Slides>`.

        If the presentation part cannot be located (e.g. the package is not fully wired up, as
        in some low-level unit tests), the existing count is left unchanged.
        """
        try:
            presentation_part = self.package.part_related_by(RT.OFFICE_DOCUMENT)
        except (AttributeError, KeyError):
            return
        presentation_elm = getattr(presentation_part, "_element", None)
        sldIdLst = None if presentation_elm is None else presentation_elm.sldIdLst
        count = 0 if sldIdLst is None else len(sldIdLst)
        if self._element.slide_count != count:
            self._element.slide_count = count
