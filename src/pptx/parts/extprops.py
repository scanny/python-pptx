"""Extended-properties part, corresponds to ``/docProps/app.xml`` part in package.

This part contains application-level document properties such as the slide count, the
authoring application's name, the company, the manager, and the slide-titles list.

The part is modeled so that python-pptx can keep the `<Slides>` count in sync with the
actual number of slides at save-time (see GitHub issue #131 -- Gmail and certain other
downstream tools refuse to preview a presentation when the count is stale) and so that
the remaining curated app-metadata fields are readable and writable (issue #105).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

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
    Exposes read/write access to the most commonly-used string-valued properties
    (``application``, ``app_version``, ``company``, ``manager``, ``hyperlink_base``,
    ``presentation_format``, ``template``) and read access to the ``slide_count``, which
    is automatically refreshed on save.
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

    # -- string-valued app.xml fields (issue #105) -----------------------

    @property
    def application(self) -> str:
        """Name of the authoring application, or '' when the `<Application>` element is absent.

        For presentations created by Microsoft PowerPoint this is typically
        ``"Microsoft Office PowerPoint"`` or ``"Microsoft Macintosh PowerPoint"``.
        """
        return self._element.application_text

    @application.setter
    def application(self, value: str) -> None:
        self._element.application_text = value

    @property
    def app_version(self) -> str:
        """Version of the authoring application, or '' when the `<AppVersion>` element is absent.

        Format is an application-specific string (e.g. ``"16.0000"`` for PowerPoint 2016+);
        python-pptx does not parse or validate this value.
        """
        return self._element.app_version_text

    @app_version.setter
    def app_version(self, value: str) -> None:
        self._element.app_version_text = value

    @property
    def company(self) -> str:
        """Value of the `<Company>` element, or '' when absent."""
        return self._element.company_text

    @company.setter
    def company(self, value: str) -> None:
        self._element.company_text = value

    @property
    def hyperlink_base(self) -> str:
        """Value of the `<HyperlinkBase>` element, or '' when absent.

        When non-empty, serves as the base URL used to resolve relative hyperlink targets.
        """
        return self._element.hyperlink_base_text

    @hyperlink_base.setter
    def hyperlink_base(self, value: str) -> None:
        self._element.hyperlink_base_text = value

    @property
    def manager(self) -> str:
        """Value of the `<Manager>` element, or '' when absent."""
        return self._element.manager_text

    @manager.setter
    def manager(self, value: str) -> None:
        self._element.manager_text = value

    @property
    def presentation_format(self) -> str:
        """Value of the `<PresentationFormat>` element, or '' when absent.

        Typical values include ``"Widescreen"``, ``"On-screen Show (4:3)"``, and
        ``"On-screen Show (16:9)"``; python-pptx does not constrain or validate the string.
        """
        return self._element.presentation_format_text

    @presentation_format.setter
    def presentation_format(self, value: str) -> None:
        self._element.presentation_format_text = value

    @property
    def template(self) -> str:
        """Value of the `<Template>` element, or '' when absent."""
        return self._element.template_text

    @template.setter
    def template(self, value: str) -> None:
        self._element.template_text = value

    # -- `<Slides>` count (issue #131) -----------------------------------
    #
    # NOTE: the slide count is re-synced from ``sldIdLst`` at save-time
    # (see :meth:`blob` below), so a manually-assigned value is clobbered on
    # save. The setter is kept for low-level callers, but end users should
    # not rely on overriding the count — the authoritative source is the
    # presentation's slide list.

    @property
    def slide_count(self) -> int:
        """Value of the `<Slides>` element, or 0 if absent.

        Automatically refreshed at save-time to match the presentation's current slide
        list; writing to this attribute directly is not usually necessary.
        """
        return self._element.slide_count

    @slide_count.setter
    def slide_count(self, value: int) -> None:
        self._element.slide_count = value

    # -- `<DocSecurity>` flag (ECMA-376 §22.2.2.6) ------------------------

    @property
    def doc_security(self) -> int | None:
        """Value of the `<DocSecurity>` element, or |None| if the element is absent.

        The flag records coarse document security state per ECMA-376:
        0 = none, 1 = password-protected, 2 = read-only recommended,
        4 = read-only enforced, 8 = locked for annotation. python-pptx does
        not enforce the flag — it is a hint consumed by viewers and
        collaboration tooling. Assigning |None| removes the element.
        """
        return self._element.doc_security

    @doc_security.setter
    def doc_security(self, value: int | None) -> None:
        self._element.doc_security = value

    # -- serialization hook ---------------------------------------------

    @property
    def blob(self) -> bytes:
        """bytes XML serialization of this part.

        Before serializing, refresh the `<Slides>` count to match the current `sldIdLst` in the
        presentation part. This is the fix for issue #131 -- previously the count was whatever
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
