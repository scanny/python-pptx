"""lxml custom element classes for extended-properties-related XML elements.

The extended-properties part corresponds to ``/docProps/app.xml`` in the OPC package. Its
``<Properties>`` root element contains an assortment of document-level metadata such as the total
slide count, the slide-titles list, the authoring application, etc.

python-pptx surfaces a curated subset of these properties (see :class:`ExtendedPropertiesPart`);
the element is also modeled so the ``<Slides>`` count can be kept in sync with the actual number
of slides in the presentation at save-time. Several downstream consumers (notably the Gmail
attachment-preview feature, but also ``pptx2html`` and assorted thumbnail generators) read
``<Slides>`` directly and fail to render any preview when the number does not match reality.
"""

from __future__ import annotations

from typing import cast

from lxml.etree import SubElement

from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import BaseOxmlElement

# -- the `ep` namespace URI; children of the Properties element live in this namespace too,
# -- conventionally emitted without an explicit prefix (they inherit the default namespace).
_EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


class CT_ExtendedProperties(BaseOxmlElement):
    """`Properties` element, root of the extended-properties part (``/docProps/app.xml``)."""

    _Properties_tmpl = (
        '<Properties xmlns="%s" xmlns:vt="%s"/>\n' % (_EP_NS, _VT_NS)
    )

    @staticmethod
    def new_extendedProperties() -> CT_ExtendedProperties:
        """Return a newly-created, default `Properties` element.

        Contains a single `<Slides>0</Slides>` child and no other properties. This is sufficient
        for round-tripping; PowerPoint and other consumers tolerate missing optional fields.
        """
        properties = cast(CT_ExtendedProperties, parse_xml(CT_ExtendedProperties._Properties_tmpl))
        properties.slide_count = 0
        return properties

    # -- `<Slides>` (int) ------------------------------------------------

    @property
    def slide_count(self) -> int:
        """Integer value of the `<Slides>` child element, or 0 if the element is absent.

        A non-integer or negative value also resolves to 0, matching the lenient reader behavior
        used for other count-style elements in the spec.
        """
        slides = self.find(qn("ep:Slides"))
        if slides is None or slides.text is None:
            return 0
        try:
            value = int(slides.text)
        except ValueError:
            return 0
        return value if value >= 0 else 0

    @slide_count.setter
    def slide_count(self, value: int) -> None:
        """Set the `<Slides>` child element's text to the string value of `value`.

        Creates the element if it is not present. `value` must be a non-negative `int`.
        """
        if not isinstance(value, int) or value < 0:  # pyright: ignore[reportUnnecessaryIsInstance]
            raise ValueError(
                "slide_count must be a non-negative integer, got %r" % (value,)
            )
        slides = self.find(qn("ep:Slides"))
        if slides is None:
            # -- xmlchemy ZeroOrOne isn't used here because the schema's sequence ordering rules
            # -- are lenient for app.xml in practice (CT_Properties is defined with xsd:all);
            # -- inserting as the last child is what Office authoring tools do when the element
            # -- is missing.
            slides = SubElement(self, qn("ep:Slides"))
        slides.text = str(value)

    # -- string-valued fields --------------------------------------------

    @property
    def application_text(self) -> str:
        """Text of the `<Application>` child element, or '' when absent."""
        return self._text_of("Application")

    @application_text.setter
    def application_text(self, value: str) -> None:
        self._set_text("Application", value)

    @property
    def app_version_text(self) -> str:
        """Text of the `<AppVersion>` child element, or '' when absent."""
        return self._text_of("AppVersion")

    @app_version_text.setter
    def app_version_text(self, value: str) -> None:
        self._set_text("AppVersion", value)

    @property
    def company_text(self) -> str:
        """Text of the `<Company>` child element, or '' when absent."""
        return self._text_of("Company")

    @company_text.setter
    def company_text(self, value: str) -> None:
        self._set_text("Company", value)

    @property
    def hyperlink_base_text(self) -> str:
        """Text of the `<HyperlinkBase>` child element, or '' when absent."""
        return self._text_of("HyperlinkBase")

    @hyperlink_base_text.setter
    def hyperlink_base_text(self, value: str) -> None:
        self._set_text("HyperlinkBase", value)

    @property
    def manager_text(self) -> str:
        """Text of the `<Manager>` child element, or '' when absent."""
        return self._text_of("Manager")

    @manager_text.setter
    def manager_text(self, value: str) -> None:
        self._set_text("Manager", value)

    @property
    def presentation_format_text(self) -> str:
        """Text of the `<PresentationFormat>` child element, or '' when absent."""
        return self._text_of("PresentationFormat")

    @presentation_format_text.setter
    def presentation_format_text(self, value: str) -> None:
        self._set_text("PresentationFormat", value)

    @property
    def template_text(self) -> str:
        """Text of the `<Template>` child element, or '' when absent."""
        return self._text_of("Template")

    @template_text.setter
    def template_text(self, value: str) -> None:
        self._set_text("Template", value)

    # -- helpers ---------------------------------------------------------

    def _text_of(self, local_name: str) -> str:
        """Return the text of `<local_name>`, or '' when the child is absent or empty.

        `local_name` is the unqualified element name (e.g. ``"Company"``); it is resolved to
        the `ep:` namespace internally.
        """
        element = self.find(qn("ep:%s" % local_name))
        if element is None or element.text is None:
            return ""
        return element.text

    def _set_text(self, local_name: str, value: str) -> None:
        """Set the text of `<local_name>` to `value`.

        Creates the child element in the `ep:` namespace if it is not already present. `value`
        must be a `str`; assigning ``""`` leaves the element in place with empty text (matching
        the way ``coreprops`` handles empty-string writes -- both round-trip cleanly and neither
        represents "delete this field" to downstream consumers).
        """
        if not isinstance(value, str):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(
                "%s must be a str, got %s" % (local_name, type(value).__name__)
            )
        element = self.find(qn("ep:%s" % local_name))
        if element is None:
            element = SubElement(self, qn("ep:%s" % local_name))
        element.text = value
