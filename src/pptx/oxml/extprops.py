"""lxml custom element classes for extended-properties-related XML elements.

The extended-properties part corresponds to ``/docProps/app.xml`` in the OPC package. Its
``<Properties>`` root element contains an assortment of document-level metadata such as the total
slide count, the slide-titles list, the authoring application, etc.

python-pptx does not presently provide a user-facing API for the majority of these properties;
the primary reason this element is modeled at all is so that the ``<Slides>`` count can be kept
in sync with the actual number of slides in the presentation at save-time. Several downstream
consumers (notably the Gmail attachment-preview feature, but also ``pptx2html`` and assorted
thumbnail generators) read ``<Slides>`` directly and fail to render any preview when the number
does not match reality.
"""

from __future__ import annotations

from typing import cast

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
            # -- are lenient for app.xml in practice; inserting as the last child is what Office
            # -- authoring tools do when the element is missing.
            from lxml.etree import SubElement

            slides = SubElement(self, qn("ep:Slides"))
        slides.text = str(value)
