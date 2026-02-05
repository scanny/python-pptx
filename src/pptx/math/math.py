"""High-level Math class for OMML equation management."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.math import CT_OMath

if TYPE_CHECKING:
    from pptx.oxml.shapes import ShapeElement


class Math:
    """High-level interface for OMML math equations."""

    def __init__(self, shape_element: ShapeElement, parent):
        """Initialize Math with shape element and parent."""
        self._shape_element = shape_element
        self._parent = parent
        self._omml_element = None

    def _create_rpr(self):
        """Create standard rPr formatting properties for PowerPoint 16.105.2."""
        rpr = OxmlElement("a:rPr")
        rpr.set("lang", "en-US")
        rpr.set("i", "1")
        rpr.set("smtClean", "0")

        latin = OxmlElement("a:latin")
        latin.set("typeface", "Cambria Math")
        latin.set("panose", "02040503050406030204")
        latin.set("pitchFamily", "18")
        latin.set("charset", "0")
        rpr.append(latin)

        return rpr

    def _ensure_formatting(self, omath_element):
        """Ensure all text runs have proper PowerPoint formatting."""
        # Add <a:rPr> formatting to all <m:r> elements that lack it
        for r_element in omath_element.xpath(".//*[local-name() = 'r']"):
            if not r_element.xpath(".//*[local-name() = 'rPr']"):
                rpr = self._create_rpr()

                # Insert rPr as first child of m:r
                if len(r_element) > 0:
                    r_element.insert(0, rpr)
                else:
                    r_element.append(rpr)

    def _add_to_slide_container(self, omath_element):
        """Add OMML element to the shape itself, not to separate container shapes."""
        # Store OMML element in this math shape
        self._omml_element = omath_element

        # Apply automatic formatting
        self._ensure_formatting(omath_element)

    @property
    def omml_element(self) -> CT_OMath:
        """Get or create the underlying OMML element."""
        if self._omml_element is None:
            # Create OMML element
            self._omml_element = OxmlElement("m:oMath")
        return self._omml_element

    def get_omml(self) -> str:
        """Get OMML XML string."""
        from pptx.oxml.xmlchemy import serialize_for_reading
        return serialize_for_reading(self.omml_element)
        
    def set_omml(self, omml_xml: str):
        """Replace current OMML with new XML string."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import qn

        # Parse the XML and extract the oMath element
        new_element = parse_xml(omml_xml)

        # Find the actual oMath element (it might be nested)
        omath_element = new_element
        expected_tags = [
            qn("m:oMath"),  # Standard namespace
            "{http://purl.oclc.org/ooxml/officeDocument/math}oMath",  # Alternative namespace
        ]

        if new_element.tag not in expected_tags:
            # Look for oMath child with proper namespace
            omath_elements = new_element.xpath(".//*[local-name() = 'oMath']")
            if omath_elements:
                omath_element = omath_elements[0]
            else:
                raise ValueError("Root element must be m:oMath")

        # Replace the OMML element in the shape
        self._omml_element = omath_element

    def add_omml(self, omml_xml: str):
        """Add OMML content to the equation with automatic formatting."""
        from pptx.oxml import parse_xml

        # Parse the XML and extract the oMath element
        new_element = parse_xml(omml_xml)

        # Find the actual oMath element (it might be nested)
        omath_element = new_element
        expected_tags = [
            "m:oMath",  # Standard namespace
            "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath",  # Test namespace
            "{http://purl.oclc.org/ooxml/officeDocument/math}oMath",  # Alternative namespace
        ]

        if new_element.tag not in expected_tags:
            # Look for oMath child with proper namespace
            omath_elements = new_element.xpath(".//*[local-name() = 'oMath']")
            if omath_elements:
                omath_element = omath_elements[0]
            else:
                raise ValueError("Root element must be m:oMath")

        # Add to slide-level container with automatic formatting
        self._add_to_slide_container(omath_element)

        # Store reference for get_omml
        self._omml_element = omath_element
