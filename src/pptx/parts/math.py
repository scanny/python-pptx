"""Math parts for OMML content."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.oxml.math import CT_OMath

if TYPE_CHECKING:
    from pptx.math.math import Math


class MathPart(XmlPart):
    """Part containing OMML (Office Math Markup Language) content."""
    
    content_type: str = CT.OFFICE_MATH
    
    def __init__(self, package, partname, content_type=None, element=None, blob=None):
        super().__init__(package, partname, content_type, element, blob)
        self._math = None
    
    @property
    def math(self) -> Math:
        """Math object providing access to OMML content."""
        if self._math is None:
            from pptx.math.math import Math
            self._math = Math(self._element)
        return self._math
    
    def get_or_add_omml_element(self) -> CT_OMath:
        """Get or create the OMML element."""
        if not isinstance(self._element, CT_OMath):
            from pptx.oxml.xmlchemy import OxmlElement
            self._element = OxmlElement("m:oMath")
        return self._element
    
    def parse_omml_string(self, omml_xml: str) -> CT_OMath:
        """Parse OMML XML string into CT_OMath element."""
        from pptx.oxml import parse_xml
        element = parse_xml(omml_xml)
        if not isinstance(element, CT_OMath):
            raise ValueError("Root element must be m:oMath")
        return element
