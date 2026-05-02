"""lxml custom element classes for DrawingML line-related XML elements."""

from __future__ import annotations

from pptx.enum.dml import (
    MSO_LINE_DASH_STYLE,
    MSO_LINE_END_LENGTH,
    MSO_LINE_END_TYPE,
    MSO_LINE_END_WIDTH,
)
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute


class CT_PresetLineDashProperties(BaseOxmlElement):
    """`a:prstDash` custom element class"""

    val = OptionalAttribute("val", MSO_LINE_DASH_STYLE)


class CT_LineEndProperties(BaseOxmlElement):
    """`a:headEnd` or `a:tailEnd` custom element class.

    Holds the arrowhead decoration at one end of a line.
    """

    type: MSO_LINE_END_TYPE | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", MSO_LINE_END_TYPE
    )
    w: MSO_LINE_END_WIDTH | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w", MSO_LINE_END_WIDTH
    )
    len: MSO_LINE_END_LENGTH | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "len", MSO_LINE_END_LENGTH
    )
