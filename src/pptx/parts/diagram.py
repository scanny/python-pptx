"""Diagram (SmartArt) related parts."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.opc.package import Part, XmlPart
from pptx.oxml import parse_xml

if TYPE_CHECKING:
    from pptx.opc.package import Package
    from pptx.opc.packuri import PackURI
    from pptx.oxml.diagram import CT_DiagramData


class DiagramDataPart(XmlPart):
    """Diagram data part - contains the SmartArt data model."""

    @property
    def data_model(self) -> CT_DiagramData:
        """The CT_DiagramData root element."""
        return cast("CT_DiagramData", self._element)

    @classmethod
    def load(
        cls, partname: PackURI, content_type: str, package: Package, blob: bytes
    ) -> DiagramDataPart:
        """Called by PartFactory to load diagram data part."""
        return cls(partname, content_type, package, element=parse_xml(blob))


class DiagramLayoutPart(Part):
    """Diagram layout part - contains layout definition."""

    pass


class DiagramColorsPart(Part):
    """Diagram colors part - contains color scheme."""

    pass


class DiagramStylePart(Part):
    """Diagram style part - contains quick style."""

    pass


class DiagramDrawingPart(XmlPart):
    """Diagram drawing part - contains pre-rendered drawing (MS extension)."""

    def update_all_text_elements(self, text_list: list[str]) -> None:
        """Update all text elements in drawing order with new texts."""
        from pptx.oxml.ns import qn

        # Find all <a:t> elements in the drawing
        text_elements = self._element.findall(".//" + qn("a:t"))

        # Update each text element with corresponding text from list
        for idx, text_el in enumerate(text_elements):
            if idx < len(text_list):
                text_el.text = text_list[idx]

    @classmethod
    def load(
        cls, partname: PackURI, content_type: str, package: Package, blob: bytes
    ) -> DiagramDrawingPart:
        """Called by PartFactory to load diagram drawing part."""
        return cls(partname, content_type, package, element=parse_xml(blob))
