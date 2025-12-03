"""lxml custom element classes for diagram (SmartArt) elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, ZeroOrOne
from pptx.oxml.simpletypes import XsdString

if TYPE_CHECKING:
    pass


class CT_DiagramData(BaseOxmlElement):
    """<dgm:dataModel> element - root element of diagram data."""

    ptLst: CT_DiagramPointList = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "dgm:ptLst"
    )
    cxnLst: CT_DiagramConnectionList = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "dgm:cxnLst"
    )


class CT_DiagramPointList(BaseOxmlElement):
    """<dgm:ptLst> element - list of diagram points (nodes)."""

    def iter_pts(self) -> Iterator[CT_DiagramPoint]:
        """Generate each <dgm:pt> child element."""
        for pt in self.findall(qn("dgm:pt")):
            yield pt


class CT_DiagramPoint(BaseOxmlElement):
    """<dgm:pt> element - a node in the diagram."""

    modelId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "modelId", XsdString
    )
    type: str | None = OptionalAttribute("type", XsdString)  # pyright: ignore[reportAssignmentType]
    t: CT_DiagramText | None = ZeroOrOne("dgm:t")  # pyright: ignore[reportAssignmentType]

    @property
    def text(self) -> str:
        """Extract all text content from this point."""
        if self.t is None:
            return ""
        # Extract text from all <a:t> elements
        text_elements = self.t.findall(".//" + qn("a:t"))
        return "".join(el.text for el in text_elements if el.text)

    @text.setter
    def text(self, value: str) -> None:
        """Set text content for this point."""
        if self.t is None:
            return
        # Find all <a:t> elements and set the text to the first one, clear others
        text_elements = self.t.findall(".//" + qn("a:t"))
        if text_elements:
            # Set text in first element, clear the rest
            text_elements[0].text = value
            for el in text_elements[1:]:
                el.text = ""


class CT_DiagramText(BaseOxmlElement):
    """<dgm:t> element - text container in a diagram point."""

    pass


class CT_DiagramConnectionList(BaseOxmlElement):
    """<dgm:cxnLst> element - list of connections between points."""

    def iter_cxns(self) -> Iterator[CT_DiagramConnection]:
        """Generate each <dgm:cxn> child element."""
        for cxn in self.findall(qn("dgm:cxn")):
            yield cxn


class CT_DiagramConnection(BaseOxmlElement):
    """<dgm:cxn> element - connection/relationship between diagram points."""

    modelId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "modelId", XsdString
    )
    srcId: str | None = OptionalAttribute("srcId", XsdString)  # pyright: ignore[reportAssignmentType]
    destId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "destId", XsdString
    )
    type: str | None = OptionalAttribute("type", XsdString)  # pyright: ignore[reportAssignmentType]
