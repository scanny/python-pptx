"""SmartArt shape and related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.shared import ParentedElementProxy

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DiagramPoint, CT_DiagramPointList
    from pptx.parts.diagram import DiagramDataPart


class SmartArt(ParentedElementProxy):
    """SmartArt diagram object contained in a GraphicFrame."""

    def __init__(self, data_part: DiagramDataPart, parent):
        super().__init__(data_part.data_model, parent)
        self._data_part = data_part

    @property
    def nodes(self) -> _SmartArtNodes:
        """Collection of nodes in this SmartArt diagram."""
        return _SmartArtNodes(self._element.ptLst, self, self._data_part)

    def iter_text(self) -> Iterator[str]:
        """Generate all text content from diagram nodes."""
        for node in self.nodes:
            text = node.text
            if text:
                yield text

    @property
    def text_content(self) -> list[str]:
        """List of all text strings from diagram nodes."""
        return list(self.iter_text())

    def _sync_drawing_part(self) -> None:
        """Sync drawing part with current data model texts."""
        # Find drawing part through parent slide relationships
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        from pptx.parts.diagram import DiagramDrawingPart

        # Get the parent slide part
        parent = self._parent
        while parent and not hasattr(parent, 'part'):
            parent = getattr(parent, '_parent', None)

        if not parent:
            return

        slide_part = parent.part

        # Find drawing part relationship
        drawing_rels = [
            rel
            for rel in slide_part.rels.values()
            if rel.reltype == RT.DIAGRAM_DRAWING
        ]

        if not drawing_rels:
            return

        # Update all text elements in drawing with texts from data model
        texts = self.text_content
        for rel in drawing_rels:
            drawing_part = rel.target_part
            if isinstance(drawing_part, DiagramDrawingPart):
                drawing_part.update_all_text_elements(texts)


class _SmartArtNodes(ParentedElementProxy):
    """Collection of SmartArt nodes."""

    def __init__(self, ptLst: CT_DiagramPointList, parent, data_part):
        super().__init__(ptLst, parent)
        self._data_part = data_part

    def __iter__(self) -> Iterator[SmartArtNode]:
        """Generate each node in the collection."""
        for pt in self._element.iter_pts():
            # Skip presentation nodes (type="pres") - they're layout metadata
            # Also skip transition nodes (parTrans, sibTrans)
            if pt.type not in ("pres", "parTrans", "sibTrans"):
                yield SmartArtNode(pt, self, self._data_part)

    def __len__(self) -> int:
        """Number of nodes."""
        return sum(1 for _ in self)


class SmartArtNode(ParentedElementProxy):
    """A single node in a SmartArt diagram."""

    def __init__(self, pt: CT_DiagramPoint, parent, data_part):
        super().__init__(pt, parent)
        self._pt = pt
        self._data_part = data_part

    @property
    def text(self) -> str:
        """Text content of this node."""
        return self._pt.text

    @text.setter
    def text(self, value: str) -> None:
        """Set text content of this node."""
        # Update data.xml
        self._pt.text = value

        # Update drawing.xml cache if it exists
        # The drawing part is a Microsoft extension that caches rendered output
        # We need to update it so apps like Keynote show the correct text
        if self._data_part is None:
            return

        # Get the parent SmartArt object by walking up the parent chain
        parent = self._parent
        while parent is not None:
            if isinstance(parent, SmartArt):
                parent._sync_drawing_part()
                break
            parent = getattr(parent, '_parent', None)

    @property
    def model_id(self) -> str | None:
        """Unique model ID of this node."""
        return self._pt.modelId

    @property
    def node_type(self) -> str | None:
        """Type of this node (e.g., 'doc', 'node', None)."""
        return self._pt.type
