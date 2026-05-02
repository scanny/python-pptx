"""Chart-drawing (``c:userShapes``) part objects.

A :class:`ChartDrawingPart` carries the annotation shapes that PowerPoint
overlays on a chart — arrows, call-outs, text boxes, and so on — pinned
to the chart's plot area so they track the chart as it resizes.

On-disk, this is a relationship part referenced from a :class:`ChartPart`
via the ``chartUserShapes`` relationship type (``c:chartSpace`` also
carries a bare ``c:userShapes`` element whose ``r:id`` points at this
part). The part's XML root is ``cdr:userShapes`` and it holds an
unbounded sequence of anchor elements — each a
``cdr:relSizeAnchor`` or ``cdr:absSizeAnchor`` wrapping a single
drawing shape (``cdr:sp`` / ``cdr:grpSp`` / ``cdr:graphicFrame`` /
``cdr:cxnSp`` / ``cdr:pic``). The grammar is defined in
``dml-chartDrawing.xsd`` (ISO/IEC 29500-1).

This MVP supports **read-only** inspection: callers can enumerate the
anchors present, read their raw XML, and count them. Authoring new
annotation shapes is deferred — see
``docs/dev/analysis/chart-user-shapes.rst`` for the full design analysis
and the reason authoring isn't in this first drop (it would pull in a
whole parallel DrawingML shape-tree stack).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, cast

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn

if TYPE_CHECKING:
    from pptx.opc.packuri import PackURI
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.package import Package


# -- `cdr:relSizeAnchor` anchors a shape to a pair of fractional plot-area
# -- coordinates (`cdr:from` / `cdr:to`), so the shape resizes with the chart.
# -- `cdr:absSizeAnchor` anchors a shape at a fractional origin but with a
# -- fixed EMU extent (`cdr:from` / `cdr:ext`), so the shape keeps its size
# -- when the chart resizes. Both appear as direct children of
# -- `cdr:userShapes` per `dml-chartDrawing.xsd`. --
_ANCHOR_TAGS = (qn("cdr:relSizeAnchor"), qn("cdr:absSizeAnchor"))


class ChartDrawingPart(XmlPart):
    """Package part carrying a chart's user-shapes (annotation) drawing.

    Partnames follow the convention ``/ppt/charts/_rels/chartN.xml.rels``
    referencing ``/ppt/charts/drawings/drawingN.xml`` — the layout
    PowerPoint writes when it emits a chart-drawing alongside a chart.
    """

    partname_template = "/ppt/charts/drawings/drawing%d.xml"

    @classmethod
    def new(cls, package: Package, partname: PackURI | None = None) -> ChartDrawingPart:
        """Return a new, empty |ChartDrawingPart| added to `package`.

        When `partname` is omitted, a fresh partname is allocated from
        :attr:`partname_template`. The part starts with an empty
        ``cdr:userShapes`` root (no anchors).
        """
        if partname is None:
            partname = package.next_partname(cls.partname_template)
        element = cast(
            "BaseOxmlElement",
            parse_xml(f'<cdr:userShapes {nsdecls("cdr", "a", "r")}/>'),
        )
        return cls(partname, CT.DML_CHARTSHAPES, package, element)

    def iter_anchor_elements(self) -> Iterator[BaseOxmlElement]:
        """Generate each ``cdr:relSizeAnchor`` / ``cdr:absSizeAnchor`` child in document order.

        Each yielded element is the raw lxml element straight from the
        user-shapes XML — it has not been wrapped in a python-pptx proxy
        class. This is the MVP read-only affordance: enough to confirm
        that annotation shapes survive round-trip and enough to let
        adventurous callers inspect the raw XML, without committing the
        library to a proxy hierarchy for chart-drawing shapes (see
        ``docs/dev/analysis/chart-user-shapes.rst``).
        """
        for child in self._element.iterchildren():
            if child.tag in _ANCHOR_TAGS:
                yield cast("BaseOxmlElement", child)

    @property
    def anchor_count(self) -> int:
        """Number of top-level anchor elements in the user-shapes drawing.

        Equivalent to ``sum(1 for _ in self.iter_anchor_elements())`` but
        avoids materialising the iterator for callers that only need to
        check presence / length.
        """
        return sum(1 for _ in self.iter_anchor_elements())
