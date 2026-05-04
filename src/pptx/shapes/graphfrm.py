"""Graphic Frame shape and related objects.

A graphic frame is a common container for table, chart, smart art, and media
objects.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from lxml import etree

from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.shapes.base import BaseShape
from pptx.shapes.model3d import Model3D
from pptx.shared import ParentedElementProxy
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_CHARTEX,
    GRAPHIC_DATA_URI_MODEL_3D,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_SMART_ART,
    GRAPHIC_DATA_URI_TABLE,
)
from pptx.table import Table
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.chart import Chart
    from pptx.dml.effect import ShadowFormat
    from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectData, CT_GraphicalObjectFrame
    from pptx.parts.chart import ChartPart
    from pptx.parts.slide import BaseSlidePart
    from pptx.types import ProvidesPart


class GraphicFrame(BaseShape):
    """Container shape for table, chart, smart art, and media objects.

    Corresponds to a `p:graphicFrame` element in the shape tree.
    """

    def __init__(self, graphicFrame: CT_GraphicalObjectFrame, parent: ProvidesPart):
        super().__init__(graphicFrame, parent)
        self._graphicFrame = graphicFrame

    @property
    def chart(self) -> Chart:
        """The |Chart| object containing the chart in this graphic frame.

        Raises |ValueError| if this graphic frame does not contain a chart. Raises
        |NotImplementedError| when the graphic frame contains an Office 2016+ extended
        (``cx:``/chartex) chart — funnel, treemap, sunburst, waterfall, histogram,
        box-and-whisker, or map. Such shapes can still be inspected via :attr:`has_chart`
        and :attr:`chart_type` (which returns :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`) and
        are preserved unchanged on save.
        """
        if not self.has_chart:
            raise ValueError("shape does not contain a chart")
        if self.has_chartex:
            raise NotImplementedError(
                "chart object not available for Office 2016+ chartex (extended) chart types;"
                " check .chart_type for XL_CHART_TYPE.UNSUPPORTED_CHARTEX"
            )
        return self.chart_part.chart

    @property
    def chart_part(self) -> ChartPart:
        """The |ChartPart| object containing the chart in this graphic frame."""
        chart_rId = self._graphicFrame.chart_rId
        if chart_rId is None:
            raise ValueError("this graphic frame does not contain a chart")
        return cast("ChartPart", self.part.related_part(chart_rId))

    @property
    def chart_type(self) -> XL_CHART_TYPE:
        """Member of :ref:`XlChartType` identifying the chart contained in this graphic frame.

        For legacy (``c:``) charts this is the type reported by the chart itself. For Office
        2016+ extended (``cx:``) charts, which this library does not yet read in detail, the
        value is :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`. Raises |ValueError| if this
        graphic frame does not contain a chart.

        .. versionadded:: 2026.05.0
        """
        if self.has_chartex:
            return XL_CHART_TYPE.UNSUPPORTED_CHARTEX
        if not self.has_chart:
            raise ValueError("shape does not contain a chart")
        return self.chart_part.chart.chart_type

    @property
    def chartex_type(self) -> str | None:
        """Layout-id string of the chartex chart in this graphic frame, or |None|.

        Returns |None| when this graphic frame does not contain an Office 2016+
        extended (``cx:``) chart (i.e. when :attr:`has_chartex` is |False|) or when
        the chartex part is unresolvable or carries no ``cx:series`` with a
        ``@layoutId`` attribute.

        When |True|-ish, the returned string is the raw ``cx:series/@layoutId``
        value of the first ``cx:series`` in the referenced ``cx:chartSpace`` part.
        Recognised Office 2016+ values include:
        ``"funnel"``, ``"treemap"``, ``"sunburst"``, ``"waterfall"``,
        ``"boxWhisker"`` (box-and-whisker), ``"clusteredColumn"``
        (histogram — distinguish Pareto by presence of a sibling
        ``paretoLine`` series), ``"paretoLine"``, and ``"regionMap"``.

        This is a discovery / round-trip hint; structured read/write of
        chartex charts is still deferred. See
        ``docs/dev/analysis/chartex-foundation.rst`` and per-kind design notes.
        """
        if not self.has_chartex:
            return None
        rId = self._graphicFrame.graphicData.cx_chart_rId
        if rId is None:
            return None
        try:
            chartex_part = self.part.related_part(rId)
        except KeyError:
            return None
        blob = cast("bytes | None", getattr(chartex_part, "blob", None))
        if not blob:
            return None
        # -- `etree.XMLSyntaxError` is not exposed in the typings used here, so
        # -- a narrow catch would require a pyright ignore. A malformed chartex
        # -- part in an already-loaded package is a pathological / unlikely case
        # -- — a raw `Exception` catch is acceptable for a best-effort
        # -- read-only discovery helper.
        try:
            root = parse_xml(blob)
        except Exception:
            return None
        series = root.find(".//" + qn("cx:series"))
        if series is None:
            return None
        layoutId = series.get("layoutId")
        return layoutId

    @property
    def has_chart(self) -> bool:
        """|True| if this graphic frame contains a chart object, |False| otherwise.

        This includes both legacy charts (``c:chart``) and Office 2016+ extended charts
        (``cx:chart``). When |True|, use :attr:`chart_type` to distinguish between them;
        :attr:`chart` is only available for legacy charts.
        """
        uri = self._graphicFrame.graphicData_uri
        return uri in (GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_CHARTEX)

    @property
    def has_chartex(self) -> bool:
        """|True| if this graphic frame contains an Office 2016+ extended (``cx:``) chart.

        Extended chart types include funnel, treemap, sunburst, waterfall, histogram,
        box-and-whisker, and map. These shapes are surfaced for discoverability and
        round-trip preservation; detailed read/write is not yet supported.

        .. versionadded:: 2026.05.0
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_CHARTEX

    @property
    def has_model_3d(self) -> bool:
        """|True| if this graphic frame contains an embedded 3D model, |False| otherwise.

        PowerPoint 365 (Office 2016+) "Insert > 3D Models" produces a graphic-frame
        with an ``am3d:model3D`` child under ``a:graphicData`` whose ``uri`` matches
        :data:`~pptx.spec.GRAPHIC_DATA_URI_MODEL_3D`. When |True|, :attr:`model_3d`
        exposes a :class:`~pptx.shapes.model3d.Model3D` proxy and :attr:`model_3d_xml`
        returns the raw XML of the ``am3d:model3D`` element. This is a read-only
        passthrough MVP — 3D-model shapes are *discoverable* and
        *round-trip-preserved*, but authoring is deliberately out of scope. See
        ``docs/dev/analysis/model-3d.rst`` for the roadmap.

        .. versionadded:: 2026.05.0
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_MODEL_3D

    @property
    def has_smart_art(self) -> bool:
        """|True| if this graphic frame contains a SmartArt diagram, |False| otherwise.

        When |True|, :attr:`smart_art` exposes the four-part SmartArt graphic as raw XML.
        This is a Foundation-F9 scaffolding capability: SmartArt shapes are *discoverable*
        and *round-trip-preserved*, but full read/write of the diagram tree, layout, and
        styling requires the four-part feature work that layers on top of this foundation.
        See ``docs/dev/analysis/f9-smartart.rst`` for the roadmap.

        .. versionadded:: 2026.05.0
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_SMART_ART

    @property
    def has_table(self) -> bool:
        """|True| if this graphic frame contains a table object, |False| otherwise.

        When |True|, the table object can be accessed using the `.table` property.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_TABLE

    @property
    def model_3d(self) -> Model3D:
        """A |Model3D| object providing read-only access to this 3D model's embedded part.

        Raises |ValueError| if this graphic frame does not contain a 3D model (i.e.
        :attr:`has_model_3d` is |False|).

        The returned object exposes the embedded model's relationship id
        (:attr:`~pptx.shapes.model3d.Model3D.embedded_rel_id`), its raw bytes
        (:attr:`~pptx.shapes.model3d.Model3D.media_blob`), and its extension
        (:attr:`~pptx.shapes.model3d.Model3D.ext`). This is a read-only MVP; structured
        authoring of camera/scene/lighting parameters is deferred to a follow-up
        iteration of issue #410. See ``docs/dev/analysis/model-3d.rst``.

        .. versionadded:: 2026.05.0
        """
        if not self.has_model_3d:
            raise ValueError("shape does not contain a 3D model")
        model3D = self._graphicFrame.graphicData.model_3d
        if model3D is None:
            raise ValueError("shape does not contain a 3D model")
        return Model3D(model3D, self._parent)

    @property
    def model_3d_xml(self) -> str | None:
        """Raw XML of the ``am3d:model3D`` element, or |None| when not present.

        Returns a Unicode string containing the serialized ``am3d:model3D`` subtree
        (including whatever namespace declarations lxml emits to make it well-formed on
        its own) when :attr:`has_model_3d` is |True|. Returns |None| when the graphic
        frame does not contain a 3D model or when the ``am3d:model3D`` child is absent.

        The returned string is a snapshot — mutating it has no effect on the
        presentation. Callers wanting to touch the element should operate on
        ``shape.element`` directly.

        .. versionadded:: 2026.05.0
        """
        if not self.has_model_3d:
            return None
        model3D = self._graphicFrame.graphicData.model_3d
        if model3D is None:
            return None
        return etree.tostring(model3D, encoding="unicode")

    @property
    def ole_format(self) -> _OleFormat:
        """_OleFormat object for this graphic-frame shape.

        Raises `ValueError` on a GraphicFrame instance that does not contain an OLE object.

        An shape that contains an OLE object will have `.shape_type` of either
        `EMBEDDED_OLE_OBJECT` or `LINKED_OLE_OBJECT`.
        """
        if not self._graphicFrame.has_oleobj:
            raise ValueError("not an OLE-object shape")
        return _OleFormat(self._graphicFrame.graphicData, self._parent)

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """Unconditionally raises |NotImplementedError|.

        Access to the shadow effect for graphic-frame objects is content-specific (i.e. different
        for charts, tables, etc.) and has not yet been implemented.
        """
        raise NotImplementedError("shadow property on GraphicFrame not yet supported")

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Optional member of `MSO_SHAPE_TYPE` identifying the type of this shape.

        Possible values are `MSO_SHAPE_TYPE.CHART`, `MSO_SHAPE_TYPE.TABLE`,
        `MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT`, `MSO_SHAPE_TYPE.LINKED_OLE_OBJECT`,
        `MSO_SHAPE_TYPE.IGX_GRAPHIC` (for SmartArt).

        This value is `None` when none of those types apply.
        """
        graphicData_uri = self._graphicFrame.graphicData_uri
        if graphicData_uri in (GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_CHARTEX):
            return MSO_SHAPE_TYPE.CHART
        elif graphicData_uri == GRAPHIC_DATA_URI_TABLE:
            return MSO_SHAPE_TYPE.TABLE
        elif graphicData_uri == GRAPHIC_DATA_URI_OLEOBJ:
            return (
                MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
                if self._graphicFrame.is_embedded_ole_obj
                else MSO_SHAPE_TYPE.LINKED_OLE_OBJECT
            )
        elif graphicData_uri == GRAPHIC_DATA_URI_SMART_ART:
            return MSO_SHAPE_TYPE.IGX_GRAPHIC
        else:
            return None  # pyright: ignore[reportReturnType]

    @property
    def smart_art(self) -> SmartArt:
        """A |SmartArt| object providing read-only access to this SmartArt diagram's parts.

        Raises |ValueError| if this graphic frame does not contain a SmartArt diagram (i.e.
        :attr:`has_smart_art` is |False|).

        The returned object exposes the four XML parts that together define a SmartArt
        graphic (data, layout, colors, quickStyle) as raw-bytes accessors. This is a
        Foundation-F9 MVP: structured authoring / editing of SmartArt nodes and layout
        selection is not yet implemented; see ``docs/dev/analysis/f9-smartart.rst``.

        .. versionadded:: 2026.05.0
        """
        if not self.has_smart_art:
            raise ValueError("shape does not contain SmartArt")
        return SmartArt(self._graphicFrame.graphicData, self._parent)

    @property
    def table(self) -> Table:
        """The |Table| object contained in this graphic frame.

        Raises |ValueError| if this graphic frame does not contain a table.
        """
        if not self.has_table:
            raise ValueError("shape does not contain a table")
        tbl = self._graphicFrame.graphic.graphicData.tbl
        return Table(tbl, self)


class _OleFormat(ParentedElementProxy):
    """Provides attributes on an embedded OLE object."""

    part: BaseSlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, graphicData: CT_GraphicalObjectData, parent: ProvidesPart):
        super().__init__(graphicData, parent)
        self._graphicData = graphicData

    @property
    def blob(self) -> bytes | None:
        """Optional bytes of OLE object, suitable for loading or saving as a file.

        This value is `None` if the embedded object does not represent a "file".
        """
        blob_rId = self._graphicData.blob_rId
        if blob_rId is None:
            return None
        return self.part.related_part(blob_rId).blob

    @property
    def prog_id(self) -> str | None:
        """str "progId" attribute of this embedded OLE object.

        The progId is a str like "Excel.Sheet.12" that identifies the "file-type" of the embedded
        object, or perhaps more precisely, the application (aka. "server" in OLE parlance) to be
        used to open this object.
        """
        return self._graphicData.progId

    @property
    def show_as_icon(self) -> bool | None:
        """True when OLE object should appear as an icon (rather than preview)."""
        return self._graphicData.showAsIcon


class SmartArt(ParentedElementProxy):
    """Provides read-only access to the four parts of a SmartArt graphic.

    A SmartArt graphic is defined by four interlinked XML parts referenced from a single
    ``dgm:relIds`` child of ``a:graphicData``:

    * *diagramData* (``r:dm``) — the semantic node/connection tree (``dgm:dataModel``)
      plus presentation element index. This is the authoring surface: adding or removing
      nodes, editing node text, or reordering branches is a diagramData operation.
    * *diagramLayout* (``r:lo``) — the layout definition (``dgm:layoutDef``) which is the
      algorithmic program that converts the data tree into positioned shapes. Layouts are
      usually referenced from the Office-installed library (e.g. "Hierarchy", "Cycle",
      "Process") rather than authored from scratch.
    * *diagramColors* (``r:cs``) — a color-variation definition (``dgm:colorsDef``)
      binding theme color slots to a color transform.
    * *diagramQuickStyle* (``r:qs``) — the style variation (``dgm:styleDef``) picking fill
      / line / effect recipes within a layout.

    MVP scope (Foundation-F9): :attr:`data_xml`, :attr:`layout_xml`, :attr:`colors_xml`,
    and :attr:`quick_style_xml` return the raw bytes of the referenced parts. Each is
    |None| when the corresponding rId is missing or unresolvable. No structured
    editing API is provided at this tier; see ``docs/dev/analysis/f9-smartart.rst`` for
    the per-subsystem roadmap.

    .. versionadded:: 2026.05.0
    """

    part: BaseSlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, graphicData: CT_GraphicalObjectData, parent: ProvidesPart):
        super().__init__(graphicData, parent)
        self._graphicData = graphicData

    @property
    def colors_xml(self) -> bytes | None:
        """Raw XML bytes of the diagramColors part, or |None| if unresolvable.

        .. versionadded:: 2026.05.0
        """
        return self._related_blob("cs_rId")

    @property
    def data_xml(self) -> bytes | None:
        """Raw XML bytes of the diagramData part, or |None| if unresolvable.

        This is the authoring surface of a SmartArt graphic — the ``dgm:dataModel``
        element and its tree of ``dgm:pt`` (points / nodes) and ``dgm:cxn`` (connections).

        .. versionadded:: 2026.05.0
        """
        return self._related_blob("dm_rId")

    @property
    def layout_xml(self) -> bytes | None:
        """Raw XML bytes of the diagramLayout part, or |None| if unresolvable.

        .. versionadded:: 2026.05.0
        """
        return self._related_blob("lo_rId")

    @property
    def quick_style_xml(self) -> bytes | None:
        """Raw XML bytes of the diagramQuickStyle part, or |None| if unresolvable.

        .. versionadded:: 2026.05.0
        """
        return self._related_blob("qs_rId")

    def _related_blob(self, rId_attr: str) -> bytes | None:
        """Return blob bytes of the part related by the rId carried in `rId_attr`.

        Returns |None| when `dgm:relIds` is absent (non-SmartArt diagram), when the named
        attribute has no value, or when the relationship cannot be resolved.
        """
        relIds = self._graphicData.dgm_relIds
        if relIds is None:
            return None
        rId = getattr(relIds, rId_attr)
        if rId is None:
            return None
        try:
            return self.part.related_part(rId).blob
        except KeyError:
            return None
