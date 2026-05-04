"""lxml custom element class for CT_GraphicalObjectFrame XML element."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.oxml import parse_xml
from pptx.oxml.chart.chart import CT_Chart
from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.simpletypes import XsdBoolean, XsdString
from pptx.oxml.table import CT_Table
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_CHARTEX,
    GRAPHIC_DATA_URI_MODEL_3D,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_SMART_ART,
    GRAPHIC_DATA_URI_TABLE,
)

if TYPE_CHECKING:
    from pptx.oxml.shapes.shared import (
        CT_ApplicationNonVisualDrawingProps,
        CT_NonVisualDrawingProps,
        CT_Transform2D,
    )


class CT_GraphicalObject(BaseOxmlElement):
    """`a:graphic` element.

    The container for the reference to or definition of the framed graphical object (table, chart,
    etc.).
    """

    graphicData: CT_GraphicalObjectData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphicData"
    )

    @property
    def chart(self) -> CT_Chart | None:
        """The `c:chart` grandchild element, or |None| if not present."""
        return self.graphicData.chart


class CT_GraphicalObjectData(BaseShapeElement):
    """`p:graphicData` element.

    The direct container for a table, a chart, or another graphical object.
    """

    chart: CT_Chart | None = ZeroOrOne("c:chart")  # pyright: ignore[reportAssignmentType]
    tbl: CT_Table | None = ZeroOrOne("a:tbl")  # pyright: ignore[reportAssignmentType]
    uri: str = RequiredAttribute("uri", XsdString)  # pyright: ignore[reportAssignmentType]

    @property
    def blob_rId(self) -> str | None:
        """Optional `r:id` attribute value of `p:oleObj` descendent element.

        This value is `None` when this `p:graphicData` element does not enclose an OLE object.
        This value could also be `None` if an enclosed OLE object does not specify this attribute
        (it is specified optional in the schema) but so far, all OLE objects we've encountered
        specify this value.
        """
        return None if self._oleObj is None else self._oleObj.rId

    @property
    def is_embedded_ole_obj(self) -> bool | None:
        """Optional boolean indicating an embedded OLE object.

        Returns `None` when this `p:graphicData` element does not enclose an OLE object. `True`
        indicates an embedded OLE object and `False` indicates a linked OLE object.
        """
        return None if self._oleObj is None else self._oleObj.is_embedded

    @property
    def progId(self) -> str | None:
        """Optional str value of "progId" attribute of `p:oleObj` descendent.

        This value identifies the "type" of the embedded object in terms of the application used
        to open it.

        This value is `None` when this `p:graphicData` element does not enclose an OLE object.
        This could also be `None` if an enclosed OLE object does not specify this attribute (it is
        specified optional in the schema) but so far, all OLE objects we've encountered specify
        this value.
        """
        return None if self._oleObj is None else self._oleObj.progId

    @property
    def showAsIcon(self) -> bool | None:
        """Optional value of "showAsIcon" attribute value of `p:oleObj` descendent.

        This value is `None` when this `p:graphicData` element does not enclose an OLE object. It
        is False when the `showAsIcon` attribute is omitted on the `p:oleObj` element.
        """
        return None if self._oleObj is None else self._oleObj.showAsIcon

    @property
    def model_3d(self) -> CT_Model3D | None:
        """Optional `am3d:model3D` child element carrying the embedded 3D model.

        Returns |None| when this `a:graphicData` does not enclose a 3D model (i.e.
        when `uri` is not :const:`~pptx.spec.GRAPHIC_DATA_URI_MODEL_3D`). The
        `am3d:model3D` element carries an `r:embed` attribute identifying the relationship
        to the embedded 3D-model part (`.glb` / `.obj` / `.fbx`) and scene/camera/lighting
        metadata. This is a read-only passthrough: the element is preserved verbatim on
        round-trip but not parsed further. See issue #410.
        """
        if self.uri != GRAPHIC_DATA_URI_MODEL_3D:
            return None
        models = cast("list[CT_Model3D]", self.xpath("./am3d:model3D"))
        return models[0] if models else None

    @property
    def cx_chart_rId(self) -> str | None:
        """Optional `r:id` attribute value of the `cx:chart` child element.

        Returns |None| when this `a:graphicData` does not enclose an Office 2016+
        extended (chartex) chart (i.e. when :attr:`uri` is not
        :const:`~pptx.spec.GRAPHIC_DATA_URI_CHARTEX`) or when the `cx:chart`
        child is absent. The relationship-id points at the ``cx:chartSpace``
        part carrying the chartex body — funnel, treemap, sunburst, waterfall,
        histogram/Pareto, box-and-whisker, or map.
        """
        if self.uri != GRAPHIC_DATA_URI_CHARTEX:
            return None
        rIds = cast("list[str]", self.xpath("./cx:chart/@r:id"))
        return rIds[0] if rIds else None

    @property
    def dgm_relIds(self) -> CT_DgmRelIds | None:
        """Optional `dgm:relIds` child element carrying the four SmartArt rIds.

        Returns |None| when this `a:graphicData` does not enclose a SmartArt diagram (i.e.
        when `uri` is not :const:`~pptx.spec.GRAPHIC_DATA_URI_SMART_ART`). The four
        attributes of `dgm:relIds` (``r:dm``, ``r:lo``, ``r:qs``, ``r:cs``) identify the
        diagramData, diagramLayout, diagramQuickStyle and diagramColors parts that
        together define a SmartArt graphic.
        """
        if self.uri != GRAPHIC_DATA_URI_SMART_ART:
            return None
        relIds = cast("list[CT_DgmRelIds]", self.xpath("./dgm:relIds"))
        return relIds[0] if relIds else None

    @property
    def _oleObj(self) -> CT_OleObject | None:
        """Optional `p:oleObj` element contained in this `p:graphicData' element.

        Returns `None` when this graphic-data element does not enclose an OLE object. Note that
        this returns the last `p:oleObj` element found. There can be more than one `p:oleObj`
        element because an `mc.AlternateContent` element may appear as the child of
        `p:graphicData` and that alternate-content subtree can contain multiple compatibility
        choices. The last one should suit best for reading purposes because it contains the lowest
        common denominator.
        """
        oleObjs = cast("list[CT_OleObject]", self.xpath(".//p:oleObj"))
        return oleObjs[-1] if oleObjs else None


class CT_GraphicalObjectFrame(BaseShapeElement):
    """`p:graphicFrame` element.

    A container for a table, a chart, or another graphical object.
    """

    nvGraphicFramePr: CT_GraphicalObjectFrameNonVisual = (  # pyright: ignore[reportAssignmentType]
        OneAndOnlyOne("p:nvGraphicFramePr")
    )
    xfrm: CT_Transform2D = OneAndOnlyOne("p:xfrm")  # pyright: ignore
    graphic: CT_GraphicalObject = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphic"
    )

    @property
    def chart(self) -> CT_Chart | None:
        """The `c:chart` great-grandchild element, or |None| if not present."""
        return self.graphic.chart

    @property
    def chart_rId(self) -> str | None:
        """The `rId` attribute of the `c:chart` great-grandchild element.

        |None| if not present.
        """
        chart = self.chart
        if chart is None:
            return None
        return chart.rId

    def get_or_add_xfrm(self) -> CT_Transform2D:
        """Return the required `p:xfrm` child element.

        Overrides version on BaseShapeElement.
        """
        return self.xfrm

    @property
    def graphicData(self) -> CT_GraphicalObjectData:
        """`a:graphicData` grandchild of this graphic-frame element."""
        return self.graphic.graphicData

    @property
    def graphicData_uri(self) -> str:
        """str value of `uri` attribute of `a:graphicData` grandchild."""
        return self.graphic.graphicData.uri

    @property
    def has_oleobj(self) -> bool:
        """`True` for graphicFrame containing an OLE object, `False` otherwise."""
        return self.graphicData.uri == GRAPHIC_DATA_URI_OLEOBJ

    @property
    def is_embedded_ole_obj(self) -> bool | None:
        """Optional boolean indicating an embedded OLE object.

        Returns `None` when this `p:graphicFrame` element does not enclose an OLE object. `True`
        indicates an embedded OLE object and `False` indicates a linked OLE object.
        """
        return self.graphicData.is_embedded_ole_obj

    @classmethod
    def new_chart_graphicFrame(
        cls, id_: int, name: str, rId: str, x: int, y: int, cx: int, cy: int
    ) -> CT_GraphicalObjectFrame:
        """Return a `p:graphicFrame` element tree populated with a chart element."""
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(id_, name, x, y, cx, cy)
        graphicData = graphicFrame.graphic.graphicData
        graphicData.uri = GRAPHIC_DATA_URI_CHART
        graphicData.append(CT_Chart.new_chart(rId))
        return graphicFrame

    @classmethod
    def new_graphicFrame(
        cls, id_: int, name: str, x: int, y: int, cx: int, cy: int
    ) -> CT_GraphicalObjectFrame:
        """Return a new `p:graphicFrame` element tree suitable for containing a table or chart.

        Note that a graphicFrame element is not a valid shape until it contains a graphical object
        such as a table.
        """
        return cast(
            CT_GraphicalObjectFrame,
            parse_xml(
                f"<p:graphicFrame {nsdecls('a', 'p')}>\n"
                f"  <p:nvGraphicFramePr>\n"
                f'    <p:cNvPr id="{id_}" name="{name}"/>\n'
                f"    <p:cNvGraphicFramePr>\n"
                f'      <a:graphicFrameLocks noGrp="1"/>\n'
                f"    </p:cNvGraphicFramePr>\n"
                f"    <p:nvPr/>\n"
                f"  </p:nvGraphicFramePr>\n"
                f"  <p:xfrm>\n"
                f'    <a:off x="{x}" y="{y}"/>\n'
                f'    <a:ext cx="{cx}" cy="{cy}"/>\n'
                f"  </p:xfrm>\n"
                f"  <a:graphic>\n"
                f"    <a:graphicData/>\n"
                f"  </a:graphic>\n"
                f"</p:graphicFrame>"
            ),
        )

    @classmethod
    def new_ole_object_graphicFrame(
        cls,
        id_: int,
        name: str,
        ole_object_rId: str,
        progId: str,
        icon_rId: str,
        x: int,
        y: int,
        cx: int,
        cy: int,
        imgW: int,
        imgH: int,
    ) -> CT_GraphicalObjectFrame:
        """Return newly-created `p:graphicFrame` for embedded OLE-object.

        `ole_object_rId` identifies the relationship to the OLE-object part.

        `progId` is a str identifying the object-type in terms of the application (program) used
        to open it. This becomes an attribute of the same name in the `p:oleObj` element.

        `icon_rId` identifies the relationship to an image part used to display the OLE-object as
        an icon (vs. a preview).
        """
        return cast(
            CT_GraphicalObjectFrame,
            parse_xml(
                f"<p:graphicFrame {nsdecls('a', 'p', 'r')}>\n"
                f"  <p:nvGraphicFramePr>\n"
                f'    <p:cNvPr id="{id_}" name="{name}"/>\n'
                f"    <p:cNvGraphicFramePr>\n"
                f'      <a:graphicFrameLocks noGrp="1"/>\n'
                f"    </p:cNvGraphicFramePr>\n"
                f"    <p:nvPr/>\n"
                f"  </p:nvGraphicFramePr>\n"
                f"  <p:xfrm>\n"
                f'    <a:off x="{x}" y="{y}"/>\n'
                f'    <a:ext cx="{cx}" cy="{cy}"/>\n'
                f"  </p:xfrm>\n"
                f"  <a:graphic>\n"
                f"    <a:graphicData"
                f'        uri="http://schemas.openxmlformats.org/presentationml/2006/ole">\n'
                f'      <p:oleObj showAsIcon="1"'
                f'                r:id="{ole_object_rId}"'
                f'                imgW="{imgW}"'
                f'                imgH="{imgH}"'
                f'                progId="{progId}">\n'
                f"        <p:embed/>\n"
                f"        <p:pic>\n"
                f"          <p:nvPicPr>\n"
                f'            <p:cNvPr id="0" name=""/>\n'
                f"            <p:cNvPicPr/>\n"
                f"            <p:nvPr/>\n"
                f"          </p:nvPicPr>\n"
                f"          <p:blipFill>\n"
                f'            <a:blip r:embed="{icon_rId}"/>\n'
                f"            <a:stretch>\n"
                f"              <a:fillRect/>\n"
                f"            </a:stretch>\n"
                f"          </p:blipFill>\n"
                f"          <p:spPr>\n"
                f"            <a:xfrm>\n"
                f'              <a:off x="{x}" y="{y}"/>\n'
                f'              <a:ext cx="{cx}" cy="{cy}"/>\n'
                f"            </a:xfrm>\n"
                f'            <a:prstGeom prst="rect">\n'
                f"              <a:avLst/>\n"
                f"            </a:prstGeom>\n"
                f"          </p:spPr>\n"
                f"        </p:pic>\n"
                f"      </p:oleObj>\n"
                f"    </a:graphicData>\n"
                f"  </a:graphic>\n"
                f"</p:graphicFrame>"
            ),
        )

    @classmethod
    def new_table_graphicFrame(
        cls, id_: int, name: str, rows: int, cols: int, x: int, y: int, cx: int, cy: int
    ) -> CT_GraphicalObjectFrame:
        """Return a `p:graphicFrame` element tree populated with a table element."""
        # -- coerce position/size to int so float inputs (e.g. `0.1 * slide_width`) don't
        # -- end up as non-integer values in the int-typed `a:off`/`a:ext` attributes.
        x, y, cx, cy = int(round(x)), int(round(y)), int(round(cx)), int(round(cy))
        graphicFrame = cls.new_graphicFrame(id_, name, x, y, cx, cy)
        graphicFrame.graphic.graphicData.uri = GRAPHIC_DATA_URI_TABLE
        graphicFrame.graphic.graphicData.append(CT_Table.new_tbl(rows, cols, cx, cy))
        return graphicFrame


class CT_GraphicalObjectFrameNonVisual(BaseOxmlElement):
    """`p:nvGraphicFramePr` element.

    This contains the non-visual properties of a graphic frame, such as name, id, etc.
    """

    cNvPr: CT_NonVisualDrawingProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cNvPr"
    )
    nvPr: CT_ApplicationNonVisualDrawingProps = (  # pyright: ignore[reportAssignmentType]
        OneAndOnlyOne("p:nvPr")
    )


class CT_DgmRelIds(BaseOxmlElement):
    """`dgm:relIds` element, the one child of `a:graphicData` for a SmartArt diagram.

    Carries four relationship-id attributes, one for each of the four SmartArt parts:

    * ``r:dm`` — diagramData (the semantic tree: nodes, text, connections)
    * ``r:lo`` — diagramLayout (algorithms and shape geometry)
    * ``r:qs`` — diagramQuickStyle (style variant within the layout)
    * ``r:cs`` — diagramColors (theme color transform)

    These four rIds are resolved against the slide part's relationships to retrieve the
    backing diagram parts. Together they constitute a single SmartArt graphic; losing
    any one renders the diagram invalid, which is why round-trip preservation is a
    foundation-level concern.
    """

    dm_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:dm", XsdString
    )
    lo_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:lo", XsdString
    )
    qs_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:qs", XsdString
    )
    cs_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:cs", XsdString
    )


class CT_Model3D(BaseOxmlElement):
    """`am3d:model3D` element, the one child of `a:graphicData` for an embedded 3D model.

    PowerPoint 365 (Office 2016+) "Insert > 3D Models" authoring produces this element
    under an `a:graphicData` whose `uri` is
    :data:`~pptx.spec.GRAPHIC_DATA_URI_MODEL_3D`. The element carries an `r:embed`
    attribute identifying the slide-relationship for the embedded 3D-model part
    (typically a binary glTF `.glb`; `.obj` / `.fbx` are also observed) plus an `ext`
    attribute recording the file extension. Additional descendants describe camera
    position, scene lighting, bounding extents, and (when the model is animated) a
    currently-selected animation and camera pose. python-pptx currently preserves the
    element verbatim on round-trip without parsing its scene parameters — issue #410
    ships *detection and passthrough* as the MVP.
    """

    embed_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:embed", XsdString
    )
    ext: str | None = OptionalAttribute("ext", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_OleObject(BaseOxmlElement):
    """`p:oleObj` element, container for an OLE object (e.g. Excel file).

    An OLE object can be either linked or embedded (hence the name).
    """

    progId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "progId", XsdString
    )
    rId: str | None = OptionalAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]
    showAsIcon: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showAsIcon", XsdBoolean, default=False
    )

    @property
    def is_embedded(self) -> bool:
        """True when this OLE object is embedded, False when it is linked."""
        return len(self.xpath("./p:embed")) > 0
