"""Initializes lxml parser, particularly the custom element classes.

Also makes available a handful of functions that wrap its typical uses.

The parser is hardened against entity-expansion attacks (e.g. the "billion
laughs" attack), external-entity attacks ("XXE"), and DTD/network lookups.
See `docs/dev/security.rst` for the full trust model.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Type

from lxml import etree

from pptx.oxml.ns import NamespacePrefixedTag

if TYPE_CHECKING:
    from pptx.oxml.xmlchemy import BaseOxmlElement


# -- configure etree XML parser ----------------------------
# - `resolve_entities=False` disables entity substitution entirely which is
#   the principal defense against both the "billion laughs" exponential
#   entity-expansion attack and XML external entity ("XXE") attacks that
#   would otherwise read local files or open network connections.
# - `no_network=True` and `load_dtd=False` are the lxml defaults but are
#   set explicitly for clarity and as defense-in-depth.
# - `huge_tree=False` (the default) keeps libxml2's built-in limits on
#   maximum tree depth and text-node length in effect.
element_class_lookup = etree.ElementNamespaceClassLookup()
oxml_parser = etree.XMLParser(
    remove_blank_text=True,
    resolve_entities=False,
    no_network=True,
    load_dtd=False,
    huge_tree=False,
)
oxml_parser.set_element_class_lookup(element_class_lookup)


def parse_from_template(template_file_name: str):
    """Return an element loaded from the XML in the template file identified by `template_name`."""
    thisdir = os.path.split(__file__)[0]
    filename = os.path.join(thisdir, "..", "templates", "%s.xml" % template_file_name)
    with open(filename, "rb") as f:
        xml = f.read()
    return parse_xml(xml)


def parse_xml(xml: str | bytes):
    """Return root lxml element obtained by parsing XML character string in `xml`."""
    return etree.fromstring(xml, oxml_parser)


def register_element_cls(nsptagname: str, cls: Type[BaseOxmlElement]):
    """Register `cls` to be constructed when oxml parser encounters element having `nsptag_name`.

    `nsptag_name` is a string of the form `nspfx:tagroot`, e.g. `"w:document"`.
    """
    nsptag = NamespacePrefixedTag(nsptagname)
    namespace = element_class_lookup.get_namespace(nsptag.nsuri)
    namespace[nsptag.local_part] = cls


from pptx.oxml.action import CT_EmbeddedWAVAudioFile, CT_Hyperlink  # noqa: E402

register_element_cls("a:hlinkClick", CT_Hyperlink)
register_element_cls("a:hlinkHover", CT_Hyperlink)
register_element_cls("a:hlinkMouseOver", CT_Hyperlink)
register_element_cls("a:snd", CT_EmbeddedWAVAudioFile)


from pptx.oxml.chart.axis import (  # noqa: E402
    CT_AxisUnit,
    CT_AxPos,
    CT_CatAx,
    CT_ChartLines,
    CT_CrossBetween,
    CT_Crosses,
    CT_DateAx,
    CT_LblOffset,
    CT_Orientation,
    CT_Scaling,
    CT_TickLblPos,
    CT_TickMark,
    CT_ValAx,
)

register_element_cls("c:axPos", CT_AxPos)
register_element_cls("c:catAx", CT_CatAx)
register_element_cls("c:crossBetween", CT_CrossBetween)
register_element_cls("c:crosses", CT_Crosses)
register_element_cls("c:dateAx", CT_DateAx)
register_element_cls("c:lblOffset", CT_LblOffset)
register_element_cls("c:majorGridlines", CT_ChartLines)
register_element_cls("c:majorTickMark", CT_TickMark)
register_element_cls("c:majorUnit", CT_AxisUnit)
register_element_cls("c:minorTickMark", CT_TickMark)
register_element_cls("c:minorUnit", CT_AxisUnit)
register_element_cls("c:orientation", CT_Orientation)
register_element_cls("c:scaling", CT_Scaling)
register_element_cls("c:tickLblPos", CT_TickLblPos)
register_element_cls("c:valAx", CT_ValAx)


from pptx.oxml.chart.chart import (  # noqa: E402
    CT_Chart,
    CT_ChartSpace,
    CT_ExternalData,
    CT_PlotArea,
    CT_Style,
    CT_StyleEx,
)

register_element_cls("c:chart", CT_Chart)
register_element_cls("c:chartSpace", CT_ChartSpace)
register_element_cls("c:externalData", CT_ExternalData)
register_element_cls("c:plotArea", CT_PlotArea)
register_element_cls("c:style", CT_Style)
register_element_cls("c14:style", CT_StyleEx)


from pptx.oxml.chart.datalabel import CT_DLbl, CT_DLblPos, CT_DLbls  # noqa: E402

register_element_cls("c:dLbl", CT_DLbl)
register_element_cls("c:dLblPos", CT_DLblPos)
register_element_cls("c:dLbls", CT_DLbls)


from pptx.oxml.chart.legend import CT_Legend, CT_LegendPos  # noqa: E402

register_element_cls("c:legend", CT_Legend)
register_element_cls("c:legendPos", CT_LegendPos)


from pptx.oxml.chart.marker import CT_Marker, CT_MarkerSize, CT_MarkerStyle  # noqa: E402

register_element_cls("c:marker", CT_Marker)
register_element_cls("c:size", CT_MarkerSize)
register_element_cls("c:symbol", CT_MarkerStyle)


from pptx.oxml.chart.plot import (  # noqa: E402
    CT_Area3DChart,
    CT_AreaChart,
    CT_BarChart,
    CT_BarDir,
    CT_BubbleChart,
    CT_BubbleScale,
    CT_DoughnutChart,
    CT_GapAmount,
    CT_Grouping,
    CT_LineChart,
    CT_Overlap,
    CT_PieChart,
    CT_RadarChart,
    CT_ScatterChart,
)

register_element_cls("c:area3DChart", CT_Area3DChart)
register_element_cls("c:areaChart", CT_AreaChart)
register_element_cls("c:barChart", CT_BarChart)
register_element_cls("c:barDir", CT_BarDir)
register_element_cls("c:bubbleChart", CT_BubbleChart)
register_element_cls("c:bubbleScale", CT_BubbleScale)
register_element_cls("c:doughnutChart", CT_DoughnutChart)
register_element_cls("c:gapWidth", CT_GapAmount)
register_element_cls("c:grouping", CT_Grouping)
register_element_cls("c:lineChart", CT_LineChart)
register_element_cls("c:overlap", CT_Overlap)
register_element_cls("c:pieChart", CT_PieChart)
register_element_cls("c:radarChart", CT_RadarChart)
register_element_cls("c:scatterChart", CT_ScatterChart)


from pptx.oxml.chart.series import (  # noqa: E402
    CT_AxDataSource,
    CT_DPt,
    CT_ErrBars,
    CT_ErrBarType,
    CT_ErrDir,
    CT_ErrValType,
    CT_Lvl,
    CT_NumDataSource,
    CT_SeriesComposite,
    CT_StrVal_NumVal_Composite,
)

register_element_cls("c:bubbleSize", CT_NumDataSource)
register_element_cls("c:cat", CT_AxDataSource)
register_element_cls("c:dPt", CT_DPt)
register_element_cls("c:errBars", CT_ErrBars)
register_element_cls("c:errBarType", CT_ErrBarType)
register_element_cls("c:errDir", CT_ErrDir)
register_element_cls("c:errValType", CT_ErrValType)
register_element_cls("c:lvl", CT_Lvl)
register_element_cls("c:minus", CT_NumDataSource)
register_element_cls("c:plus", CT_NumDataSource)
register_element_cls("c:pt", CT_StrVal_NumVal_Composite)
register_element_cls("c:ser", CT_SeriesComposite)
register_element_cls("c:val", CT_NumDataSource)
register_element_cls("c:xVal", CT_NumDataSource)
register_element_cls("c:yVal", CT_NumDataSource)


from pptx.oxml.chart.shared import (  # noqa: E402
    CT_Boolean,
    CT_Boolean_Explicit,
    CT_Double,
    CT_Layout,
    CT_LayoutMode,
    CT_ManualLayout,
    CT_NumFmt,
    CT_Title,
    CT_Tx,
    CT_UnsignedInt,
)

register_element_cls("c:autoTitleDeleted", CT_Boolean_Explicit)
register_element_cls("c:autoUpdate", CT_Boolean)
register_element_cls("c:bubble3D", CT_Boolean)
register_element_cls("c:crossAx", CT_UnsignedInt)
register_element_cls("c:crossesAt", CT_Double)
register_element_cls("c:date1904", CT_Boolean)
register_element_cls("c:delete", CT_Boolean)
register_element_cls("c:idx", CT_UnsignedInt)
register_element_cls("c:invertIfNegative", CT_Boolean_Explicit)
register_element_cls("c:layout", CT_Layout)
register_element_cls("c:noEndCap", CT_Boolean)
register_element_cls("c:manualLayout", CT_ManualLayout)
register_element_cls("c:max", CT_Double)
register_element_cls("c:min", CT_Double)
register_element_cls("c:numFmt", CT_NumFmt)
register_element_cls("c:order", CT_UnsignedInt)
register_element_cls("c:overlay", CT_Boolean_Explicit)
register_element_cls("c:ptCount", CT_UnsignedInt)
register_element_cls("c:showCatName", CT_Boolean_Explicit)
register_element_cls("c:showLegendKey", CT_Boolean_Explicit)
register_element_cls("c:showPercent", CT_Boolean_Explicit)
register_element_cls("c:showSerName", CT_Boolean_Explicit)
register_element_cls("c:showVal", CT_Boolean_Explicit)
register_element_cls("c:smooth", CT_Boolean)
register_element_cls("c:title", CT_Title)
register_element_cls("c:tx", CT_Tx)
register_element_cls("c:varyColors", CT_Boolean)
register_element_cls("c:x", CT_Double)
register_element_cls("c:xMode", CT_LayoutMode)


from pptx.oxml.comments import (  # noqa: E402
    CT_Comment,
    CT_CommentAuthor,
    CT_CommentAuthorList,
    CT_CommentList,
)

register_element_cls("p:cm", CT_Comment)
register_element_cls("p:cmAuthor", CT_CommentAuthor)
register_element_cls("p:cmAuthorLst", CT_CommentAuthorList)
register_element_cls("p:cmLst", CT_CommentList)


from pptx.oxml.coreprops import CT_CoreProperties  # noqa: E402

register_element_cls("cp:coreProperties", CT_CoreProperties)


from pptx.oxml.extprops import CT_ExtendedProperties  # noqa: E402

register_element_cls("ep:Properties", CT_ExtendedProperties)


from pptx.oxml.dml.color import (  # noqa: E402
    CT_Color,
    CT_HslColor,
    CT_Percentage,
    CT_PresetColor,
    CT_SchemeColor,
    CT_ScRgbColor,
    CT_SRgbColor,
    CT_SystemColor,
)

register_element_cls("a:bgClr", CT_Color)
register_element_cls("a:fgClr", CT_Color)
register_element_cls("a:hslClr", CT_HslColor)
register_element_cls("a:lumMod", CT_Percentage)
register_element_cls("a:lumOff", CT_Percentage)
register_element_cls("a:prstClr", CT_PresetColor)
register_element_cls("a:schemeClr", CT_SchemeColor)
register_element_cls("a:scrgbClr", CT_ScRgbColor)
register_element_cls("a:srgbClr", CT_SRgbColor)
register_element_cls("a:sysClr", CT_SystemColor)


from pptx.oxml.dml.fill import (  # noqa: E402
    CT_Blip,
    CT_BlipFillProperties,
    CT_GradientFillProperties,
    CT_GradientStop,
    CT_GradientStopList,
    CT_GroupFillProperties,
    CT_LinearShadeProperties,
    CT_NoFillProperties,
    CT_PatternFillProperties,
    CT_RelativeRect,
    CT_SolidColorFillProperties,
)

register_element_cls("a:blip", CT_Blip)
register_element_cls("a:blipFill", CT_BlipFillProperties)
register_element_cls("a:gradFill", CT_GradientFillProperties)
register_element_cls("a:grpFill", CT_GroupFillProperties)
register_element_cls("a:gs", CT_GradientStop)
register_element_cls("a:gsLst", CT_GradientStopList)
register_element_cls("a:lin", CT_LinearShadeProperties)
register_element_cls("a:noFill", CT_NoFillProperties)
register_element_cls("a:pattFill", CT_PatternFillProperties)
register_element_cls("a:solidFill", CT_SolidColorFillProperties)
register_element_cls("a:srcRect", CT_RelativeRect)


from pptx.oxml.dml.effect import (  # noqa: E402
    CT_BlurEffect,
    CT_EffectList,
    CT_GlowEffect,
    CT_InnerShadowEffect,
    CT_OuterShadowEffect,
    CT_PresetShadowEffect,
    CT_ReflectionEffect,
    CT_SoftEdgesEffect,
)

register_element_cls("a:blur", CT_BlurEffect)
register_element_cls("a:effectLst", CT_EffectList)
register_element_cls("a:glow", CT_GlowEffect)
register_element_cls("a:innerShdw", CT_InnerShadowEffect)
register_element_cls("a:outerShdw", CT_OuterShadowEffect)
register_element_cls("a:prstShdw", CT_PresetShadowEffect)
register_element_cls("a:reflection", CT_ReflectionEffect)
register_element_cls("a:softEdge", CT_SoftEdgesEffect)


from pptx.oxml.dml.line import (  # noqa: E402
    CT_LineEndProperties,
    CT_PresetLineDashProperties,
)

register_element_cls("a:prstDash", CT_PresetLineDashProperties)
register_element_cls("a:headEnd", CT_LineEndProperties)
register_element_cls("a:tailEnd", CT_LineEndProperties)


from pptx.oxml.presentation import (  # noqa: E402
    CT_EmbeddedFontDataId,
    CT_EmbeddedFontList,
    CT_EmbeddedFontListEntry,
    CT_Extension,
    CT_ExtensionList,
    CT_Presentation,
    CT_Section,
    CT_SectionList,
    CT_SectionSlideId,
    CT_SectionSlideIdList,
    CT_SlideId,
    CT_SlideIdList,
    CT_SlideMasterIdList,
    CT_SlideMasterIdListEntry,
    CT_SlideSize,
)

register_element_cls("p:bold", CT_EmbeddedFontDataId)
register_element_cls("p:boldItalic", CT_EmbeddedFontDataId)
register_element_cls("p:embeddedFont", CT_EmbeddedFontListEntry)
register_element_cls("p:embeddedFontLst", CT_EmbeddedFontList)
register_element_cls("p:ext", CT_Extension)
register_element_cls("p:extLst", CT_ExtensionList)
register_element_cls("p:italic", CT_EmbeddedFontDataId)
register_element_cls("p:presentation", CT_Presentation)
register_element_cls("p:regular", CT_EmbeddedFontDataId)
register_element_cls("p:sldId", CT_SlideId)
register_element_cls("p:sldIdLst", CT_SlideIdList)
register_element_cls("p:sldMasterId", CT_SlideMasterIdListEntry)
register_element_cls("p:sldMasterIdLst", CT_SlideMasterIdList)
register_element_cls("p:sldSz", CT_SlideSize)
register_element_cls("p14:section", CT_Section)
register_element_cls("p14:sectionLst", CT_SectionList)
register_element_cls("p14:sldId", CT_SectionSlideId)
register_element_cls("p14:sldIdLst", CT_SectionSlideIdList)


from pptx.oxml.shapes.autoshape import (  # noqa: E402
    CT_AdjPoint2D,
    CT_CustomGeometry2D,
    CT_GeomGuide,
    CT_GeomGuideList,
    CT_NonVisualDrawingShapeProps,
    CT_Path2D,
    CT_Path2DClose,
    CT_Path2DLineTo,
    CT_Path2DList,
    CT_Path2DMoveTo,
    CT_PresetGeometry2D,
    CT_Shape,
    CT_ShapeNonVisual,
)

register_element_cls("a:avLst", CT_GeomGuideList)
register_element_cls("a:custGeom", CT_CustomGeometry2D)
register_element_cls("a:gd", CT_GeomGuide)
register_element_cls("a:close", CT_Path2DClose)
register_element_cls("a:lnTo", CT_Path2DLineTo)
register_element_cls("a:moveTo", CT_Path2DMoveTo)
register_element_cls("a:path", CT_Path2D)
register_element_cls("a:pathLst", CT_Path2DList)
register_element_cls("a:prstGeom", CT_PresetGeometry2D)
register_element_cls("a:pt", CT_AdjPoint2D)
register_element_cls("p:cNvSpPr", CT_NonVisualDrawingShapeProps)
register_element_cls("p:nvSpPr", CT_ShapeNonVisual)
register_element_cls("p:sp", CT_Shape)


from pptx.oxml.shapes.connector import (  # noqa: E402
    CT_Connection,
    CT_Connector,
    CT_ConnectorNonVisual,
    CT_NonVisualConnectorProperties,
)

register_element_cls("a:endCxn", CT_Connection)
register_element_cls("a:stCxn", CT_Connection)
register_element_cls("p:cNvCxnSpPr", CT_NonVisualConnectorProperties)
register_element_cls("p:cxnSp", CT_Connector)
register_element_cls("p:nvCxnSpPr", CT_ConnectorNonVisual)


from pptx.oxml.shapes.graphfrm import (  # noqa: E402
    CT_GraphicalObject,
    CT_GraphicalObjectData,
    CT_GraphicalObjectFrame,
    CT_GraphicalObjectFrameNonVisual,
    CT_OleObject,
)

register_element_cls("a:graphic", CT_GraphicalObject)
register_element_cls("a:graphicData", CT_GraphicalObjectData)
register_element_cls("p:graphicFrame", CT_GraphicalObjectFrame)
register_element_cls("p:nvGraphicFramePr", CT_GraphicalObjectFrameNonVisual)
register_element_cls("p:oleObj", CT_OleObject)


from pptx.oxml.shapes.groupshape import (  # noqa: E402
    CT_GroupShape,
    CT_GroupShapeNonVisual,
    CT_GroupShapeProperties,
)

register_element_cls("p:grpSp", CT_GroupShape)
register_element_cls("p:grpSpPr", CT_GroupShapeProperties)
register_element_cls("p:nvGrpSpPr", CT_GroupShapeNonVisual)
register_element_cls("p:spTree", CT_GroupShape)


from pptx.oxml.shapes.picture import CT_Picture, CT_PictureNonVisual  # noqa: E402

register_element_cls("p:blipFill", CT_BlipFillProperties)
register_element_cls("p:nvPicPr", CT_PictureNonVisual)
register_element_cls("p:pic", CT_Picture)


from pptx.oxml.shapes.shared import (  # noqa: E402
    CT_AlternateContent,
    CT_ApplicationNonVisualDrawingProps,
    CT_LineProperties,
    CT_NonVisualDrawingProps,
    CT_Placeholder,
    CT_Point2D,
    CT_PositiveSize2D,
    CT_ShapeProperties,
    CT_Transform2D,
)

register_element_cls("a:chExt", CT_PositiveSize2D)
register_element_cls("a:chOff", CT_Point2D)
register_element_cls("a:ext", CT_PositiveSize2D)
register_element_cls("a:ln", CT_LineProperties)
# -- table-cell border elements share the `CT_LineProperties` schema (per dml-main.xsd
# -- `CT_TableCellProperties`) but appear under distinct tag names.
register_element_cls("a:lnB", CT_LineProperties)
register_element_cls("a:lnBlToTr", CT_LineProperties)
register_element_cls("a:lnL", CT_LineProperties)
register_element_cls("a:lnR", CT_LineProperties)
register_element_cls("a:lnT", CT_LineProperties)
register_element_cls("a:lnTlToBr", CT_LineProperties)
register_element_cls("a:off", CT_Point2D)
register_element_cls("a:xfrm", CT_Transform2D)
register_element_cls("c:spPr", CT_ShapeProperties)
register_element_cls("mc:AlternateContent", CT_AlternateContent)
register_element_cls("p:cNvPr", CT_NonVisualDrawingProps)
register_element_cls("p:nvPr", CT_ApplicationNonVisualDrawingProps)
register_element_cls("p:ph", CT_Placeholder)
register_element_cls("p:spPr", CT_ShapeProperties)
register_element_cls("p:xfrm", CT_Transform2D)


from pptx.oxml.slide import (  # noqa: E402
    CT_Background,
    CT_BackgroundProperties,
    CT_CommonSlideData,
    CT_HeaderFooter,
    CT_NotesMaster,
    CT_NotesSlide,
    CT_Slide,
    CT_SlideLayout,
    CT_SlideLayoutIdList,
    CT_SlideLayoutIdListEntry,
    CT_SlideMaster,
    CT_SlideTiming,
    CT_SlideTransition,
    CT_TimeNodeList,
    CT_TLCommonTimeNodeData,
    CT_TLMediaNodeVideo,
    CT_TLTimeNodeParallel,
    CT_TLTimeNodeSequence,
    CT_TransitionMorph,
    CT_TransitionVariant,
)

register_element_cls("p:bg", CT_Background)
register_element_cls("p:bgPr", CT_BackgroundProperties)
register_element_cls("p:childTnLst", CT_TimeNodeList)
register_element_cls("p:cSld", CT_CommonSlideData)
register_element_cls("p:cTn", CT_TLCommonTimeNodeData)
register_element_cls("p:hf", CT_HeaderFooter)
register_element_cls("p:notes", CT_NotesSlide)
register_element_cls("p:notesMaster", CT_NotesMaster)
register_element_cls("p:par", CT_TLTimeNodeParallel)
register_element_cls("p:seq", CT_TLTimeNodeSequence)
register_element_cls("p:sld", CT_Slide)
register_element_cls("p:sldLayout", CT_SlideLayout)
register_element_cls("p:sldLayoutId", CT_SlideLayoutIdListEntry)
register_element_cls("p:sldLayoutIdLst", CT_SlideLayoutIdList)
register_element_cls("p:sldMaster", CT_SlideMaster)
register_element_cls("p:timing", CT_SlideTiming)
register_element_cls("p:tnLst", CT_TimeNodeList)
register_element_cls("p:transition", CT_SlideTransition)
register_element_cls("p:video", CT_TLMediaNodeVideo)
# -- transition-variant elements from `p:transition`'s inner <xsd:choice> --
for _variant_tag in (
    "p:blinds",
    "p:checker",
    "p:circle",
    "p:dissolve",
    "p:comb",
    "p:cover",
    "p:cut",
    "p:diamond",
    "p:fade",
    "p:newsflash",
    "p:plus",
    "p:pull",
    "p:push",
    "p:random",
    "p:randomBar",
    "p:split",
    "p:strips",
    "p:wedge",
    "p:wheel",
    "p:wipe",
    "p:zoom",
):
    register_element_cls(_variant_tag, CT_TransitionVariant)
del _variant_tag
# -- Office 2010 MORPH transition extension (different namespace) --
register_element_cls("p14:morph", CT_TransitionMorph)


from pptx.oxml.table import (  # noqa: E402
    CT_Table,
    CT_TableCell,
    CT_TableCellProperties,
    CT_TableCol,
    CT_TableGrid,
    CT_TableProperties,
    CT_TableRow,
)

register_element_cls("a:gridCol", CT_TableCol)
register_element_cls("a:tbl", CT_Table)
register_element_cls("a:tblGrid", CT_TableGrid)
register_element_cls("a:tblPr", CT_TableProperties)
register_element_cls("a:tc", CT_TableCell)
register_element_cls("a:tcPr", CT_TableCellProperties)
register_element_cls("a:tr", CT_TableRow)


from pptx.oxml.text import (  # noqa: E402
    CT_RegularTextRun,
    CT_TextAutonumberBullet,
    CT_TextBody,
    CT_TextBodyProperties,
    CT_TextBulletColor,
    CT_TextBulletSizePercent,
    CT_TextBulletSizePoint,
    CT_TextCharacterProperties,
    CT_TextCharBullet,
    CT_TextField,
    CT_TextFont,
    CT_TextLineBreak,
    CT_TextNoBullet,
    CT_TextNormalAutofit,
    CT_TextParagraph,
    CT_TextParagraphProperties,
    CT_TextSpacing,
    CT_TextSpacingPercent,
    CT_TextSpacingPoint,
)

register_element_cls("a:bodyPr", CT_TextBodyProperties)
register_element_cls("a:br", CT_TextLineBreak)
register_element_cls("a:buAutoNum", CT_TextAutonumberBullet)
register_element_cls("a:buChar", CT_TextCharBullet)
register_element_cls("a:buClr", CT_TextBulletColor)
register_element_cls("a:buFont", CT_TextFont)
register_element_cls("p:font", CT_TextFont)
register_element_cls("a:buNone", CT_TextNoBullet)
register_element_cls("a:buSzPct", CT_TextBulletSizePercent)
register_element_cls("a:buSzPts", CT_TextBulletSizePoint)
register_element_cls("a:cs", CT_TextFont)
register_element_cls("a:defRPr", CT_TextCharacterProperties)
register_element_cls("a:ea", CT_TextFont)
register_element_cls("a:endParaRPr", CT_TextCharacterProperties)
register_element_cls("a:fld", CT_TextField)
register_element_cls("a:latin", CT_TextFont)
register_element_cls("a:lnSpc", CT_TextSpacing)
register_element_cls("a:normAutofit", CT_TextNormalAutofit)
register_element_cls("a:r", CT_RegularTextRun)
register_element_cls("a:p", CT_TextParagraph)
register_element_cls("a:pPr", CT_TextParagraphProperties)
register_element_cls("c:rich", CT_TextBody)
register_element_cls("a:rPr", CT_TextCharacterProperties)
register_element_cls("a:spcAft", CT_TextSpacing)
register_element_cls("a:spcBef", CT_TextSpacing)
register_element_cls("a:spcPct", CT_TextSpacingPercent)
register_element_cls("a:spcPts", CT_TextSpacingPoint)
register_element_cls("a:txBody", CT_TextBody)
register_element_cls("c:txPr", CT_TextBody)
register_element_cls("p:txBody", CT_TextBody)


from pptx.oxml.theme import CT_OfficeStyleSheet  # noqa: E402

register_element_cls("a:theme", CT_OfficeStyleSheet)
