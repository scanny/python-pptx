======================================
Miscellaneous Open XML schema elements
======================================

Analysis
========

All shape types (except *contentPart*) have a required cNvPr element as their
first grandchild, for example::

   <p:sp>
     <p:nvSpPr>
       <p:cNvPr id="2" name="Title Placeholder 1"/>
       ...

So the XPath expression::

   slide.xpath('./p:cSld/p:spTree/*[position()>2]/*[1]', namespaces=nsmap)

selects all cNvPr elements in all the shapes in *slide* (actually, this
doesn't take the possible *extLst* element into account, but is roughly
correct).


782 dml - cNvPr -- CT_NonVisualDrawingProps -- §19.3.1.12
---------------------------------------------------------

attributes
^^^^^^^^^^

================  ===  ===================  ==========
name              use  type                 default
================  ===  ===================  ==========
id                 1   ST_DrawingElementId  no default
name               1   xsd:string           no default
descr              ?   xsd:string           no default
hidden             ?   xsd:boolean          false
title              ?   xsd:string           ""
================  ===  ===================  ==========


child elements
^^^^^^^^^^^^^^

================  ===  ================================  ========
name               #   type                              line
================  ===  ================================  ========
hlinkClick         ?   CT_Hyperlink
hlinkHover         ?   CT_Hyperlink
extLst             ?   CT_OfficeArtExtensionList
================  ===  ================================  ========


First-level slide XML components
================================

1282 - spTree -- CT_GroupShape
------------------------------

attributes
^^^^^^^^^^

None.


child elements
^^^^^^^^^^^^^^

============  ===  =======================  ========
name           #   type                     line
============  ===  =======================  ========
nvGrpSpPr      1   CT_GroupShapeNonVisual   1273
grpSpPr        1   CT_GroupShapeProperties
sp            ? |  CT_Shape                 1209
grpSp         ? |  CT_GroupShape            1282
graphicFrame  ? |  CT_GraphicalObjectFrame  1263
cxnSp         ? |  CT_Connector             1228
pic           ? |  CT_Picture               1245
contentPart   ?    CT_Rel
extLst         ?   CT_ExtensionListModify   775>767
============  ===  =======================  ========


1273 - nvGrpSpPr -- CT_GroupShapeNonVisual
------------------------------------------

attributes
^^^^^^^^^^

None.


child elements
^^^^^^^^^^^^^^

============  ===  ===================================  ========
name           #   type                                 line
============  ===  ===================================  ========
cNvPr          1   CT_NonVisualDrawingProps             782 dml
cNvGrpSpPr     1   CT_NonVisualGroupDrawingShapeProps
nvPr           1   CT_ApplicationNonVisualDrawingProps
============  ===  ===================================  ========


1209 - sp -- CT_Shape
------------------------------------------

attributes
^^^^^^^^^^

================  ===  ==================  ========
name              use  type                default
================  ===  ==================  ========
useBgFill          ?   xsd:boolean         false
================  ===  ==================  ========


child elements
^^^^^^^^^^^^^^

======  ===  ======================  ========
name     #   type                    line
======  ===  ======================  ========
nvSpPr   1   CT_ShapeNonVisual       1201
spPr     1   CT_ShapeProperties      2210 dml
style    ?   CT_ShapeStyle           2245 dml
txBody   ?   CT_TextBody             2640 dml
extLst   ?   CT_ExtensionListModify  775>767
======  ===  ======================  ========


1263 - graphicFrame -- CT_GraphicalObjectFrame
----------------------------------------------

attributes
^^^^^^^^^^

================  ===  ===================  ==========
name              use  type                 default
================  ===  ===================  ==========
bwMode             ?   a:ST_BlackWhiteMode  no default
================  ===  ===================  ==========


child elements
^^^^^^^^^^^^^^

================  ===  ================================  ========
name               #   type                              line
================  ===  ================================  ========
nvGraphicFramePr   1   CT_GraphicalObjectFrameNonVisual  1254
xfrm               1   a:CT_Transform2D                  613 dml
a:graphic          1   a:CT_GraphicalObject              835 dml
extLst             ?   CT_ExtensionListModify            775>767
================  ===  ================================  ========


1228 - cxnSp -- CT_Connector
----------------------------

attributes
^^^^^^^^^^

None.


child elements
^^^^^^^^^^^^^^

================  ===  ================================  ========
name               #   type                              line
================  ===  ================================  ========
nvCxnSpPr          1   CT_ConnectorNonVisual             1219
spPr               1   a:CT_ShapeProperties              2210 dml
style              ?   a:CT_ShapeStyle                   2245 dml
extLst             ?   CT_ExtensionListModify            775>767
================  ===  ================================  ========


1297 - contentPart -- CT_Rel  -- §19.3.1.13
-------------------------------------------

From ISO/IEC 29500-1:

   This element specifies a reference to XML content in a format not defined
   by ISO/IEC 29500. [Note: This part allows the native use of other commonly
   used interchange formats, such as:
   
   * MathML (http://www.w3.org/TR/MathML2/)
   * SMIL (http://www.w3.org/TR/REC-smil/)
   * SVG (http://www.w3.org/TR/SVG11/)

attributes
^^^^^^^^^^

================  ===  ===================  ============
name              use  type                 default
================  ===  ===================  ============
r:id               1   *not specified*      *no default*
================  ===  ===================  ============


child elements
^^^^^^^^^^^^^^

None.


