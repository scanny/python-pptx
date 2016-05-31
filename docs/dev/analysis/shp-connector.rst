
Line/Connector Shape
====================

Lines are a sub-category of auto shape that differ in certain properties and
behaviors.

Initial scope is limited to connectors, shapes based on the ``<p:cxnSp>``
element and having a preset geometry such as ``'line'``.


Notes
-----

* Might be a good idea to get LineFormat (e.g. shape.line) object in place
  before developing this, so it can be connected up to set color and so on.
* Might want to move adjustments to shapes.shared
* Probably worth adding LockAspectRatio to shape properties along the way, also
  have a look at shp-properties.feature to see if it's current with coding
  standards. Rotation is another property that might be good to pop in along
  the way.
* Rename AdjustmentCollection to Adjustments
* Might want to review AdjustmentCollection for thin-proxy refactoring
* Add backlog item for ConnectorFormat object, allowing connectors to be
  attached to shape connection sites such that their connected end points are
  moved when the shape is moved.


Analysis
--------

* Seems like the most common lines are connectors
* Connectors are a distinct shape type. They are very similar to
  regular ``<p:sp>``-based auto shapes, but lack a text frame.
* Hypothesis: There are really two types, connectors and free-form.
  
  + Connectors are based on the ``<p:cxnSp>`` element and have a preset
    geometry (``<a:prstGeom>`` child of ``<p:spPr>``). 
  + Free-form lines are based on the ``<p:sp>`` element and have a custom
    geometry (``<a:custGeom>`` child of ``<p:spPr>``).

* Connectors don't have a fill. Free-form shapes do. Fill of free-form shapes
  extends between the line and a line connecting the end points, whether
  present or not. Since the lines can cross, this produces some possibly
  surprising fill behaviors; there is no clear concept of inside and outside
  for such a shape.


Experiments
-----------

* [ ] see how a connector returns for `connector.AutoShapeType`
* [ ] what behaviors for height and width on a connector? There are no StartX
      or EndY properties on shape.


Protocols
---------

Properties::

    >>> connector.is_connector
    True
    >>> connector.type
    MSO_SHAPE_TYPE.LINE
    >>> connector.start_x
    914400
    >>> connector.end_y
    914400
    >>> connector.adjustments
    <pptx.shapes.shared.Adjustments instance at 0x123456789>

Creation protocol::

    >>> line = shapes.add_connector(
    ...     MSO_CONNECTOR.STRAIGHT, start_x, start_y, end_x, end_y
    ... )


MS API Protocol
---------------



Behaviors
---------

UI display
~~~~~~~~~~

Connector shapes do not display a bounding box when selected in the UI. Rather
their start and end points are highlighted with a small gray circle that can be
moved. If there is an adjustment, such as a mid-point, it is indicated in the
normal way with a small yellow diamond. These endpoints and adjustment diamond
cannot be individually selected.


Naming
~~~~~~

What determines the name automatically applied to a line or connector shape?
Can it change after shape creation, for example by applying an arrowhead?


Enumerations
------------

MsoConnectorType

http://msdn.microsoft.com/en-us/library/office/ff860918(v=office.15).aspx

=====================  =====  ===============================================
Name                   Value  Description
=====================  =====  ===============================================
msoConnectorCurve        3    Curved connector.
msoConnectorElbow        2    Elbow connector.
msoConnectorStraight     1    Straight line connector.
msoConnectorTypeMixed   -2    Return value only; indicates a combination of
                              the other states.
=====================  =====  ===============================================


Specimen XML
------------

.. highlight:: xml

Straight line (Connector)::

   <p:cxnSp>
     <p:nvCxnSpPr>
       <p:cNvPr id="5" name="Straight Connector 4"/>
       <p:cNvCxnSpPr/>
       <p:nvPr/>
     </p:nvCxnSpPr>
     <p:spPr>
       <a:xfrm>
         <a:off x="950964" y="1101493"/>
         <a:ext cx="1257921" cy="0"/>
       </a:xfrm>
       <a:prstGeom prst="line">
         <a:avLst/>
       </a:prstGeom>
     </p:spPr>
     <p:style>
       <a:lnRef idx="2">
         <a:schemeClr val="accent1"/>
       </a:lnRef>
       <a:fillRef idx="0">
         <a:schemeClr val="accent1"/>
       </a:fillRef>
       <a:effectRef idx="1">
         <a:schemeClr val="accent1"/>
       </a:effectRef>
       <a:fontRef idx="minor">
         <a:schemeClr val="tx1"/>
       </a:fontRef>
     </p:style>
   </p:cxnSp>

Straight arrow Connector::

   <p:cxnSp>
     <p:nvCxnSpPr>
       <p:cNvPr id="7" name="Straight Arrow Connector 6"/>
       <p:cNvCxnSpPr/>
       <p:nvPr/>
     </p:nvCxnSpPr>
     <p:spPr>
       <a:xfrm>
         <a:off x="950964" y="1673307"/>
         <a:ext cx="1257921" cy="0"/>
       </a:xfrm>
       <a:prstGeom prst="straightConnector1">
         <a:avLst/>
       </a:prstGeom>
       <a:ln>
         <a:tailEnd type="arrow"/>
       </a:ln>
     </p:spPr>
     <p:style>
       <a:lnRef idx="2">
         <a:schemeClr val="accent1"/>
       </a:lnRef>
       <a:fillRef idx="0">
         <a:schemeClr val="accent1"/>
       </a:fillRef>
       <a:effectRef idx="1">
         <a:schemeClr val="accent1"/>
       </a:effectRef>
       <a:fontRef idx="minor">
         <a:schemeClr val="tx1"/>
       </a:fontRef>
     </p:style>
   </p:cxnSp>

Straight segment jointed connector::

   <p:cxnSp>
     <p:nvCxnSpPr>
       <p:cNvPr id="9" name="Elbow Connector 8"/>
       <p:cNvCxnSpPr/>
       <p:nvPr/>
     </p:nvCxnSpPr>
     <p:spPr>
       <a:xfrm>
         <a:off x="950964" y="2124739"/>
         <a:ext cx="1257921" cy="415317"/>
       </a:xfrm>
       <a:prstGeom prst="bentConnector3">
         <a:avLst/>
       </a:prstGeom>
     </p:spPr>
     <p:style>
       <a:lnRef idx="2">
         <a:schemeClr val="accent1"/>
       </a:lnRef>
       <a:fillRef idx="0">
         <a:schemeClr val="accent1"/>
       </a:fillRef>
       <a:effectRef idx="1">
         <a:schemeClr val="accent1"/>
       </a:effectRef>
       <a:fontRef idx="minor">
         <a:schemeClr val="tx1"/>
       </a:fontRef>
     </p:style>
   </p:cxnSp>

Curved (S-like) connector::

   <p:cxnSp>
     <p:nvCxnSpPr>
       <p:cNvPr id="11" name="Curved Connector 10"/>
       <p:cNvCxnSpPr/>
       <p:nvPr/>
     </p:nvCxnSpPr>
     <p:spPr>
       <a:xfrm>
         <a:off x="950964" y="2925277"/>
         <a:ext cx="1257921" cy="619967"/>
       </a:xfrm>
       <a:prstGeom prst="curvedConnector3">
         <a:avLst/>
       </a:prstGeom>
     </p:spPr>
     <p:style>
       <a:lnRef idx="2">
         <a:schemeClr val="accent1"/>
       </a:lnRef>
       <a:fillRef idx="0">
         <a:schemeClr val="accent1"/>
       </a:fillRef>
       <a:effectRef idx="1">
         <a:schemeClr val="accent1"/>
       </a:effectRef>
       <a:fontRef idx="minor">
         <a:schemeClr val="tx1"/>
       </a:fontRef>
     </p:style>
   </p:cxnSp>

Freeform connector::

   <p:sp>
     <p:nvSpPr>
       <p:cNvPr id="12" name="Freeform 11"/>
       <p:cNvSpPr/>
       <p:nvPr/>
     </p:nvSpPr>
     <p:spPr>
       <a:xfrm>
         <a:off x="981058" y="4086962"/>
         <a:ext cx="1372277" cy="686176"/>
       </a:xfrm>
       <a:custGeom>
         <a:avLst/>
         <a:gdLst>
           <a:gd name="connsiteX0" fmla="*/ 0 w 1372277"/>
           <a:gd name="connsiteY0" fmla="*/ 0 h 686176"/>
           <a:gd name="connsiteX1" fmla="*/ 379182 w 1372277"/>
           <a:gd name="connsiteY1" fmla="*/ 306973 h 686176"/>
           <a:gd name="connsiteX2" fmla="*/ 944945 w 1372277"/>
           <a:gd name="connsiteY2" fmla="*/ 48152 h 686176"/>
         </a:gdLst>
         <a:ahLst/>
         <a:cxnLst>
           <a:cxn ang="0">
             <a:pos x="connsiteX0" y="connsiteY0"/>
           </a:cxn>
           <a:cxn ang="0">
             <a:pos x="connsiteX1" y="connsiteY1"/>
           </a:cxn>
           <a:cxn ang="0">
             <a:pos x="connsiteX2" y="connsiteY2"/>
           </a:cxn>
         </a:cxnLst>
         <a:rect l="l" t="t" r="r" b="b"/>
         <a:pathLst>
           <a:path w="1372277" h="686176">
             <a:moveTo>
               <a:pt x="0" y="0"/>
             </a:moveTo>
             <a:cubicBezTo>
               <a:pt x="110845" y="149474"/>
               <a:pt x="221691" y="298948"/>
               <a:pt x="379182" y="306973"/>
             </a:cubicBezTo>
             <a:cubicBezTo>
               <a:pt x="536673" y="314998"/>
               <a:pt x="811529" y="4012"/>
               <a:pt x="944945" y="48152"/>
             </a:cubicBezTo>
           </a:path>
         </a:pathLst>
       </a:custGeom>
     </p:spPr>
     <p:style>
       <a:lnRef idx="2">
         <a:schemeClr val="accent1"/>
       </a:lnRef>
       <a:fillRef idx="0">
         <a:schemeClr val="accent1"/>
       </a:fillRef>
       <a:effectRef idx="1">
         <a:schemeClr val="accent1"/>
       </a:effectRef>
       <a:fontRef idx="minor">
         <a:schemeClr val="tx1"/>
       </a:fontRef>
     </p:style>
     <p:txBody>
       <a:bodyPr rtlCol="0" anchor="ctr"/>
       <a:lstStyle/>
       <a:p>
         <a:pPr algn="ctr"/>
         <a:endParaRPr lang="en-US"/>
       </a:p>
     </p:txBody>
   </p:sp>

Completely free-form line::

   <p:sp>
     <p:nvSpPr>
       <p:cNvPr id="13" name="Freeform 12"/>
       <p:cNvSpPr/>
       <p:nvPr/>
     </p:nvSpPr>
     <p:spPr>
       <a:xfrm>
         <a:off x="1005133" y="5483390"/>
         <a:ext cx="1360239" cy="379203"/>
       </a:xfrm>
       <a:custGeom>
         <a:avLst/>
         <a:gdLst>
           <a:gd name="connsiteX0" fmla="*/ 0 w 1360239"/>
           <a:gd name="connsiteY0" fmla="*/ 0 h 379203"/>
           <a:gd name="connsiteX1" fmla="*/ 0 w 1360239"/>
           <a:gd name="connsiteY1" fmla="*/ 0 h 379203"/>
           <a:gd name="connsiteX2" fmla="*/ 96300 w 1360239"/>
           <a:gd name="connsiteY2" fmla="*/ 6020 h 379203"/>
           <a:gd name="connsiteX3" fmla="*/ 138431 w 1360239"/>
           <a:gd name="connsiteY3" fmla="*/ 18058 h 379203"/>
           <a:gd name="connsiteX4" fmla="*/ 222694 w 1360239"/>
           <a:gd name="connsiteY4" fmla="*/ 24077 h 379203"/>
           <a:gd name="connsiteX5" fmla="*/ 511594 w 1360239"/>
           <a:gd name="connsiteY5" fmla="*/ 24077 h 379203"/>
         </a:gdLst>
         <a:ahLst/>
         <a:cxnLst>
           <a:cxn ang="0">
             <a:pos x="connsiteX0" y="connsiteY0"/>
           </a:cxn>
           <a:cxn ang="0">
             <a:pos x="connsiteX1" y="connsiteY1"/>
           </a:cxn>
           <a:cxn ang="0">
             <a:pos x="connsiteX2" y="connsiteY2"/>
           </a:cxn>
           <a:cxn ang="0">
             <a:pos x="connsiteX3" y="connsiteY3"/>
           </a:cxn>
           <a:cxn ang="0">
             <a:pos x="connsiteX4" y="connsiteY4"/>
           </a:cxn>
           <a:cxn ang="0">
             <a:pos x="connsiteX5" y="connsiteY5"/>
           </a:cxn>
         </a:cxnLst>
         <a:rect l="l" t="t" r="r" b="b"/>
         <a:pathLst>
           <a:path w="1360239" h="379203">
             <a:moveTo>
               <a:pt x="0" y="0"/>
             </a:moveTo>
             <a:lnTo>
               <a:pt x="0" y="0"/>
             </a:lnTo>
             <a:cubicBezTo>
               <a:pt x="32100" y="2007"/>
               <a:pt x="64408" y="1860"/>
               <a:pt x="96300" y="6020"/>
             </a:cubicBezTo>
             <a:cubicBezTo>
               <a:pt x="110783" y="7909"/>
               <a:pt x="123972" y="15992"/>
               <a:pt x="138431" y="18058"/>
             </a:cubicBezTo>
             <a:cubicBezTo>
               <a:pt x="166307" y="22040"/>
               <a:pt x="194606" y="22071"/>
               <a:pt x="222694" y="24077"/>
             </a:cubicBezTo>
             <a:cubicBezTo>
               <a:pt x="333136" y="60893"/>
               <a:pt x="138800" y="-1634"/>
               <a:pt x="511594" y="24077"/>
             </a:cubicBezTo>
             <a:lnTo>
               <a:pt x="1360239" y="343089"/>
             </a:lnTo>
           </a:path>
         </a:pathLst>
       </a:custGeom>
     </p:spPr>
     <p:style>
       <a:lnRef idx="2">
         <a:schemeClr val="accent1"/>
       </a:lnRef>
       <a:fillRef idx="0">
         <a:schemeClr val="accent1"/>
       </a:fillRef>
       <a:effectRef idx="1">
         <a:schemeClr val="accent1"/>
       </a:effectRef>
       <a:fontRef idx="minor">
         <a:schemeClr val="tx1"/>
       </a:fontRef>
     </p:style>
     <p:txBody>
       <a:bodyPr rtlCol="0" anchor="ctr"/>
       <a:lstStyle/>
       <a:p>
         <a:pPr algn="ctr"/>
         <a:endParaRPr lang="en-US"/>
       </a:p>
     </p:txBody>
   </p:sp>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_GroupShape">
    <xsd:sequence>
      <xsd:element name="nvGrpSpPr" type="CT_GroupShapeNonVisual"/>
      <xsd:element name="grpSpPr"   type="a:CT_GroupShapeProperties"/>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="sp"           type="CT_Shape"/>
        <xsd:element name="grpSp"        type="CT_GroupShape"/>
        <xsd:element name="graphicFrame" type="CT_GraphicalObjectFrame"/>
        <xsd:element name="cxnSp"        type="CT_Connector"/>
        <xsd:element name="pic"          type="CT_Picture"/>
        <xsd:element name="contentPart"  type="CT_Rel"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Connector">
    <xsd:sequence>
      <xsd:element name="nvCxnSpPr" type="CT_ConnectorNonVisual"/>
      <xsd:element name="spPr"      type="a:CT_ShapeProperties"/>
      <xsd:element name="style"     type="a:CT_ShapeStyle"        minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ConnectorNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"      type="a:CT_NonVisualDrawingProps"/>
      <xsd:element name="cNvCxnSpPr" type="a:CT_NonVisualConnectorProperties"/>
      <xsd:element name="nvPr"       type="CT_ApplicationNonVisualDrawingProps"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"               type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group   ref="EG_Geometry"                                          minOccurs="0"/>
      <xsd:group   ref="EG_FillProperties"                                    minOccurs="0"/>
      <xsd:element name="ln"                 type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties"                                  minOccurs="0"/>
      <xsd:element name="scene3d"            type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"               type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"             type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_ShapeType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="line"/>
      <xsd:enumeration value="straightConnector1"/>
      <xsd:enumeration value="bentConnector2"/>
      <xsd:enumeration value="bentConnector3"/>
      <xsd:enumeration value="bentConnector4"/>
      <xsd:enumeration value="bentConnector5"/>
      <xsd:enumeration value="curvedConnector2"/>
      <xsd:enumeration value="curvedConnector3"/>
      <xsd:enumeration value="curvedConnector4"/>
      <xsd:enumeration value="curvedConnector5"/>
      ... other shape types removed ...
    </xsd:restriction>
  </xsd:simpleType>
