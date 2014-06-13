
Shape position and size
=======================

Shapes of the following types can appear in the shape tree of a slide
(``<p:spTree>``) and each will require support for querying size and position.

* **sp** -- Auto Shape *(completed)*
* **pic** -- Picture *(completed)*
* **graphicFrame** -- container for table and chart shapes
* **grpSp** -- Group Shape
* **cxnSp** -- Connector (line)
* **contentPart** -- has no position, but should return None instead of raising
  an exception


Protocol
--------

::

    >>> assert isinstance(shape, pptx.shapes.autoshape.Shape)
    >>> shape.left
    914400
    >>> shape.left = Inches(0.5)
    >>> shape.left
    457200


XML specimens
-------------

.. highlight:: xml

Here is a representative sample of shape XML showing the placement of the
position and size elements (in the <a:xfrm> element).

*Auto shape (rounded rectangle in this case)*::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Rounded Rectangle 2"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="760096" y="562720"/>
          <a:ext cx="2520824" cy="914400"/>
        </a:xfrm>
        <a:prstGeom prst="roundRect">
          <a:avLst>
            <a:gd name="adj" fmla="val 30346"/>
          </a:avLst>
        </a:prstGeom>
      </p:spPr>
      <p:style>
        <a:lnRef idx="1">
          <a:schemeClr val="accent1"/>
        </a:lnRef>
        <a:fillRef idx="3">
          <a:schemeClr val="accent1"/>
        </a:fillRef>
        <a:effectRef idx="2">
          <a:schemeClr val="accent1"/>
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="lt1"/>
        </a:fontRef>
      </p:style>
      <p:txBody>
        <a:bodyPr rtlCol="0" anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" dirty="0" smtClean="0"/>
            <a:t>This is text inside a rounded rectangle</a:t>
          </a:r>
          <a:endParaRPr lang="en-US" dirty="0"/>
        </a:p>
      </p:txBody>
    </p:sp>

*Example picture shape*::

    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="6" name="Picture 5" descr="python-logo.gif"/>
        <p:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </p:cNvPicPr>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip r:embed="rId2"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm>
          <a:off x="5580112" y="1988840"/>
          <a:ext cx="2679700" cy="901700"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <a:ln>
          <a:solidFill>
            <a:schemeClr val="bg1">
              <a:lumMod val="85000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </p:spPr>
    </p:pic>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_Shape">
    <xsd:sequence>
      <xsd:element name="nvSpPr" type="CT_ShapeNonVisual"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties"   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="style"  type="a:CT_ShapeStyle"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="txBody" type="a:CT_TextBody"          minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="useBgFill" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Picture">
    <xsd:sequence>
      <xsd:element name="nvPicPr"  type="CT_PictureNonVisual"     minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blipFill" type="a:CT_BlipFillProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="spPr"     type="a:CT_ShapeProperties"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="style"    type="a:CT_ShapeStyle"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"   type="CT_ExtensionListModify"  minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm" type="CT_Transform2D" minOccurs="0" maxOccurs="1"/>
      ...
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Transform2D">
    <xsd:sequence>
      <xsd:element name="off" type="CT_Point2D" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="rot" type="ST_Angle" use="optional" default="0"/>
    <xsd:attribute name="flipH" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="flipV" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Point2D">
    <xsd:attribute name="x" type="ST_Coordinate" use="required"/>
    <xsd:attribute name="y" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PositiveSize2D">
    <xsd:attribute name="cx" type="ST_PositiveCoordinate" use="required"/>
    <xsd:attribute name="cy" type="ST_PositiveCoordinate" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_PositiveCoordinate">
    <xsd:restriction base="xsd:long">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="27273042316900"/>
    </xsd:restriction>
  </xsd:simpleType>
