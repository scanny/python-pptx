
Shape position and size
=======================

Shapes of the following types can appear in the shape tree of a slide
(``<p:spTree>``) and each will require support for querying size and position.

* **sp** -- Auto Shape
* **pic** -- Picture
* **grpSp** -- Group Shape
* **graphicFrame** -- container for table and chart shapes
* **cxnSp** -- Connector (line)
* **contentPart** -- has no position, but should return None instead of raising
  an exception


Acceptance test
---------------

.. highlight:: cucumber

::

    shp-pos-and-size.feature

    Feature: Query and set shape position and size
      In order to manipulate shapes on an existing slide
      As an application developer
      I need to get the position and size of a shape

    Scenario: get position and size of existing shape
       Given a shape of known position and size
        When I get the position and size of the shape
        Then it matches the known position and size of the shape

    Scenario: change position and size of an existing shape
       Given a shape of known position and size
        When I change the position and size of the shape
         And I get the position and size of the shape
        Then it matches the new position and size of the shape


Unit tests
----------

::

    DescribeShape

    it has a position
    it has dimensions
    it can change its position
    it can change its dimensions


Candidate API
-------------

* Shape.left
* Shape.top
* Shape.width
* Shape.height


Protocol
--------

::

    >>> assert isinstance(shape, pptx.shapes.autoshape.Shape)
    >>> shape.left
    914400
    >>> shape.left = Inches(0.5)
    >>> shape.left
    457200


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

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_Transform2D"            minOccurs="0" maxOccurs="1"/>
      ...
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

  <!-- Supporting elements -->

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
