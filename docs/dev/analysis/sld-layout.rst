
Slide Layout
============

A slide layout acts as an property inheritance base for zero or more slides.
This provides a certain amount of separation between formatting and content
and contributes to visual consistency across the slides of a presentation.

Remove Layout
-------------

Preliminary notes
~~~~~~~~~~~~~~~~~

* Images on a removed slide layout appear to disappear automatically (without explicitly
  dropping their relationship from the slide layout before removing it).

* A "used" layout cannot be removed. Being "used" in this context means the presentation
  contains one or more slides *based* on that layout. The layout provides inherited
  placeholders, styling, and perhaps other background objects and will cause a repair
  error if missing.


Protocol
~~~~~~~~

The default slide layout collection (the one belonging to the first slide master) is
accessible directly from the presentation object::

  >>> from pptx import Presentation
  >>> prs = Presentation()
  >>> slide_layouts = prs.slide_layouts

Remove an (unused) slide layout::

  >>> slide_layouts.remove(slide_layouts[3])

Identify unused slide layouts::

  >>> [layout for layout in slide_layouts if not layout.used_by_slides])
  [
      <pptx.slide.SlideLayout object at 0x..a>
      <pptx.slide.SlideLayout object at 0x..b>
      ...
  ]


Specimen XML
------------

Each slide master contains a list of its slide layouts. Each layout is uniquely
identified by a slide-layout id (although this is the only place it appears) and is
accessed by traversing the relationship-id (rId) in the `r:id` attribute::

  <p:sldMaster>
    <!-- ... -->
    <p:sldLayoutIdLst>
      <p:sldLayoutId id="2147483649" r:id="rId1"/>
      <p:sldLayoutId id="2147483650" r:id="rId2"/>
      <p:sldLayoutId id="2147483651" r:id="rId3"/>
      <p:sldLayoutId id="2147483652" r:id="rId4"/>
      <p:sldLayoutId id="2147483653" r:id="rId5"/>
      <p:sldLayoutId id="2147483654" r:id="rId6"/>
      <p:sldLayoutId id="2147483655" r:id="rId7"/>
      <p:sldLayoutId id="2147483656" r:id="rId8"/>
      <p:sldLayoutId id="2147483657" r:id="rId9"/>
      <p:sldLayoutId id="2147483658" r:id="rId10"/>
      <p:sldLayoutId id="2147483659" r:id="rId11"/>
    </p:sldLayoutIdLst>
    <!-- ... -->
  </p:sldMaster>

The layout used by a slide is specified (only) in the `.rels` file for that slide::

  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship
      Id="rId1"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
      Target="../slideLayouts/slideLayout1.xml"/>
  </Relationships>


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:element name="sldLayout" type="CT_SlideLayout"/>

  <xsd:complexType name="CT_SlideLayout">
    <xsd:sequence>
      <xsd:element name="cSld"       type="CT_CommonSlideData"/>
      <xsd:element name="clrMapOvr"  type="a:CT_ColorMappingOverride" minOccurs="0"/>
      <xsd:element name="transition" type="CT_SlideTransition"        minOccurs="0"/>
      <xsd:element name="timing"     type="CT_SlideTiming"            minOccurs="0"/>
      <xsd:element name="hf"         type="CT_HeaderFooter"           minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_ExtensionListModify"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="showMasterSp"     type="xsd:boolean"        default="true"/>
    <xsd:attribute name="showMasterPhAnim" type="xsd:boolean"        default="true"/>
    <xsd:attribute name="matchingName"     type="xsd:string"         default=""/>
    <xsd:attribute name="type"             type="ST_SlideLayoutType" default="cust"/>
    <xsd:attribute name="preserve"         type="xsd:boolean"        default="false"/>
    <xsd:attribute name="userDrawn"        type="xsd:boolean"        default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_CommonSlideData">
    <xsd:sequence>
      <xsd:element name="bg"          type="CT_Background"       minOccurs="0"/>
      <xsd:element name="spTree"      type="CT_GroupShape"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="controls"    type="CT_ControlList"      minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="name" type="xsd:string" use="optional" default=""/>
  </xsd:complexType>
