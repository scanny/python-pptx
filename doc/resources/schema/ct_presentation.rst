===================
``CT_Presentation``
===================

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_Presentation
   Spec Name    , Presentation
   Tag(s)       , p:presentation
   Namespace    , presentationml (pml.xsd)
   Schema Line  , 1038
   Spec Section , 19.2.1.26


Analysis
========

The ``<p:presentation>`` element is the root element for a presentation part.

The id values for <p:sldMasterId> elements in <p:sldMasterIdLst> start at a
strangely high number, id="2147483648" is the minimum stated in the schema.
It's possible this is so they don't conflict with bookmark ids (internal
hyperlinks). The range for those ids is up to exactly one less than that
number. 2147483647 also happens to be the maximum value of a 32-bit signed
int.

Id values for slides in <p:sldIdLst> start at 256.


attributes
^^^^^^^^^^

========================  ===  =====================  ============
name                      use  type                   default
========================  ===  =====================  ============
serverZoom                 ?   a:ST_Percentage        50%
firstSlideNum              ?   xsd:int                1
showSpecialPlsOnTitleSld   ?   xsd:boolean            true
rtl                        ?   xsd:boolean            false
removePersonalInfoOnSave   ?   xsd:boolean            false
compatMode                 ?   xsd:boolean            false
strictFirstAndLastChars    ?   xsd:boolean            true
embedTrueTypeFonts         ?   xsd:boolean            false
saveSubsetFonts            ?   xsd:boolean            false
autoCompressPictures       ?   xsd:boolean            true
bookmarkIdSeed             ?   ST_BookmarkIdSeed      1
conformance                ?   s:ST_ConformanceClass  transitional
========================  ===  =====================  ============



child elements
^^^^^^^^^^^^^^

==================  ===  ======================  ========
name                 #   type                    line
==================  ===  ======================  ========
sldMasterIdLst       ?   CT_SlideMasterIdList    894
notesMasterIdLst     ?   CT_NotesMasterIdList    906
handoutMasterIdLst   ?   CT_HandoutMasterIdList  918
sldIdLst             ?   CT_SlideIdList          877
sldSz                ?   CT_SlideSize
notesSz              ?   a:CT_PositiveSize2D
smartTags            ?   CT_SmartTags
embeddedFontLst      ?   CT_EmbeddedFontList
custShowLst          ?   CT_CustomShowList
photoAlbum           ?   CT_PhotoAlbum
custDataLst          ?   CT_CustomerDataList
kinsoku              \*  CT_Kinsoku
defaultTextStyle     ?   a:CT_TextListStyle
modifyVerifier       ?   CT_ModifyVerifier
extLst               ?   CT_ExtensionList
==================  ===  ======================  ========


Spec text
^^^^^^^^^

   This element specifies within it fundamental presentation-wide properties.


Schema excerpt
^^^^^^^^^^^^^^

::

  <xsd:complexType name="CT_Presentation">
    <xsd:sequence>
      <xsd:element name="sldMasterIdLst"     type="CT_SlideMasterIdList"   minOccurs="0" maxOccurs="1"/>
      <xsd:element name="notesMasterIdLst"   type="CT_NotesMasterIdList"   minOccurs="0" maxOccurs="1"/>
      <xsd:element name="handoutMasterIdLst" type="CT_HandoutMasterIdList" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sldIdLst"           type="CT_SlideIdList"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sldSz"              type="CT_SlideSize"           minOccurs="0" maxOccurs="1"/>
      <xsd:element name="notesSz"            type="a:CT_PositiveSize2D"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="smartTags"          type="CT_SmartTags"           minOccurs="0" maxOccurs="1"/>
      <xsd:element name="embeddedFontLst"    type="CT_EmbeddedFontList"    minOccurs="0" maxOccurs="1"/>
      <xsd:element name="custShowLst"        type="CT_CustomShowList"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="photoAlbum"         type="CT_PhotoAlbum"          minOccurs="0" maxOccurs="1"/>
      <xsd:element name="custDataLst"        type="CT_CustomerDataList"    minOccurs="0" maxOccurs="1"/>
      <xsd:element name="kinsoku"            type="CT_Kinsoku"             minOccurs="0"/>
      <xsd:element name="defaultTextStyle"   type="a:CT_TextListStyle"     minOccurs="0" maxOccurs="1"/>
      <xsd:element name="modifyVerifier"     type="CT_ModifyVerifier"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"             type="CT_ExtensionList"       minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="serverZoom"               type="a:ST_Percentage"   use="optional" default="50%"/>
    <xsd:attribute name="firstSlideNum"            type="xsd:int"           use="optional" default="1"/>
    <xsd:attribute name="showSpecialPlsOnTitleSld" type="xsd:boolean"       use="optional" default="true"/>
    <xsd:attribute name="rtl"                      type="xsd:boolean"       use="optional" default="false"/>
    <xsd:attribute name="removePersonalInfoOnSave" type="xsd:boolean"       use="optional" default="false"/>
    <xsd:attribute name="compatMode"               type="xsd:boolean"       use="optional" default="false"/>
    <xsd:attribute name="strictFirstAndLastChars"  type="xsd:boolean"       use="optional" default="true"/>
    <xsd:attribute name="embedTrueTypeFonts"       type="xsd:boolean"       use="optional" default="false"/>
    <xsd:attribute name="saveSubsetFonts"          type="xsd:boolean"       use="optional" default="false"/>
    <xsd:attribute name="autoCompressPictures"     type="xsd:boolean"       use="optional" default="true"/>
    <xsd:attribute name="bookmarkIdSeed"           type="ST_BookmarkIdSeed" use="optional" default="1"/>
    <xsd:attribute name="conformance"              type="s:ST_ConformanceClass"/>
  </xsd:complexType>

  <!-- Slide Master ID List -->
  
  <xsd:complexType name="CT_SlideMasterIdList">
    <xsd:sequence>
      <xsd:element name="sldMasterId" type="CT_SlideMasterIdListEntry" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideMasterIdListEntry">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="id" type="ST_SlideMasterId" use="optional"/>
    <xsd:attribute ref="r:id" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_SlideMasterId">
    <xsd:restriction base="xsd:unsignedInt">
      <xsd:minInclusive value="2147483648"/>
    </xsd:restriction>
  </xsd:simpleType>

  <!-- Slide ID List -->

  <xsd:complexType name="CT_SlideIdList">
    <xsd:sequence>
      <xsd:element name="sldId" type="CT_SlideIdListEntry" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideIdListEntry">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="id" type="ST_SlideId" use="required"/>
    <xsd:attribute ref="r:id" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_SlideId">
    <xsd:restriction base="xsd:unsignedInt">
      <xsd:minInclusive value="256"/>
      <xsd:maxExclusive value="2147483648"/>
    </xsd:restriction>
  </xsd:simpleType>

