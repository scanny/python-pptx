
Presentation properties
=======================


Overview
--------

The presentation object has a few interesting properties. Right now I'm
interested in slide size.



Protocol
--------

::

    >>> assert isinstance(prs, pptx.parts.presentation.PresentationPart)
    >>> prs.slide_width
    9144000
    >>> prs.slide_height
    6858000
    >>> prs.slide_width = 11887200  # 13 inches
    >>> prs.slide_height = 6686550  # 7.3125 inches
    >>> prs.slide_width, prs.slide_height
    (11887200, 6686550)


XML specimens
-------------

.. highlight:: xml

Example presentation.xml contents::

  <p:presentation>
    <p:sldMasterIdLst>
      <p:sldMasterId id="2147483897" r:id="rId1"/>
    </p:sldMasterIdLst>
    <p:notesMasterIdLst>
      <p:notesMasterId r:id="rId153"/>
    </p:notesMasterIdLst>
    <p:sldIdLst>
      <p:sldId id="772" r:id="rId3"/>
      <p:sldId id="1244" r:id="rId4"/>
    </p:sldIdLst>
    <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
    <p:notesSz cx="6858000" cy="9296400"/>
    <p:defaultTextStyle>
      <a:defPPr>
        <a:defRPr lang="en-US"/>
      </a:defPPr>
      <a:lvl1pPr algn="l" rtl="0" fontAlgn="base">
        <a:spcBef>
          <a:spcPct val="0"/>
        </a:spcBef>
        <a:spcAft>
          <a:spcPct val="0"/>
        </a:spcAft>
        <a:defRPr kern="1200">
          <a:solidFill>
            <a:schemeClr val="tx1"/>
          </a:solidFill>
          <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
          <a:ea typeface="ＭＳ Ｐゴシック" pitchFamily="34" charset="-128"/>
          <a:cs typeface="+mn-cs"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:defaultTextStyle>
  </p:presentation>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:element name="presentation" type="CT_Presentation"/>

  <xsd:complexType name="CT_Presentation">
    <xsd:sequence>
      <xsd:element name="sldMasterIdLst"     type="CT_SlideMasterIdList"   minOccurs="0"/>
      <xsd:element name="notesMasterIdLst"   type="CT_NotesMasterIdList"   minOccurs="0"/>
      <xsd:element name="handoutMasterIdLst" type="CT_HandoutMasterIdList" minOccurs="0"/>
      <xsd:element name="sldIdLst"           type="CT_SlideIdList"         minOccurs="0"/>
      <xsd:element name="sldSz"              type="CT_SlideSize"           minOccurs="0"/>
      <xsd:element name="notesSz"            type="a:CT_PositiveSize2D"/>
      <xsd:element name="smartTags"          type="CT_SmartTags"           minOccurs="0"/>
      <xsd:element name="embeddedFontLst"    type="CT_EmbeddedFontList"    minOccurs="0"/>
      <xsd:element name="custShowLst"        type="CT_CustomShowList"      minOccurs="0"/>
      <xsd:element name="photoAlbum"         type="CT_PhotoAlbum"          minOccurs="0"/>
      <xsd:element name="custDataLst"        type="CT_CustomerDataList"    minOccurs="0"/>
      <xsd:element name="kinsoku"            type="CT_Kinsoku"             minOccurs="0"/>
      <xsd:element name="defaultTextStyle"   type="a:CT_TextListStyle"     minOccurs="0"/>
      <xsd:element name="modifyVerifier"     type="CT_ModifyVerifier"      minOccurs="0"/>
      <xsd:element name="extLst"             type="CT_ExtensionList"       minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="serverZoom"               type="a:ST_Percentage"       default="50%"/>
    <xsd:attribute name="firstSlideNum"            type="xsd:int"               default="1"/>
    <xsd:attribute name="showSpecialPlsOnTitleSld" type="xsd:boolean"           default="true"/>
    <xsd:attribute name="rtl"                      type="xsd:boolean"           default="false"/>
    <xsd:attribute name="removePersonalInfoOnSave" type="xsd:boolean"           default="false"/>
    <xsd:attribute name="compatMode"               type="xsd:boolean"           default="false"/>
    <xsd:attribute name="strictFirstAndLastChars"  type="xsd:boolean"           default="true"/>
    <xsd:attribute name="embedTrueTypeFonts"       type="xsd:boolean"           default="false"/>
    <xsd:attribute name="saveSubsetFonts"          type="xsd:boolean"           default="false"/>
    <xsd:attribute name="autoCompressPictures"     type="xsd:boolean"           default="true"/>
    <xsd:attribute name="bookmarkIdSeed"           type="ST_BookmarkIdSeed"     default="1"/>
    <xsd:attribute name="conformance"              type="s:ST_ConformanceClass"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideSize">
    <xsd:attribute name="cx"   type="ST_SlideSizeCoordinate" use="required"/>
    <xsd:attribute name="cy"   type="ST_SlideSizeCoordinate" use="required"/>
    <xsd:attribute name="type" type="ST_SlideSizeType" default="custom"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_SlideSizeType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="screen4x3"/>
      <xsd:enumeration value="letter"/>
      <xsd:enumeration value="A4"/>
      <xsd:enumeration value="35mm"/>
      <xsd:enumeration value="overhead"/>
      <xsd:enumeration value="banner"/>
      <xsd:enumeration value="custom"/>
      <xsd:enumeration value="ledger"/>
      <xsd:enumeration value="A3"/>
      <xsd:enumeration value="B4ISO"/>
      <xsd:enumeration value="B5ISO"/>
      <xsd:enumeration value="B4JIS"/>
      <xsd:enumeration value="B5JIS"/>
      <xsd:enumeration value="hagakiCard"/>
      <xsd:enumeration value="screen16x9"/>
      <xsd:enumeration value="screen16x10"/>
    </xsd:restriction>
  </xsd:simpleType>
