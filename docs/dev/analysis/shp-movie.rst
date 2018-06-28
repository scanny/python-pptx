
Movie
=====

PowerPoint allows a video to be inserted into a slide and run in slide show
mode. In the PowerPoint UI, a shape of this type is called a *movie*.

A movie is implemented as a variant form of a `p:pic` element, which makes it
closely related to a `Picture` object. This close relationship makes sense,
because a movie acts very much like an image when it is not playing and the
two shapes therefore share much of their behavior. When not playing, a movie
shape displays a static image, called a *poster frame*.


Candidate Protocol
------------------

A video can be added to a slide by specifying a video file, position, size,
and optional poster frame::

    >>> movie = shapes.add_movie(
    ...     movie_file, left, top, width, height, poster_frame_image,
    ...     mime_type
    ... )

    >>> movie
    <pptx.shape.picture.Movie object @ 0x12345678>

    >>> movie.shape_type
    MSO.MEDIA

    >>> movie.media_type
    PP_MEDIA_TYPE.MOVIE

    >>> movie.media_format
    <pptx.shape.picture.MediaFormat object @ 0x123456ab>


Animation/timing resources
--------------------------

* Working with animation (Open XML SDK)
  https://msdn.microsoft.com/en-us/library/office/gg278329.aspx

* AnimationBehaviors (Collection) Members (PowerPoint)
  https://msdn.microsoft.com/en-us/library/office/ff746028.aspx

* AnimationBehavior Members (PowerPoint)
  https://msdn.microsoft.com/en-us/library/office/ff746141.aspx

* <p:par> stands for *parallel*, and is one of the available time containers
  (like "parallel" timing, or executing at the same time).
  https://folk.uio.no/annembek/inf3210/how_2_SMIL.html

* https://technet.microsoft.com/en-au/library/gg278329.aspx

* http://openxmldeveloper.org/discussions/formats/f/15/t/6838.aspx


Possibly interesting resources
------------------------------

.. both of these formats work

* `Apache POI discussion <http://apache-poi.1045710.n5.nabble.com/Question-a
  bout-embedded-video-in-PPTX-files-td5718461.html>`_

* `Apache POI discussion on inserting video`_

.. _`Apache POI discussion on inserting video`:
   http://apache-poi.1045710.n5.nabble.com/Question-about-embedded-video-in-P
   PTX-files-td5718461.html


MS API
------

* Slide.Timeline
  https://msdn.microsoft.com/en-us/library/office/ff745382.aspx

* TimeLine Object
  https://msdn.microsoft.com/en-us/library/office/ff743965.aspx

* Shapes.AddMediaObject() (deprecated in PowerPoint 2013, see AddMediaObject2)
  https://msdn.microsoft.com/EN-US/library/office/ff745385.aspx

* Shapes.AddMediaObject2()
  https://msdn.microsoft.com/en-us/library/office/ff744080.aspx

* MediaFormat
  https://msdn.microsoft.com/en-us/library/office/ff745983.aspx


Enumerations
------------

PP_MEDIA_TYPE - https://msdn.microsoft.com/en-us/library/office/ff746008.aspx

+---------------------+----+--------+
| PP_MEDIA_TYPE.MOVIE | 3  | Movie  |
+---------------------+----+--------+
| PP_MEDIA_TYPE.SOUND | 2  | Sound  |
+---------------------+----+--------+
| PP_MEDIA_TYPE.OTHER | 1  | Others |
+---------------------+----+--------+
| PP_MEDIA_TYPE.MIXED | -2 |  Mixed |
+---------------------+----+--------+


Click Behavior
--------------

The video's "play on click" behavior in slideshow mode is implemented by the
use of a `<p:timing>` element in the `<p:sld>` element. No "play" button or
slider control appears when this element is not present.


Poster Frame
------------

A *poster frame* is the static image displayed in the video location when the
video is not playing. Each image that appears on the YouTube home page
representing a video is an example of a poster frame.

The poster frame is perhaps most frequently a frame from the video itself. In
some contexts, the first frame is used by default. The poster frame can be
undefined, or empty, and it can also be an unrelated image.

Some of the example videos for this feature get a poster frame upon
insertion; however at least one does not. In that case, a media "speaker"
icon (stretched to fit) is shown instead.


XML Semantics
-------------

* `id=` of p:cTn element just needs to be unique among `p:cTn` elements,
  apparently.


Example XML
-----------

.. highlight:: xml

Inserted MPEG-4 H.264 video::

  <!--slide{n}.xml-->

  <p:pic>
    <p:nvPicPr>
      <p:cNvPr id="6" name="video-filename.mp4">
        <a:hlinkClick r:id="" action="ppaction://media"/>
      </p:cNvPr>
      <p:cNvPicPr>
        <a:picLocks noChangeAspect="1"/>
      </p:cNvPicPr>
      <p:nvPr>
        <a:videoFile r:link="rId2"/>
        <p:extLst>
          <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
            <p14:media
                xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                r:embed="rId1"/>
          </p:ext>
        </p:extLst>
      </p:nvPr>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="rId4"/>
      <a:stretch>
        <a:fillRect/>
      </a:stretch>
    </p:blipFill>
    <p:spPr>
      <a:xfrm>
        <a:off x="5059279" y="876300"/>
        <a:ext cx="2390526" cy="5184274"/>
      </a:xfrm>
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
    </p:spPr>
  </p:pic>

Regular picture shape, for comparison::

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
    </p:spPr>
  </p:pic>


  <!--minimal (I think) p:video element-->

  <p:sld
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <p:cSld>
      <!-- ... -->
    </p:cSld>
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
    <p:timing>
      <p:tnLst>
        <p:par>
          <p:cTn xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:video>
                <p:cMediaNode vol="80000">
                  <p:cTn id="7" fill="hold" display="0">
                    <p:stCondLst>
                      <p:cond delay="indefinite"/>
                    </p:stCondLst>
                  </p:cTn>
                  <p:tgtEl>
                    <p:spTgt spid="3"/>
                  </p:tgtEl>
                </p:cMediaNode>
              </p:video>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
    </p:timing>
  </p:sld>


  <!--p:timing element for two videos in slide, as added by PowerPoint-->

  <p:timing>
    <p:tnLst>
      <p:par>
        <p:cTn xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" id="1" dur="indefinite" restart="never" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">
                <p:stCondLst>
                  <p:cond evt="onClick" delay="0">
                    <p:tgtEl>
                      <p:spTgt spid="3"/>
                    </p:tgtEl>
                  </p:cond>
                </p:stCondLst>
                <p:endSync evt="end" delay="0">
                  <p:rtn val="all"/>
                </p:endSync>
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst>
                        <p:cond delay="0"/>
                      </p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" fill="hold">
                            <p:stCondLst>
                              <p:cond delay="0"/>
                            </p:stCondLst>
                            <p:childTnLst>
                              <p:par>
                                <p:cTn id="5" presetID="2" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect">
                                  <p:stCondLst>
                                    <p:cond delay="0"/>
                                  </p:stCondLst>
                                  <p:childTnLst>
                                    <p:cmd type="call" cmd="togglePause">
                                      <p:cBhvr>
                                        <p:cTn id="6" dur="1" fill="hold"/>
                                        <p:tgtEl>
                                          <p:spTgt spid="3"/>
                                        </p:tgtEl>
                                      </p:cBhvr>
                                    </p:cmd>
                                  </p:childTnLst>
                                </p:cTn>
                              </p:par>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
              <p:nextCondLst>
                <p:cond evt="onClick" delay="0">
                  <p:tgtEl>
                    <p:spTgt spid="3"/>
                  </p:tgtEl>
                </p:cond>
              </p:nextCondLst>
            </p:seq>
            <p:video>
              <p:cMediaNode vol="80000">
                <p:cTn id="7" fill="hold" display="0">
                  <p:stCondLst>
                    <p:cond delay="indefinite"/>
                  </p:stCondLst>
                </p:cTn>
                <p:tgtEl>
                  <p:spTgt spid="3"/>
                </p:tgtEl>
              </p:cMediaNode>
            </p:video>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="8" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">
                <p:stCondLst>
                  <p:cond evt="onClick" delay="0">
                    <p:tgtEl>
                      <p:spTgt spid="4"/>
                    </p:tgtEl>
                  </p:cond>
                </p:stCondLst>
                <p:endSync evt="end" delay="0">
                  <p:rtn val="all"/>
                </p:endSync>
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="9" fill="hold">
                      <p:stCondLst>
                        <p:cond delay="0"/>
                      </p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="10" fill="hold">
                            <p:stCondLst>
                              <p:cond delay="0"/>
                            </p:stCondLst>
                            <p:childTnLst>
                              <p:par>
                                <p:cTn id="11" presetID="2" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect">
                                  <p:stCondLst>
                                    <p:cond delay="0"/>
                                  </p:stCondLst>
                                  <p:childTnLst>
                                    <p:cmd type="call" cmd="togglePause">
                                      <p:cBhvr>
                                        <p:cTn id="12" dur="1" fill="hold"/>
                                        <p:tgtEl>
                                          <p:spTgt spid="4"/>
                                        </p:tgtEl>
                                      </p:cBhvr>
                                    </p:cmd>
                                  </p:childTnLst>
                                </p:cTn>
                              </p:par>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
              <p:nextCondLst>
                <p:cond evt="onClick" delay="0">
                  <p:tgtEl>
                    <p:spTgt spid="4"/>
                  </p:tgtEl>
                </p:cond>
              </p:nextCondLst>
            </p:seq>
            <p:video>
              <p:cMediaNode vol="80000">
                <p:cTn id="13" fill="hold" display="0">
                  <p:stCondLst>
                    <p:cond delay="indefinite"/>
                  </p:stCondLst>
                </p:cTn>
                <p:tgtEl>
                  <p:spTgt spid="4"/>
                </p:tgtEl>
              </p:cMediaNode>
            </p:video>
          </p:childTnLst>
        </p:cTn>
      </p:par>
    </p:tnLst>
  </p:timing>



  <!--slide{n}.xml.rels-->

  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1"
        Type="http://schemas.microsoft.com/office/2007/relationships/media"
        Target="../media/media1.mp4"/>
    <Relationship Id="rId2"
        Type="http://sc.../officeDocument/2006/relationships/video"
        Target="../media/media1.mp4"/>
    <Relationship Id="rId3"
        Type="http://sc.../officeDocument/2006/relationships/slideLayout"
        Target="../slideLayouts/slideLayout1.xml"/>
    <!-- this one is the poster frame -->
    <Relationship Id="rId4"
        Type="http://sc.../officeDocument/2006/relationships/image"
        Target="../media/image1.png"/>
  </Relationships>


  <!--[Content_Types].xml-->

  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <!-- ... -->
    <Default Extension="mp4" ContentType="video/unknown"/>
    <Default Extension="png" ContentType="image/png"/>
    <Default Extension="jpeg" ContentType="image/jpeg"/>
    <!-- ... -->
  </Types>


p:video element
---------------

Provides playback controls.

http://openxmldeveloper.org/discussions/formats/f/15/p/1124/2842.aspx#2842


XML Semantics
-------------

* Extension DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230 is described as a
  `media extension`_. It appears to allow:

  + "cropping" the video period (set start and stop time markers)
  + provide for "fade-in"
  + allow for setting bookmarks in the video for fast jumps to a particular
    location

.. _media extension:
   https://msdn.microsoft.com/en-us/library/dd947021(v=office.12).aspx

* This and other extensions are documented in `this PDF <media_pdf>`_.

.. _media_pdf:
   http://interoperability.blob.core.windows.net/files/MS-PPTX/[MS-PPTX].pdf


Related Schema Definitions
--------------------------

.. highlight:: xml

The root element of a picture shape is a `p:pic (CT_Picture)` element::

  <xsd:complexType name="CT_Picture">
    <xsd:sequence>
      <xsd:element name="nvPicPr"  type="CT_PictureNonVisual"/>
      <xsd:element name="blipFill" type="a:CT_BlipFillProperties"/>
      <xsd:element name="spPr"     type="a:CT_ShapeProperties"/>
      <xsd:element name="style"    type="a:CT_ShapeStyle"        minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PictureNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"    type="a:CT_NonVisualDrawingProps"          minOccurs="1" maxOccurs="1"/>
      <xsd:element name="cNvPicPr" type="a:CT_NonVisualPictureProperties"     minOccurs="1" maxOccurs="1"/>
      <xsd:element name="nvPr"     type="CT_ApplicationNonVisualDrawingProps" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="hlinkClick" type="CT_Hyperlink"              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkHover" type="CT_Hyperlink"              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="id"     type="ST_DrawingElementId" use="required"/>
    <xsd:attribute name="name"   type="xsd:string"          use="required"/>
    <xsd:attribute name="descr"  type="xsd:string"          use="optional" default=""/>
    <xsd:attribute name="hidden" type="xsd:boolean"         use="optional" default="false"/>
    <xsd:attribute name="title"  type="xsd:string"          use="optional" default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualPictureProperties">
    <xsd:sequence>
      <xsd:element name="picLocks" type="CT_PictureLocking"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"   type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="preferRelativeResize" type="xsd:boolean" use="optional" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ApplicationNonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="ph"          type="CT_Placeholder"      minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="a:EG_Media"                              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="isPhoto"   type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="userDrawn" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:group name="EG_Media">
    <xsd:choice>
      <xsd:element name="audioCd"       type="CT_AudioCD"/>
      <xsd:element name="wavAudioFile"  type="CT_EmbeddedWAVAudioFile"/>
      <xsd:element name="audioFile"     type="CT_AudioFile"/>
      <xsd:element name="videoFile"     type="CT_VideoFile"/>
      <xsd:element name="quickTimeFile" type="CT_QuickTimeFile"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_BlipFillProperties">
    <xsd:sequence>
      <xsd:element name="blip"    type="CT_Blip"         minOccurs="0"/>
      <xsd:element name="srcRect" type="CT_RelativeRect" minOccurs="0"/>
      <xsd:group   ref="EG_FillModeProperties"           minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="dpi"          type="xsd:unsignedInt" use="optional"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"     use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Blip">
    <xsd:sequence>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="alphaBiLevel" type="CT_AlphaBiLevelEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaCeiling" type="CT_AlphaCeilingEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaFloor"   type="CT_AlphaFloorEffect"         minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaInv"     type="CT_AlphaInverseEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaMod"     type="CT_AlphaModulateEffect"      minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaModFix"  type="CT_AlphaModulateFixedEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaRepl"    type="CT_AlphaReplaceEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="biLevel"      type="CT_BiLevelEffect"            minOccurs="1" maxOccurs="1"/>
        <xsd:element name="blur"         type="CT_BlurEffect"               minOccurs="1" maxOccurs="1"/>
        <xsd:element name="clrChange"    type="CT_ColorChangeEffect"        minOccurs="1" maxOccurs="1"/>
        <xsd:element name="clrRepl"      type="CT_ColorReplaceEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="duotone"      type="CT_DuotoneEffect"            minOccurs="1" maxOccurs="1"/>
        <xsd:element name="fillOverlay"  type="CT_FillOverlayEffect"        minOccurs="1" maxOccurs="1"/>
        <xsd:element name="grayscl"      type="CT_GrayscaleEffect"          minOccurs="1" maxOccurs="1"/>
        <xsd:element name="hsl"          type="CT_HSLEffect"                minOccurs="1" maxOccurs="1"/>
        <xsd:element name="lum"          type="CT_LuminanceEffect"          minOccurs="1" maxOccurs="1"/>
        <xsd:element name="tint"         type="CT_TintEffect"               minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attributeGroup ref="AG_Blob"/>
    <xsd:attribute name="cstate" type="ST_BlipCompression" use="optional" default="none"/>
  </xsd:complexType>

  <xsd:attributeGroup name="AG_Blob">
    <xsd:attribute ref="r:embed" use="optional" default=""/>
    <xsd:attribute ref="r:link"  use="optional" default=""/>
  </xsd:attributeGroup>

  <xsd:group name="EG_FillModeProperties">
    <xsd:choice>
      <xsd:element name="tile"    type="CT_TileInfoProperties"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="stretch" type="CT_StretchInfoProperties" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_StretchInfoProperties">
    <xsd:sequence>
      <xsd:element name="fillRect" type="CT_RelativeRect" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_RelativeRect">
    <xsd:attribute name="l" type="ST_Percentage" use="optional" default="0%"/>
    <xsd:attribute name="t" type="ST_Percentage" use="optional" default="0%"/>
    <xsd:attribute name="r" type="ST_Percentage" use="optional" default="0%"/>
    <xsd:attribute name="b" type="ST_Percentage" use="optional" default="0%"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group ref="EG_Geometry"                                 minOccurs="0"/>
      <xsd:group ref="EG_FillProperties"                           minOccurs="0"/>
      <xsd:element name="ln"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group ref="EG_EffectProperties"                         minOccurs="0"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"    type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Transform2D">
    <xsd:sequence>
      <xsd:element name="off" type="CT_Point2D"        minOccurs="0"/>
      <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rot"   type="ST_Angle"    default="0"/>
    <xsd:attribute name="flipH" type="xsd:boolean" default="false"/>
    <xsd:attribute name="flipV" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Point2D">
    <xsd:attribute name="x" type="ST_Coordinate" use="required"/>
    <xsd:attribute name="y" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PositiveSize2D">
    <xsd:attribute name="cx" type="ST_PositiveCoordinate" use="required"/>
    <xsd:attribute name="cy" type="ST_PositiveCoordinate" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_Geometry">
    <xsd:choice>
      <xsd:element name="custGeom" type="CT_CustomGeometry2D"/>
      <xsd:element name="prstGeom" type="CT_PresetGeometry2D"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_FillProperties">
    <xsd:choice>
      <xsd:element name="noFill"    type="CT_NoFillProperties"/>
      <xsd:element name="solidFill" type="CT_SolidColorFillProperties"/>
      <xsd:element name="gradFill"  type="CT_GradientFillProperties"/>
      <xsd:element name="blipFill"  type="CT_BlipFillProperties"/>
      <xsd:element name="pattFill"  type="CT_PatternFillProperties"/>
      <xsd:element name="grpFill"   type="CT_GroupFillProperties"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_EffectProperties">
    <xsd:choice>
      <xsd:element name="effectLst" type="CT_EffectList"/>
      <xsd:element name="effectDag" type="CT_EffectContainer"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_ShapeStyle">
    <xsd:sequence>
      <xsd:element name="lnRef"     type="CT_StyleMatrixReference"/>
      <xsd:element name="fillRef"   type="CT_StyleMatrixReference"/>
      <xsd:element name="effectRef" type="CT_StyleMatrixReference"/>
      <xsd:element name="fontRef"   type="CT_FontReference"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:element name="notes" type="CT_NotesSlide"/>

  <xsd:complexType name="CT_NotesSlide">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="cSld"      type="CT_CommonSlideData"/>
      <xsd:element name="clrMapOvr" type="a:CT_ColorMappingOverride" minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionListModify"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="showMasterSp"     type="xsd:boolean" default="true"/>
    <xsd:attribute name="showMasterPhAnim" type="xsd:boolean" default="true"/>
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


`p:timing` related Schema Definitions
-------------------------------------

.. highlight:: xml

The `p:timing` element is a child of the `p:sld (CT_Slide)` element::

  <xsd:complexType name="CT_Slide">  <!-- denormalized -->
    <xsd:sequence minOccurs="1" maxOccurs="1">
      <xsd:element name="cSld"       type="CT_CommonSlideData"/>
      <xsd:element name="clrMapOvr"  type="a:CT_ColorMappingOverride" minOccurs="0"/>
      <xsd:element name="transition" type="CT_SlideTransition"        minOccurs="0"/>
      <xsd:element name="timing"     type="CT_SlideTiming"            minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_ExtensionListModify"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="showMasterSp"     type="xsd:boolean" use="optional" default="true"/>
    <xsd:attribute name="showMasterPhAnim" type="xsd:boolean" use="optional" default="true"/>
    <xsd:attribute name="show"             type="xsd:boolean" use="optional" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideTiming">
    <xsd:sequence>
      <xsd:element name="tnLst"  type="CT_TimeNodeList"        minOccurs="0"/>
      <xsd:element name="bldLst" type="CT_BuildList"           minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TimeNodeList">
    <xsd:choice minOccurs="1" maxOccurs="unbounded">
      <xsd:element name="par"        type="CT_TLTimeNodeParallel"/>
      <xsd:element name="seq"        type="CT_TLTimeNodeSequence"/>
      <xsd:element name="excl"       type="CT_TLTimeNodeExclusive"/>
      <xsd:element name="anim"       type="CT_TLAnimateBehavior"/>
      <xsd:element name="animClr"    type="CT_TLAnimateColorBehavior"/>
      <xsd:element name="animEffect" type="CT_TLAnimateEffectBehavior"/>
      <xsd:element name="animMotion" type="CT_TLAnimateMotionBehavior"/>
      <xsd:element name="animRot"    type="CT_TLAnimateRotationBehavior"/>
      <xsd:element name="animScale"  type="CT_TLAnimateScaleBehavior"/>
      <xsd:element name="cmd"        type="CT_TLCommandBehavior"/>
      <xsd:element name="set"        type="CT_TLSetBehavior"/>
      <xsd:element name="audio"      type="CT_TLMediaNodeAudio"/>
      <xsd:element name="video"      type="CT_TLMediaNodeVideo"/>
    </xsd:choice>
  </xsd:complexType>

  <xsd:complexType name="CT_TLTimeNodeParallel">
    <xsd:sequence>
      <xsd:element name="cTn" type="CT_TLCommonTimeNodeData"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TLTimeNodeSequence">
    <xsd:sequence>
      <xsd:element name="cTn"         type="CT_TLCommonTimeNodeData"/>
      <xsd:element name="prevCondLst" type="CT_TLTimeConditionList" minOccurs="0"/>
      <xsd:element name="nextCondLst" type="CT_TLTimeConditionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="concurrent" type="xsd:boolean"/>
    <xsd:attribute name="prevAc"     type="ST_TLPreviousActionType"/>
    <xsd:attribute name="nextAc"     type="ST_TLNextActionType"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TLCommonTimeNodeData">
    <xsd:sequence>
      <xsd:element name="stCondLst"  type="CT_TLTimeConditionList" minOccurs="0"/>
      <xsd:element name="endCondLst" type="CT_TLTimeConditionList" minOccurs="0"/>
      <xsd:element name="endSync"    type="CT_TLTimeCondition"     minOccurs="0"/>
      <xsd:element name="iterate"    type="CT_TLIterateData"       minOccurs="0"/>
      <xsd:element name="childTnLst" type="CT_TimeNodeList"        minOccurs="0"/>
      <xsd:element name="subTnLst"   type="CT_TimeNodeList"        minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id"            type="ST_TLTimeNodeID"/>
    <xsd:attribute name="presetID"      type="xsd:int"/>
    <xsd:attribute name="presetClass"   type="ST_TLTimeNodePresetClassType"/>
    <xsd:attribute name="presetSubtype" type="xsd:int"/>
    <xsd:attribute name="dur"           type="ST_TLTime"/>
    <xsd:attribute name="repeatCount"   type="ST_TLTime"                     default="1000"/>
    <xsd:attribute name="repeatDur"     type="ST_TLTime"/>
    <xsd:attribute name="spd"           type="a:ST_Percentage"               default="100%"/>
    <xsd:attribute name="accel"         type="a:ST_PositiveFixedPercentage"  default="0%"/>
    <xsd:attribute name="decel"         type="a:ST_PositiveFixedPercentage"  default="0%"/>
    <xsd:attribute name="autoRev"       type="xsd:boolean"                   default="false"/>
    <xsd:attribute name="restart"       type="ST_TLTimeNodeRestartType"/>
    <xsd:attribute name="fill"          type="ST_TLTimeNodeFillType"/>
    <xsd:attribute name="syncBehavior"  type="ST_TLTimeNodeSyncType"/>
    <xsd:attribute name="tmFilter"      type="xsd:string"/>
    <xsd:attribute name="evtFilter"     type="xsd:string"/>
    <xsd:attribute name="display"       type="xsd:boolean"/>
    <xsd:attribute name="masterRel"     type="ST_TLTimeNodeMasterRelation"/>
    <xsd:attribute name="bldLvl"        type="xsd:int"/>
    <xsd:attribute name="grpId"         type="xsd:unsignedInt"/>
    <xsd:attribute name="afterEffect"   type="xsd:boolean"/>
    <xsd:attribute name="nodeType"      type="ST_TLTimeNodeType"/>
    <xsd:attribute name="nodePh"        type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TLMediaNodeVideo">
    <xsd:sequence>
      <xsd:element name="cMediaNode" type="CT_TLCommonMediaNodeData"/>
    </xsd:sequence>
    <xsd:attribute name="fullScrn" type="xsd:boolean" default="false"/>
  </xsd:complexType>
