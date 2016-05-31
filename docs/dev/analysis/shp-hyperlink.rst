
Shape hyperlink
===============

In addition to hyperlinks in text, PowerPoint supports hyperlinks on shapes.
However, this feature is just a subset of the more general *click action*
functionality on shapes.

A run or shape can actually have two distinct mouse event behaviors, one for
(left) clicking and another for rolling over with the mouse. These are
independent; a run or shape can have one, the other, both, or neither. These
are the two hyperlink types reported by the Hyperlink.type attribute and
using enumeration values from MsoHyperlinkType.

A "true" hyperlink in PowerPoint is indicated by having no `action` attribute
on the `a:hlinkClick` element. The general hyperlink mechanism of storing
a URL in a mapped relationship is used by other verbs such as `hlinkfile` and
`program`, but these are considered distinct from hyperlinks for the sake of
this analysis. That distinction is reflected in the API.

See also: _Run.hyperlink, pptx.text._Hyperlink


Glossary
--------

**click action**
    The behavior to be executed when the user clicks on a shape. There are 14
    possible behaviors, such as navigate to a different slide, open a web
    page, and run a macro. The possible behaviors are described by the
    members of the `PP_ACTION_TYPE` enumeration.

**hover action**
    All click action behaviors can also be triggered by a mouse-over (hover)
    event.

**hyperlink**
    A hyperlink is a particular class of click action that roughly
    corresponds to opening a web page, although it can also be used to send
    an email. While a similar mechanism is used to specify other actions,
    such as open a file, this term here is reserved for the action of
    navigating to a web URL (including `mailto://`).

**action**
    The click or hover action is specified in the XML using a URL on the
    `ppaction://` protocol contained in the `action` attribute of the
    `a:hlinkClick` (or `a:hlinkHover`) element. A hyperlink action is implied
    when no `action` attribute is present.

**action verb**
    The specific action to be performed is contained in the *host* field of
    the `ppaction://` URL. For instance, `customshow` appears in
    `ppaction://customshow?id=0&return=true` to indicate
    a `PP_ACTION.NAMED_SLIDE_SHOW` action.

**OLE verb**
    The term *verb* also appears in this context to indicate an OLE verb such
    as `Open` or `Edit`. This is not to be confused with an `action verb`.


Candidate Protocol
------------------

.. highlight:: python

* Shape
  
  + `.click_action` - unconditionally returns an `ActionSetting` object,
    regardless of whether a click action is defined. Returns an
    `ActionSetting` object even when the shape type does not support a click
    action (such as a table).

* ActionSetting

  + `.action` - returns a member of the `PP_ACTION_TYPE` (`PP_ACTION`)
    enumeration

  + `.action_url` - returns the `ppaction://` URL as a string, in its
    entirety, or None if no action attribute is present. Maybe this should
    do XML character entity decoding.

  + `.action_verb` - returns the verb in the `ppaction://` URL, or None if no
    action URL is present. e.g. `'hlinksldjump'`

  + `.action_fields` - returns a dictionary containing the fields in the query
    string of the action URL.

  + `.hyperlink` - returns a Hyperlink object that represents the hyperlink
    defined for the shape. A Hyperlink object is always returned.

  + `.target_slide` - returns a Slide object when the action is a jump to
    another slide in the same presentation. This is the case when action is
    `FIRST_SLIDE`, `LAST_SLIDE`, `PREVIOUS_SLIDE`, `NEXT_SLIDE`, or
    `NAMED_SLIDE`.

* Hyperlink

  + `.address` - returns the URL contained in the relationship for this
    hyperlink.

  + `.screen_tip` - tool-tip text displayed on mouse rollover is slideshow
    mode. Put in the XML hooks for this but API call is second priority

Detect that a shape has a hyperlink::

    for shape in shapes:
        click_action = shape.click_action
        if click_action.action == PP_ACTION.HYPERLINK:
            print(click_action.hyperlink)


Add a hyperlink::

    p = shape.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = 'link to python-pptx @ GitHub'
    hlink = r.hyperlink
    hlink.address = 'https://github.com/scanny/python-pptx'

Delete a hyperlink::

    r.hyperlink = None

    # or -----------

    r.hyperlink.address = None  # empty string '' will do it too

A Hyperlink instance is lazy-created on first reference. The object persists
until garbage collected once created. The link XML is not written until
.address is specified. Setting ``hlink.address`` to None or '' causes the
hlink entry to be removed if present.


`PP_ACTION_TYPE` mapping logic
------------------------------

::

    # _ClickAction.action property

    hlinkClick = shape_elm.hlinkClick

    if hlinkClick is None:
        return PP_ACTION.NONE

    action_verb = hlinkClick.action_verb

    if action_verb == 'hlinkshowjump':
        relative_target = hlinkClick.action_fields['jump']
        return {
            'firstslide':      PP_ACTION.FIRST_SLIDE,
            'lastslide':       PP_ACTION.LAST_SLIDE,
            'lastslideviewed': PP_ACTION.LAST_SLIDE_VIEWED,
            'nextslide':       PP_ACTION.NEXT_SLIDE,
            'previousslide':   PP_ACTION.PREVIOUS_SLIDE,
            'endshow':         PP_ACTION.END_SHOW,
        }.relative_target

    return {
        None:           PP_ACTION.HYPERLINK,
        'hlinksldjump': PP_ACTION.NAMED_SLIDE,
        'hlinkpres':    PP_ACTION.PLAY,
        'hlinkfile':    PP_ACTION.OPEN_FILE,
        'customshow':   PP_ACTION.NAMED_SLIDE_SHOW,
        'ole':          PP_ACTION.OLE_VERB,
        'macro':        PP_ACTION.RUN_MACRO,
        'program':      PP_ACTION.RUN_PROGRAM,
    }.action_verb


PowerPoint® application behavior
--------------------------------

The general domain here is mouse event behaviors, with respect to a shape.
So far, the only two mouse events are (left) click and hover (mouse over).
These can trigger a variety of actions. I'm not sure if all actions can be
triggered by either event, but the XML appears to support it.

Action inventory
~~~~~~~~~~~~~~~~

The following behaviors can be triggered by a click:

* Jump to a relative slide in same presentation (first, last, next, previous,
  etc.).
* Jump to specific slide in same presentation (by slide index, perhaps title
  as fallback)
* Jump to a slide in different presentation (by slide index)
* End the slide show
* Jump to bookmark in Microsoft Word document
* Open an arbitrary file on the same computer
* Web link - Open a browser and navigate to a specified web page
* Run a macro
* Run an arbitrary program
* Execute an OLE action

In addition to performing one of these actions, zero, one, or both of two auxiliarly actions can be triggered by clicking:

* Play a sound
* Highlight the shape with a dashed line for a short time

Hyperlinkable shapes
~~~~~~~~~~~~~~~~~~~~

These shape types can have hyperlinks:

  + Autoshapes
  + Textbox
  + Picture
  + Connector (Line)
  + Chart

These shape types cannot:

  + Table
  + Group shape


UI procedures
-------------

Hyperlink autoshape to other slide by title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* Right-click > Hyperlink... (Cmd-K)
* Select Document panel
* Anchor: > Locate... > Slide Titles
* select slide by number and title, e.g. "2 

Add Anchor point in a document (or perhaps a slide)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* A hyperlink can link to a bookmark in a Word document
* It appears that maximum granularity in PowerPoint is to an entire slide
  (not to a range of text in a shape, for example)


MS API
------

Shape.ActionSettings(ppMouseClick | ppMouseOver)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The Shape object has an ActionSettings property, which is a collection of two
ActionSetting objects, one for click and the other for hover.
https://msdn.microsoft.com/EN-US/library/office/ff745656.aspx

ActionSetting
~~~~~~~~~~~~~

* Shape.ActionSettings(ppMouseClick | ppMouseOver) => ActionSetting

* ActionSetting.Action
  
  + one of: ppActionHyperlink, ppActionFirstSlide, ppActionPlay, or several
    others: https://msdn.microsoft.com/EN-US/library/office/ff744511.aspx

* ActionSetting.Hyperlink => Hyperlink

* Hyperlink members:
  
  + Address
  + SubAddress
  + TextToDisplay
  + ScreenTip
  + EmailSubject
  + Type (read-only, one of msoHyperlinkRange (run) or msoHyperlinkShape)


XML specimens
-------------

.. highlight:: xml

These are representative samples of shape XML showing the hyperlinks
associated the shape (as opposed to text contained by the shape).

* The `a:hlinkClick` element can be present or absent.

* Its parent, `p:cNvPr` is always present (is a required element).

* All of its attributes are optional, but an `a:hlinkClick` having no
  attributes has no meaning (or may trigger an error).

* Its `r:id` element is always present on click actions created by PowerPoint.
  Its value is an empty string when the action is first, last, next, previous,
  macro, and perhaps others.

* Adding a `highlightClick` attribute set True causes the shape to get
  a dashed line border for a short time when it is clicked.

* There are some more obscure attributes like "stop playing sound before
  navigating" that are available on `CT_Hyperlink`, perhaps meant for
  kiosk-style applications.

Summary
~~~~~~~

The action to perform on a mouse click is specified by the `action` attribute
of the `a:hlinkClick` element. Its value is a URL having the `ppaction://`
protocol, a verb, and an optional query string.

Some actions reference a relationship that specifies the target of the
action.

============= ======== =======================================================
verb          rId      behavior
============= ======== =======================================================
none          external Open a browser and navigate to URL in relationship
hlinkshowjump none     Jump to a relative slide in the same presentation
hlinksldjump  internal Jump to a specified slide in the same presentation
hlinkpres     external Jump to a specified slide in another presentation
hlinkfile     external Open an arbitrary file on the same computer
customshow    none     Start a custom slide show, option to return after
ole           none     Execute an OLE action (open, edit)
macro         none     Run an embedded VBA macro
program       external Execute an arbitrary program on same computer
============= ======== =======================================================

Jump to relative slide within presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**hlinkshowjump** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="7" name="Rounded Rectangle 6">
        <!-- this element does the needful -->
        <a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=firstslide"/>
      </p:cNvPr>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="1020781" y="1684235"/>
        <a:ext cx="1495562" cy="1775031"/>
      </a:xfrm>
      <a:prstGeom prst="roundRect">
        <a:avLst/>
      </a:prstGeom>
    </p:spPr>
    <p:txBody>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0"/>
          <a:t>Click to go to Foobar Slide</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0" smtClean="0"/>
      </a:p>
    </p:txBody>
  </p:sp>

* `jump` key can have value `firstslide`, `lastslide`, `previousslide`,
  `nextslide`, `lastslideviewed`, `endshow`.
* Note that `r:id` attribute is empty string; no relationship is required to
  determine target slide.

Jump to specific slide within presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**hlinksldjump** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="7" name="Rounded Rectangle 6">
        <a:hlinkClick r:id="rId2" action="ppaction://hlinksldjump"/>
      </p:cNvPr>
      ...
  </p:sp>

The corresponding `Relationship` element must be of type `slide`, be
internal, and point to the target slide in the package::

  <Relationship
    Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    Target="slide1.xml"/>

Jump to slide in another presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**hlinkpres** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="7" name="Rounded Rectangle 6">
        <a:hlinkClick r:id="rId3" action="ppaction://hlinkpres?slideindex=3&amp;slidetitle=Key Questions"/>
      </p:cNvPr>
      ...
  </p:sp>

The corresponding `Relationship` element must be of type `hyperlink`, be
*external*, and point to the target presentation with a URL (using the
`file://` protocol for a local file). The slide number and slide title are
provided in the `ppaction://` URL in the `a:hlinkClick` element::

  <Relationship
    Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="file://localhost/Users/scanny/Documents/checksec-prelim-analysis.pptx"
    TargetMode="External"/>

Web link (hyperlink)
~~~~~~~~~~~~~~~~~~~~

Note: The `action` attribute of `a:hlinkClick` has no value in this case.

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="4" name="Rounded Rectangle 3">
        <a:hlinkClick r:id="rId3"/>
      ...
  </p:sp>

The corresponding `Relationship` element must be of type `hyperlink`, be
*external*, and point to the target URL (using a web protocol).

The target is often a web URL, such as https://github/scanny/python-pptx,
including an optional anchor (e.g. #sub-heading suffix to jump mid-page). The
target can also be an email address, launching the local email client.
A mailto: URI is used in this case, with subject specifiable using
a '?subject=xyz' suffix.

An optional ScreenTip, a roll-over tool-tip sort of message, can also be
specified for a hyperlink. The XML schema does not limit its use to
hyperlinks, although the PowerPoint UI may not provide access to this field
in non-hyperlink cases.::

  <Relationship
    Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://www.google.com/"
    TargetMode="External"/>

Open an arbitrary file on the same computer
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**hlinkfile** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="7" name="Rounded Rectangle 6">
        <a:hlinkClick r:id="rId2" action="ppaction://hlinkfile"/>
        ...
  </p:sp>

* PowerPoint opens the file (after a warning dialog) using the default
  application for the file.

The corresponding `Relationship` element must be of type `hyperlink`, be
*external*, and point to the target file with a `file://` protocol URL::

  <Relationship
    Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="file:///C:\Install.log"
    TargetMode="External"/>

Run Custom SlideShow
~~~~~~~~~~~~~~~~~~~~

**customshow** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="4" name="Rounded Rectangle 3">
        <a:hlinkClick r:id="" action="ppaction://customshow?id=0&amp;return=true"/>
        ...
  </p:sp>

* The `return` query field determines whether focus returns to the current show
  after running the linked show. This field can be omitted, and defaults to
  `false`.

Execute an OLE action
~~~~~~~~~~~~~~~~~~~~~

**ole** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="9" name="Object 8">
        <a:hlinkClick r:id="" action="ppaction://ole?verb=0"/>
      </p:cNvPr>
    ...
  </p:sp>

This option is only available on an embedded (OLE) object. The verb field is
'0' for Edit and '1' for Open.

Run macro
~~~~~~~~~

**macro** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="4" name="Rounded Rectangle 3">
        <a:hlinkClick r:id="" action="ppaction://macro?name=Hello"/>
      </p:cNvPr>
    ...
  </p:sp>

Run a program
~~~~~~~~~~~~~

**program** action

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="4" name="Rounded Rectangle 3">
        <a:hlinkClick r:id="rId2" action="ppaction://program"/>
      ...
  </p:sp>

The corresponding `Relationship` element must be of type `hyperlink`, be
*external*, and point to the target application with a `file://` protocol
URL. ::

  <Relationship
    Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="file:///C:\Program%20Files%20(x86)\Vim\vim74\gvim.exe"
    TargetMode="External"/>

Play a sound
~~~~~~~~~~~~

Playing a sound is not a distinct action; rather, like highlighting, it is an
optional additional action to be performed on a click or hover event.

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="5" name="Rounded Rectangle 4">
        <a:hlinkClick r:id="" action="ppaction://..any..">
          <a:snd r:embed="rId3" name="applause.wav"/>
        </a:hlinkClick>
      ...
  </p:sp>

The corresponding `Relationship` element must be of type `audio`, be
internal, and point to a sound file embedded in the presentation::

  <Relationship
    Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
    Target="../media/audio1.wav"/>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_Shape">
    <xsd:sequence>
      <xsd:element name="nvSpPr" type="CT_ShapeNonVisual"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties"/>
      <xsd:element name="style"  type="a:CT_ShapeStyle"        minOccurs="0"/>
      <xsd:element name="txBody" type="a:CT_TextBody"          minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="useBgFill" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"   type="a:CT_NonVisualDrawingProps"/>
      <xsd:element name="cNvSpPr" type="a:CT_NonVisualDrawingShapeProps"/>
      <xsd:element name="nvPr"    type="CT_ApplicationNonVisualDrawingProps"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="hlinkClick" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="hlinkHover" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id"     type="ST_DrawingElementId" use="required"/>
    <xsd:attribute name="name"   type="xsd:string"          use="required"/>
    <xsd:attribute name="descr"  type="xsd:string"          default=""/>
    <xsd:attribute name="hidden" type="xsd:boolean"         default="false"/>
    <xsd:attribute name="title"  type="xsd:string"          default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_Hyperlink">
    <xsd:sequence>
      <xsd:element name="snd"    type="CT_EmbeddedWAVAudioFile"   minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute ref="r:id"/>
    <xsd:attribute name="invalidUrl"     type="xsd:string"  default=""/>
    <xsd:attribute name="action"         type="xsd:string"  default=""/>
    <xsd:attribute name="tgtFrame"       type="xsd:string"  default=""/>
    <xsd:attribute name="tooltip"        type="xsd:string"  default=""/>
    <xsd:attribute name="history"        type="xsd:boolean" default="true"/>
    <xsd:attribute name="highlightClick" type="xsd:boolean" default="false"/>
    <xsd:attribute name="endSnd"         type="xsd:boolean" default="false"/>
  </xsd:complexType>
