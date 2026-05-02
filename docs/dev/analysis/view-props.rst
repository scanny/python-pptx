Presentation view-properties part (``ppt/viewProps.xml``)
=========================================================

:Issue: `#94 <https://github.com/scanny/python-pptx/issues/94>`_
:Status: MVP shipped — see :class:`pptx.presentation.ViewProps`.
:Last updated: 2026-05-02

Scope
-----

This analysis covers the ``/ppt/viewProps.xml`` part — the OOXML container
that records editor-view settings PowerPoint restores when the
presentation is re-opened. It is an OPC **part** in its own right (not an
element of ``presentation.xml``), related to the presentation part by a
``http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps``
relationship.

Related is the ``p:presentation/@firstSlideNum`` attribute on the
presentation part itself, which controls the *display* number of the
first slide. Although functionally part of the "open settings" grouping
the issue calls out, ``firstSlideNum`` lives on a different part.

OOXML grammar
-------------

From ``spec/ISO-IEC-29500-1/schemas/xsd/pml.xsd`` §``CT_ViewProperties``::

    <xsd:complexType name="CT_ViewProperties">
      <xsd:sequence minOccurs="0" maxOccurs="1">
        <xsd:element name="normalViewPr"    type="CT_NormalViewProperties"         minOccurs="0"/>
        <xsd:element name="slideViewPr"     type="CT_SlideViewProperties"          minOccurs="0"/>
        <xsd:element name="outlineViewPr"   type="CT_OutlineViewProperties"        minOccurs="0"/>
        <xsd:element name="notesTextViewPr" type="CT_NotesTextViewProperties"      minOccurs="0"/>
        <xsd:element name="sorterViewPr"    type="CT_SlideSorterViewProperties"    minOccurs="0"/>
        <xsd:element name="notesViewPr"     type="CT_NotesViewProperties"          minOccurs="0"/>
        <xsd:element name="gridSpacing"     type="a:CT_PositiveSize2D"             minOccurs="0"/>
        <xsd:element name="extLst"          type="CT_ExtensionList"                minOccurs="0"/>
      </xsd:sequence>
      <xsd:attribute name="lastView"      type="ST_ViewType" default="sldView"/>
      <xsd:attribute name="showComments"  type="xsd:boolean" default="true"/>
    </xsd:complexType>

Key attributes and children:

* ``@lastView`` — which editor view PowerPoint restores when the deck is
  re-opened. Enumerated: ``sldView``, ``sldMasterView``, ``notesView``,
  ``handoutView``, ``notesMasterView``, ``outlineView``, ``sldSorterView``,
  ``sldThumbnailView``.
* ``@showComments`` — whether the comments pane is visible.
* ``p:normalViewPr`` — left/top pane-split positions and a handful of UI
  toggles (``showOutlineIcons``, ``snapVertSplitter``, ``vertBarState``,
  ``horzBarState``, ``preferSingleView``). Does **not** carry a zoom.
* ``p:slideViewPr`` / ``p:notesViewPr`` — wrap a ``p:cSldViewPr`` which
  wraps a ``p:cViewPr``. The zoom lives on ``p:cViewPr/p:scale``.
* ``p:outlineViewPr`` / ``p:sorterViewPr`` — wrap ``p:cViewPr`` directly.
* ``p:sorterViewPr/@showFormatting`` — whether the sorter renders text
  formatting in its slide thumbnails.
* ``p:cViewPr`` — contains ``p:scale`` (``a:CT_Scale2D`` with ``a:sx`` /
  ``a:sy`` n/d ratio pairs) and ``p:origin`` (scroll position).

The "zoom" attribute the issue text mentions is *not* a plain attribute
— it is the ``a:sx n="…" d="…"`` ratio under ``p:cViewPr/p:scale``. A
zoom of 124 % is written as ``<a:sx n="124" d="100"/>``.

First-slide-on-open: ``p:presentation/@firstSlideNum``
------------------------------------------------------

Default ``1``. Setting a non-default value causes PowerPoint to display
``firstSlideNum`` as the slide number of the first slide (and number
successive slides from there) in slide-number placeholders. It does not
reorder slides or change which slide is initially shown; the slide shown
first is always the first slide in ``sldIdLst``.

What PowerPoint actually writes
-------------------------------

A snapshot of ``ppt/viewProps.xml`` from a blank deck saved by
PowerPoint-for-Mac v16.92 looks like::

    <p:viewPr xmlns:a="…/drawingml/2006/main"
              xmlns:r="…/officeDocument/2006/relationships"
              xmlns:p="…/presentationml/2006/main"
              lastView="sldThumbnailView">
      <p:normalViewPr>
        <p:restoredLeft sz="15620"/>
        <p:restoredTop sz="94660"/>
      </p:normalViewPr>
      <p:slideViewPr>
        <p:cSldViewPr snapToGrid="0" snapToObjects="1">
          <p:cViewPr varScale="1">
            <p:scale><a:sx n="124" d="100"/><a:sy n="124" d="100"/></p:scale>
            <p:origin x="-1512" y="-112"/>
          </p:cViewPr>
          <p:guideLst>
            <p:guide orient="horz" pos="2160"/>
            <p:guide pos="2880"/>
          </p:guideLst>
        </p:cSldViewPr>
      </p:slideViewPr>
      <p:notesTextViewPr>
        <p:cViewPr>
          <p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale>
          <p:origin x="0" y="0"/>
        </p:cViewPr>
      </p:notesTextViewPr>
      <p:gridSpacing cx="76200" cy="76200"/>
    </p:viewPr>

python-pptx's own default template ships the same structure.

MVP surface
-----------

The MVP exposes six read/write properties on
:class:`pptx.presentation.ViewProps` (plus an escape-hatch ``element``
property for raw-XML access):

===========================  ==============================================  ====================
Python API                    OOXML source                                    Default when absent
===========================  ==============================================  ====================
``view_type``                 ``p:viewPr/@lastView``                          ``PP_VIEW_TYPE.NORMAL``
``show_comments``             ``p:viewPr/@showComments``                      ``True``
``show_formatting``           ``p:sorterViewPr/@showFormatting``              ``True``
``slide_view_zoom``           ``p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale``  ``1.0``
``notes_view_zoom``           ``p:notesViewPr/p:cSldViewPr/p:cViewPr/p:scale``  ``1.0``
``outline_view_zoom``         ``p:outlineViewPr/p:cViewPr/p:scale``            ``1.0``
``sorter_view_zoom``          ``p:sorterViewPr/p:cViewPr/p:scale``             ``1.0``
===========================  ==============================================  ====================

Plus :attr:`.Presentation.first_slide_num` which wraps
``p:presentation/@firstSlideNum``.

Zoom values are ``float`` ratios (1.0 = 100 %), encoded on disk as
``<a:sx n="…" d="100000"/>`` for consistent precision across save cycles.

Out of scope for this MVP
-------------------------

Intentionally not modeled (yet):

* ``p:normalViewPr`` pane-split geometry and UI toggles (``restoredLeft``,
  ``restoredTop``, ``showOutlineIcons``, ``snapVertSplitter``,
  ``vertBarState``, ``horzBarState``, ``preferSingleView``).
* ``p:cSldViewPr`` slide-view toggles (``snapToGrid``, ``snapToObjects``,
  ``showGuides``) and the ``p:guideLst`` list of editor guides.
* ``p:notesTextViewPr`` zoom (the notes-text view uses ``p:cViewPr``
  *directly*, not under ``p:cSldViewPr``, which would require a separate
  pair of accessors; the notes-page zoom is already covered by
  ``notes_view_zoom``).
* ``p:outlineViewPr/p:sldLst`` outline-collapse state.
* ``p:origin`` scroll-position.
* ``p:gridSpacing`` grid-line spacing.
* Every attribute on ``p:viewPr/p:extLst`` (2010+ extensions).

These round-trip verbatim — :attr:`ViewProps.element` is the escape
hatch for callers who need to read or write them today.

Related parts not covered by this issue
---------------------------------------

A few related OPC parts carry *other* presentation-level "open settings"
and are out of scope for issue #94:

* ``ppt/presProps.xml`` (``CT_PresentationProperties``) — print
  properties, show properties (loop / narration / timings), pen colour,
  color MRU.
* ``ppt/tableStyles.xml`` — table-style definitions (aesthetic, not open
  settings).
* ``docProps/app.xml`` — already modelled by
  :class:`.ExtendedPropertiesPart`.

These may be surfaced in follow-up issues.

Test coverage
-------------

* ``tests/oxml/test_viewprops.py`` — oxml-level round-tripping of
  ``CT_ViewProperties`` / ``CT_Scale2D`` / ``CT_SlideSorterViewProperties``.
* ``tests/parts/test_viewprops.py`` — :class:`.ViewPropsPart` default
  construction, content-type, partname.
* ``tests/test_presentation_view_props.py`` — end-to-end accessor tests
  (read the default-template values, write then reopen).
* ``features/prs-view-settings.feature`` — Gherkin round-trip through
  ``Presentation.save``.
