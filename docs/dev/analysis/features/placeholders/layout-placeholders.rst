
Layout Placeholders
===================


Candidate protocol
------------------

::

    >>> slide_layout = prs.slide_layouts[0]
    >>> slide_layout
    <pptx.parts.slidelayout.SlideLayout object at 0x10a5d42d0>
    >>> slide_layout.shapes
    <pptx.parts.slidelayout._LayoutShapeTree object at 0x104e60000>
    >>> slide_layout.shapes[0]
    <pptx.shapes.placeholder.LayoutPlaceholder object at 0x104e60020>
    >>> layout_placeholders = slide_layout.placeholders
    >>> layout_placeholders
    <pptx.shapes.placeholder.LayoutPlaceholders object at 0x104e60040>
    >>> len(layout_placeholders)
    5
    >>> layout_placeholders[1].type
    'body'
    >>> layout_placeholders[2].idx
    10
    >>> layout_placeholders.get(idx=0)
    <pptx.shapes.placeholder.MasterPlaceholder object at 0x104e60020>
    >>> layout_placeholders.get(idx=666)
    None
    >>> layout_placeholders[1]._sp.spPr.xfrm
    None
    >>> layout_placeholders[1].width  # should inherit from master placehldr
    8229600


Example XML
-----------

Slide layout placeholder::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Content Placeholder 2"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph idx="1"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr lvl="0"/>
          <a:r>
            <a:rPr lang="en-US" smtClean="0"/>
            <a:t>Click to edit Master text styles</a:t>
          </a:r>
        </a:p>
        <a:p>
          ... and others through lvl="4", five total
        </a:p>
      </p:txBody>
    </p:sp>


Behaviors
---------

* Hypothesis:

  + Layout inherits properties from master based only on type (e.g. body,
    dt), with no bearing on idx value.
  + Slide inherits from layout strictly on idx value.

  + Need to experiment to see if position and size are inherited on body
    placeholder by type or if idx has any effect

    - expanded/repackaged, change body ph idx on layout and see if still
      inherits position.
    - result: slide layout still appears with the proper position and size,
      however the slide placeholder appears at (0, 0) and apparently with
      size of (0, 0). 

  + What behavior is exhibited when a slide layout contains two body
    placeholders, neither of which has a matching idx value and neither of
    which has a specified size or position?

    - result: both placeholders are present, directly superimposed upon each
      other. They both inherit position and fill.


Layout placeholder inheritance
------------------------------

Objective: determine layout placeholder inheritee for each ph type

Observations:

Layout placeholder with *lyt-ph-type* inherits color from master placeholder
with *mst-ph-type*, noting idx value match.

===========  ===========  ===================================================
lyt-ph-type  mst-ph-type  notes
===========  ===========  ===================================================
ctrTitle     title        title layout - idx value matches (None, => 0)
subTitle     body         title layout - idx value matches (1)
dt           dt           title layout - idx 10 != 2
ftr          ftr          title layout - idx 11 != 3
sldNum       sldNum       title layout - idx 12 != 4
title        title        bullet layout - idx value matches (None, => 0)
None (obj)   body         bullet layout - idx value matches (1)
body         body         sect hdr - idx value matches (1)
None (obj)   body         two content - idx 2 != 1
body         body         comparison - idx 3 != 1
pic          body         picture - idx value matches (1)
chart        body         manual - idx 9 != 1
clipArt      body         manual - idx 9 != 1
dgm          body         manual - idx 9 != 1
media        body         manual - idx 9 != 1
tbl          body         manual - idx 9 != 1

hdr          repair err   valid only in Notes and Handout Slide
sldImg       repair err   valid only in Notes and Handout Slide
===========  ===========  ===================================================
