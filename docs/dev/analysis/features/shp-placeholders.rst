
Placeholders
============

Overview
--------

A *placeholder* is a shape that participates in a property inheritance
hierarchy, either as an inheritor, an inheritee, or both. Placeholder shapes
occur on masters, slide layouts, and slides. Placeholder shapes on masters
are inheritees only. Conversely placeholder shapes on slides are inheritors
only. Placeholders on slide layouts are both, a possible inheritor from a
slide master placeholder and an inheritee to placeholders on slides linked to
that layout.

From a user perspective, a *placeholder* acts as a pre-formatted container
into which content can be inserted.

A layout inherits from its master differently than how a slide inherits from
its layout. A layout placeholder inherits from the master placeholder sharing
the same type. A slide placeholder inherits from the layout placeholder
having the same `idx` value.


Candidate protocol -- MasterPlaceholders
----------------------------------------

::

    >>> slide_master = prs.slide_master
    >>> slide_master.shapes
    <pptx.shapes.slidemaster.MasterShapeTree object at 0x10a4df150>

    >>> slide_master.shapes[0]
    <pptx.shapes.placeholder._MasterPlaceholder object at 0x104e60290>
    >>> master_placeholders = slide_master.placeholders
    >>> master_placeholders
    <pptx.shapes.placeholder._MasterPlaceholders object at 0x104371290>
    >>> len(master_placeholders)
    5
    >>> master_placeholders[0].type
    'title'
    >>> master_placeholders[0].idx
    0
    >>> master_placeholders.get(type='body')
    <pptx.shapes.placeholder.MasterPlaceholder object at 0x104e60290>
    >>> master_placeholders.get(idx=666)
    None


Candidate protocol -- LayoutPlaceholders
----------------------------------------

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


Candidate protocol -- SlidePlaceholders
---------------------------------------

::

    >>> slide = prs.slides[0]
    >>> slide.shapes
    <pptx.shapes.shapetree.SlideShapeTree object at 0x104e60000>
    >>> slide.shapes[0]
    <pptx.shapes.placeholder.ShapePlaceholder object at 0x104e60020>
    >>> slide_placeholders = slide.placeholders
    >>> slide_placeholders
    <pptx.shapes.placeholder._SlidePlaceholders object at 0x104e60040>
    >>> len(slide_placeholders)
    2
    >>> slide_placeholders[1]
    <pptx.shapes.placeholder.ContentPlaceholder object at 0x104e60060>
    >>> slide_placeholders[1].type
    'body'
    >>> slide_placeholders.get(idx=1)
    AttributeError ...
    >>> slide_placeholders[1]._sp.spPr.xfrm
    None
    >>> slide_placeholders[1].width  # should inherit from layout placeholder
    8229600
    >>> table = slide_placeholders[1].insert_table(rows=2, cols=2)
    >>> len(slide_placeholders)
    2


ContentPlaceholder protocol
---------------------------

* ContentPlaceholder.insert_table()
* TablePlaceholder.insert_table()

* GraphicFramePlaceholder. ...
* PicturePlaceholder. ...

* add text
* all? shape query operations? such as get effective size and position?
* all? shape operations, such as set position and size

* [ ] object_placeholder inherits from slide_layout_placeholder where one with
      a matching `idx` value exists. Only `<p:pic>` and `<p:graphicFrame>`
      elements can be object placeholders.

Hypothesis: inheritance behaviors for an object placeholder are the same as
those for a slide_placeholder. Only object insertion behaviors are different.
There is some evidence object placeholder behaviors may be on a case-by-case
basis.


Shapes that can have a substituted placeholder
----------------------------------------------

Table, chart, smart art, picture, movie

* Picture
* GraphicFrame (Table, SmartArt and Chart)


Definitions
-----------

placeholder shape
    A shape on a slide that inherits from a layout placeholder.

layout placeholder
    a shorthand name for the placeholder shape on the slide layout from which
    a particular placeholder on a slide inherits shape properties

master placeholder
    the placeholder shape on the slide master which a layout placeholder
    inherits from, if any.


Inheritance behaviors
---------------------

A placeholder shape on a slide is initially little more than a reference to
its "parent" placeholder shape on the slide layout. If it is a placeholder
shape that can accept text, it contains a `<p:txBody>` element. Position,
size, and even geometry are inherited from the layout placeholder, which may
in turn inherit one or more of those properties from a master placeholder.


Substitution behaviors
----------------------

Content may be placed into a placeholder shape two ways, by *insertion* and
by *substitution*. Insertion is simply placing the text insertion point in
the placeholder and typing or pasting in text. Substitution occurs when an
object such as a table or picture is inserted into a placeholder by clicking
on a placeholder button.

An empty placeholder is always a `<p:sp>` (autoshape) element. When an object
such as a table is inserted into the placehoder by clicking on a placeholder
button, the `<p:sp>` element is replaced with the appropriate new shape
element, a table element in this case. The `<p:ph>` element is retained in
the new shape element and preserves the linkage to the layout placeholder
such that the 'empty' placeholder shape can be restored if the inserted
object is deleted.


Operations
----------

* clone on slide create
* query inherited property values
* substitution


Behavior
--------

* Content of a placeholder shape is retained and displayed, even when the
  slide layout is changed to one without a matching layout placeholder.

* The behavior when placeholders are added to a slide layout (from the slide
  master) may also be worth characterizing.

  + ... show master placeholder ...
  + ... add (arbitrary) placeholder ...


Sample XML
----------

.. highlight:: xml

Baseline textbox shape::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="2" name="TextBox 1"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="3016188" y="3273093"/>
          <a:ext cx="1133644" cy="369332"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="none" rtlCol="0">
          <a:spAutoFit/>
        </a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" dirty="0" smtClean="0"/>
            <a:t>Some text</a:t>
          </a:r>
          <a:endParaRPr lang="en-US" dirty="0"/>
        </a:p>
      </p:txBody>
    </p:sp>


Content placeholder::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="5" name="Content Placeholder 4"/>
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
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </p:txBody>
    </p:sp>


Notable differences:

* placeholder has `<a:spLocks>` element
* placeholder has `<p:ph>` element
* placeholder has no `<p:spPr>` child elements, this may imply both that:

  + all shape properties are initially inherited from the layout placeholder,
    including position, size, and geometry
  + any specific shape property value may be overridden by specifying it on
    the inheriting shape


Matching slide layout placeholder::

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


Matching slide master placeholder::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Text Placeholder 2"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="body" idx="1"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="457200" y="1600200"/>
          <a:ext cx="8229600" cy="4525963"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440"
                  bIns="45720" rtlCol="0">
          <a:normAutofit/>
        </a:bodyPr>
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


Note:

* master specifies size, position, and geometry
* master specifies text body properties, such as margins (inset) and autofit

# ----

A slide placeholder may be either an `<p:sp>` (autoshape) element or
a `<p:pic>` or `<p:graphicFrame>` element. In either case, its relationship
to its layout placeholder is preserved.


Experimental findings
---------------------

* Switching the layout of an empty title slide to the blank layout resulted
  in the placeholder shapes (title, subtitle) being removed.
* The same switch when the shapes had content (text), resulted in the shapes
  being preserved, complete with their `<p:ph>` element. Position and
  dimension values were added that preserve the height, width, top position
  but set the left position to zero.
* Restoring the original layout caused those position and dimension values to
  be removed (and "re-inherited").
* Applying a new (or the same) style to a slide appears to reset selected
  properties such that they are re-inherited from the new layout. Size and
  position are both reset. Background color and font, at least, are not
  reset.
* The "Reset Layout to Default Settings" option appears to reset all shape
  properties to inherited, without exception.


Other behaviors ...
~~~~~~~~~~~~~~~~~~~

* It appears a slide master permits at most five placeholders, at most one
  each of title, body, date, footer, and slide number

  + Hypothesis:

    - Layout inherits properties from master based only on type (e.g. body,
      dt), with no bearing on idx value.
    - Slide inherits from layout strictly on idx value.

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

* GUI doesn't seem to provide a way to add additional "footer" placeholders

* Does the MS API include populated object placeholders in the Placeholders
  collection?


Layout placeholder inheritance
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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


Resolved Design Questions
-------------------------

Q: What placeholder types should be members of `Slide.placeholders`? More
   specifically, should object shapes that were inserted into placeholders be
   members?

A: Start with including all placeholders, such that a placeholder's index in
   the collection is stable across object insertions and corresponds to the
   items visible on the slide. Reduce to only ShapePlaceholder instances
   (`<p:sp>` elements) later if some use case demonstrates a compelling
   reason not to include them.
