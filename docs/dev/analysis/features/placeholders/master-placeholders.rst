
Master Placeholders
===================


Candidate protocol
------------------

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


Example XML
-----------

Slide master placeholder::

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


Behaviors
---------

* Master placeholder specifies size, position, geometry, and text body
  properties, such as margins (inset) and autofit

* It appears a slide master permits at most five placeholders, at most one
  each of title, body, date, footer, and slide number

* A master placeholder is always a `<p:sp>` (autoshape) element. It does not
  allow the insertion of content, although text entered in the shape somehow
  becomes the prompt text for inheriting placeholders.

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

* GUI doesn't seem to provide a way to add additional "footer" placeholders

* Does the MS API include populated object placeholders in the Placeholders
  collection?
