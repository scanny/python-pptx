
Slide Placeholders
==================

.. toctree::
   :titlesonly:

   picture-placeholder
   table-placeholder
   placeholders-in-new-slide


Candidate protocol
------------------

::

    >>> slide = prs.slides[0]
    >>> slide.shapes
    <pptx.shapes.shapetree.SlideShapes object at 0x104e60000>
    >>> slide.shapes[0]
    <pptx.shapes.placeholder.ShapePlaceholder object at 0x104e60020>
    >>> slide_placeholders = slide.placeholders
    >>> slide_placeholders
    <pptx.shapes.shapetree.SlidePlaceholders object at 0x104e60040>
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


Example XML
-----------

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
* placeholder has no `<p:spPr>` child elements, implying both that:

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


Behaviors
---------

* A placeholder is accessed through a slide-type object's shape collection,
  e.g. ``slide.shapes.placeholders``. The contents of the placeholders
  collection is a subset of the shapes in the shape collection.

* The title placeholder, if it exists, always appears first in the
  placeholders collection.

* Placeholders can only be top-level shapes, they cannot be nested in a group
  shape.

* A slide placeholder may be either an `<p:sp>` (autoshape) element or
  a `<p:pic>` or `<p:graphicFrame>` element. In either case, its relationship
  to its layout placeholder is preserved.

* Slide inherits from layout strictly on idx value.

* The placeholder with idx="0" is the title placeholder. "0" is the default
  for the *idx* attribute on <p:ph>, so one with no idx attribute is the
  title placeholder.

* The document order of placeholders signifies their z-order and has no
  bearing on their index order. :attr:`_ShapeCollection.placeholders` contains
  the placeholders in order of *idx* value, which means the title placeholder
  always appears first in the sequence, if it is present.

Automatic naming
~~~~~~~~~~~~~~~~

* Most placeholders are automatically named '{root_name} Placeholder {num}'
  where *root_name* is something like ``Chart`` and num is a positive integer.
  A typical example is ``Table Placeholder 3``.

* A placeholder with vertical orientation, i.e. ``<p:ph orient="vert">``, is
  prefixed with ``'Vertical '``, e.g. ``Vertical Text Placeholder 2``.

* The word ``'Placeholder'`` is omitted when the type is 'title', 'ctrTitle',
  or 'subTitle'.

On slide creation
~~~~~~~~~~~~~~~~~

When a new slide is added to a presentation, empty placeholders are added to
it based on the placeholders present in the specified slide layout.

Only content placeholders are created automatically. Special placeholders
must be created individually with additional calls. Special placeholders
include header, footer, page numbers, date, and slide number; and perhaps
others. In this respect, it seems like special placeholders on a slide layout
are simply default settings for those placeholders in case they are added
later.

During slide life-cycle
~~~~~~~~~~~~~~~~~~~~~~~

* Placeholders can be deleted from slides and can be restored by calling
  methods like AddTitle on the shape collection. AddTitle raises if there's
  already a title placeholder, so need to call .hasTitle() beforehand.


Experimental findings -- Applying layouts
-----------------------------------------

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
* Content of a placeholder shape is retained and displayed, even when the
  slide layout is changed to one without a matching layout placeholder.
* The behavior when placeholders are added to a slide layout (from the slide
  master) may also be worth characterizing.

  + ... show master placeholder ...
  + ... add (arbitrary) placeholder ...
