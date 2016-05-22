=========================
Placeholders in new slide
=========================

Topic of inquiry
================

What placeholder shapes are added to a new slide when it is added to a
presentation?


Abstract
========

Slides were created based on a variety of slide layouts so their placeholder
elements could be inspected to reveal slide construction behaviors that depend
on the slide layout chosen. A placeholder shape appeared in each slide
corresponding to each of the placeholder shapes in the slide layout, excepting
those of footer, date, and slide number type. Placeholder shapes were
``<p:sp>`` elements for all placeholder types rather than matching the element
type of the shape they hold a place for. Each slide-side placeholder had a
``<p:ph>`` element that exactly matched the corresponding element in the slide
layout, with the exception that a *hasCustomPrompt* attribute, if present in
the slide layout, did not appear in the slide. Placeholders in the slide were
named and numbered differently. All placeholders in the slide layout had a
``<p:txBody>`` element, but only certain types had one in the slide. All
placeholders were locked against grouping.


Procedure
=========

This procedure was run using PowerPoint for Mac 2011 running on an Intel-based
Mac Pro running Mountain Lion 10.8.2.

1. Create a new presentation based on the "White" theme.

#. Add a new slide layout and add placeholders of the following types. The
   new slide layout is (automatically) named *Custom Layout*.

   * Chart
   * Table
   * SmartArt Graphic
   * Media
   * Clip Art

#. Add a new slide for each of the following layouts. Do not enter any content
   into the placeholders.

   * Title Slide (added automatically on File > New)
   * Title and Content
   * Two Content
   * Comparison
   * Content with Caption
   * Picture with Caption
   * Vertical Title and Text
   * Custom Layout

#. Save and unpack the presentation. This command works on OS X: ``unzip
   white.pptx -d white``

#. Inspect the slide.xml files and compare placeholders with those present in
   their corresponding slideLayout.xml file.


Observations
============

Typical XML for a placeholder::

   <p:sp>
     <p:nvSpPr>
       <p:cNvPr id="2" name="Title 1"/>
       <p:cNvSpPr>
         <a:spLocks noGrp="1"/>
       </p:cNvSpPr>
       <p:nvPr>
         <p:ph type="ctrTitle"/>
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


elements
^^^^^^^^

* Placeholder shapes **WERE** created corresponding to layout placeholders of
  the following types in the slide layout:

  * ``'title'``
  * ``'body'``
  * ``'ctrTitle'``
  * ``'subTitle'``
  * ``'obj'``
  * ``'chart'``
  * ``'tbl'``
  * ``'clipArt'``
  * ``'dgm'`` (SmartArt Graphic, perhaps abbreviation of 'diagram')
  * ``'media'``
  * ``'pic'``
  * ``None`` (type attribute default is ``'obj'``)

* Placeholder shapes **WERE NOT** created corresponding to layout placeholders
  of these types in the slide layout:

  * ``'dt'`` (date)
  * ``'sldNum'`` (slide number)
  * ``'ftr'`` (footer)

* The following placeholder types do not apply to slides (they apply to notes
  pages and/or handouts).

  * ``'hdr'``
  * ``'sldImg'``

* **<p:txBody> element**. Placeholder shapes of the following types had a
  ``<p:txBody>`` element. Placeholders of all other types did not have a
  ``<p:txBody>`` element.

  * ``'title'``
  * ``'body'``
  * ``'ctrTitle'``
  * ``'subTitle'``
  * ``'obj'`` (and those without a ``type`` attribute)

* **<p:cNvPr> element**. Each placeholder shape in each slide had a
  ``<p:cNvPr>`` element (it is a required element) populated with *id* and
  *name* attributes.

  * The value of the *id* attribute of ``p:cNvPr`` does not correspond to its
    value in the slide layout. Shape ids were renumbered starting from 2 and
    proceeding in sequence (e.g. 2, 3, 4, 5, ...). Note that id="1" belongs to
    the spTree element, at ``/p:sld/p:cSld/p:spTree/p:nvGrpSpPr/p:cNvPr@id``.

  * The value of the *name* attribute is of the form '%s %d' % (type_name,
    num), where type_name is determined by the placeholder type (e.g. 'Media
    Placeholder') and num appears to default to ``id - 1``. The assigned name
    in each case was similar to its counterpart in the slide layout (same
    *type_name*), but its number was generally different.

* **<p:cNvSpPr>** element. Placeholder shapes of all types had the element
  ``<a:spLocks noGrp="1"/>`` specifying the placeholder could not be grouped
  with any other shapes.

* **<p:spPr>** element. All placeholder shapes had an empty ``<p:spPr>``
  element.


attributes
^^^^^^^^^^

* **type**. The ``type`` attribute of the ``<p:ph>`` element was copied
  without change. Where no ``type`` attribute appeared (defaults to "obj",
  indicating a "content" placeholder), no ``type`` attribute was added to the
  ``<p:ph>`` element.

* **orient**. The ``orient`` attribute of the ``<p:ph>`` element was copied
  without change. Where no ``orient`` attribute appeared (defaults to "horz"),
  no ``orient`` attribute was added to the ``<p:ph>`` element.

* **sz**. The ``sz`` (size) attribute of the ``<p:ph>`` element was copied
  without change. Where no ``sz`` attribute appeared (defaults to "full"), no
  ``sz`` attribute was added to the ``<p:ph>`` element.

* **idx**. The ``idx`` attribute of the ``<p:ph>`` element was copied without
  change. Where no ``idx`` attribute appeared (defaults to "0", indicating the
  singleton title placeholder), no ``idx`` attribute was added to the
  ``<p:ph>`` element.

* **hasCustomPrompt**. The ``hasCustomPrompt`` attribute of the ``<p:ph>``
  element was NOT copied from the slide layout to the slide.


Other observations
==================

Placeholders created from the slide layout were always ``<p:sp>`` elements,
even when they were a placeholder for a shape of another type, e.g.
``<p:graphicFrame>``. When the on-screen placeholder for a table was clicked
to add a table, the ``<p:sp`` element was replaced by a ``<p:graphicFrame>``
element containing the table. The ``<p:ph>`` element was retained in the new
shape. Reapplying the layout to the slide caused the table to move to the
original location of the placeholder (the idx may have been used for
matching). Deleting the table caused the original placeholder to reappear in
its original position.


Conclusions
===========

* Placeholder shapes are added to new slides based on the placeholders present
  in the slide layout specified on slide creation.
* Placeholders for footer, slide number, and date are excluded.
* Each new placeholder is assigned an id starting from 2 and proceeding in
  increments of 1 (e.g. 2, 3, 4, ...).
* Each new placeholder is named based on its type and the id assigned.
* Placeholders are locked against grouping.
* Placeholders of selected types get a minimal ``<p:txBody>`` element.
* Slide placeholders receive an exact copy of the ``<p:ph>`` element in the
  corresponding slide layout placeholder, except that the optional
  ``hasCustomPrompt`` attribute, if present, is not copied.

Open Questions
==============

None.


