#####
Shape
#####


Summary
=======

The term *shape* is used to designate both a specific object and a category of
objects. The specific object corresponds to the ``<p:sp>`` element. The
category is composed of the object types that can be contained in a group
shape, including GroupShape itself. The following table summarizes the six
top-level shape types, which correspond to the XML elements that are valid
members of a slide's shape tree (``<p:spTree>``) or a group shape
(``<p:grpSp>``).

**Shape types**

============  ====================
shape type    element
============  ====================
shape         ``<p:sp>``
group shape   ``<p:grpSp>``
graphicFrame  ``<p:graphicFrame>``
connector     ``<p:cxnSp>``
picture       ``<p:pic>``
content part  ``<p:contentPart>``
============  ====================

|

Some of these shape types have important sub-types. For example, a placeholder,
a text box, and a preset geometry shape all have an ``<p:sp>`` element as their
root element.

* `Shape.Type Property`_ page on MSDN

.. _Shape.Type Property:
   http://msdn.microsoft.com/en-us/library/office/ff744590(v=office.14).aspx


``<p:sp>`` shape elements
=========================

The ``<p:sp>`` element is used for three types of shape: placeholder, text box,
and geometric shapes. A geometric shape with preset geometry is referred to as
an *auto shape*. Placeholder shapes are documented on the :doc:`placeholder`
page. Auto shapes are documented on the :doc:`autoshape` page.


Geometric shapes
----------------

Geometric shapes are the familiar shapes that may be placed on a slide such as
a rectangle or an ellipse. In the PowerPoint UI they are simply called shapes.
There are two types of geometric shapes, preset geometry shapes and custom
geometry shapes.


XML produced by PowerPoint® client
----------------------------------

.. highlight:: xml

::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Rounded Rectangle 2"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="760096" y="562720"/>
          <a:ext cx="2520824" cy="914400"/>
        </a:xfrm>
        <a:prstGeom prst="roundRect">
          <a:avLst>
            <a:gd name="adj" fmla="val 30346"/>
          </a:avLst>
        </a:prstGeom>
      </p:spPr>
      <p:style>
        <a:lnRef idx="1">
          <a:schemeClr val="accent1"/>
        </a:lnRef>
        <a:fillRef idx="3">
          <a:schemeClr val="accent1"/>
        </a:fillRef>
        <a:effectRef idx="2">
          <a:schemeClr val="accent1"/>
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="lt1"/>
        </a:fontRef>
      </p:style>
      <p:txBody>
        <a:bodyPr rtlCol="0" anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" dirty="0" smtClean="0"/>
            <a:t>This is text inside a rounded rectangle</a:t>
          </a:r>
          <a:endParaRPr lang="en-US" dirty="0"/>
        </a:p>
      </p:txBody>
    </p:sp>


Behaviors
=========

Ids
---

test _BaseSlide._next_id::

    @property
    def _next_id(self):
        """
        Next available id number in slide, starting from 1 and making use
        of any gaps in numbering.
        """
        sld_elm = self._element
        cNvPrs = sld_elm.xpath('//p:cNvPr', namespaces=self.nsmap)
        ids = [int(cNvPr.get('id')) for cNvPr in cNvPrs]
        ids.sort()
        # first gap in sequence wins, or falls off the end as max(ids)+1
        next_id = 1
        for id in ids:
            if id > next_id:
                break
            next_id += 1
        return next_id

* specifically looking for drawing element ids (ST_DrawingElementId)

* these are attribute of cNvPr for the element

* other ids in slide are cTn@id and a:fld@id. cTn@id appears to use a distinct
  id space because it is also numbered starting with 1 and overlaps drawing
  element id numberspace. a:fld@id uses GUIDs, not ints.

* CT_Connection uses ids (as id="999") to indicate endpoints (id references
  rather than id definitions), so just searching for numeric ids won't do the
  trick.


Naming
------

The current UI doesn't seem to care about duplicate names, but at least one
web reference states otherwise.

http://www.pptools.com/starterset/FAQ00036.htm

   **Duplicate names not allowed**. PowerPoint won't allow more than one shape
   on a slide to have the same name. If you already have a shape named XXX,
   you can't give another shape that same name.
   
   When this happens, you'll see a message that explains why the name can't be
   changed, and the name will be reset to the original name in the Object
   Properties dialog box.

   Note: You can have two or more shapes with the same name, as long as
   they're on different slides. The same "no duplicates" rule applies to
   slides within a presentation; no two slides can have the same name.


Child Objects
=============

Other characteristics
---------------------

**Text Frame.** A shape may have a text frame that can contain display text.
A shape is created with or without a text frame, but one cannot be added after
the shape is created.

Only shape-type shapes (``<p:sp>``) can have text associated with them. The
text is contained in a ``<p:txBody>`` element that is a direct child of the
``<p:sp>`` element. The ``<p:txBody>`` element is of type ``CT_TextBody``
defined in the dml schema.


Resources
=========

* `DrawingML Shapes`_ on officeopenxml.com

.. _DrawingML Shapes:
   http://officeopenxml.com/drwShape.php

* `Shape Object MSDN page`_

.. _Shape Object MSDN page:
   http://msdn.microsoft.com/en-us/library/office/ff744177(v=office.14).aspx

* `MsoShapeType Enumeration`_

.. _MsoShapeType Enumeration:
   http://msdn.microsoft.com/en-us/library/office/aa432678(v=office.14).aspx
