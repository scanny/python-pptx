=====
Shape
=====

Summary
=======

The term *shape* is used to designate both a specific object and a category of
objects. The specific object corresponds to the ``<p:sp>`` element. The
category is composed of the object types that can be contained in a group
shape, including GroupShape itself.

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


* `Shape.Type Property`_ page on MSDN

.. _Shape.Type Property:
   http://msdn.microsoft.com/en-us/library/office/ff744590(v=office.14).aspx


Description
===========

... graphic frame is used to hold table, perhaps other objects ...

**Text Frame.** A shape may have a text frame that can contain display text.
A shape is created with or without a text frame, but one cannot be added after
the shape is created.


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

* Only shape-type shapes (``<p:sp>``) can have text associated with them. The
  text is contained in a ``<p:txBody>`` element that is a direct child of the
  ``<p:sp>`` element. The ``<p:txBody>`` element is of type ``CT_TextBody``
  defined in the dml schema.


Specifications
==============


Related Specifications
======================


Resources
=========

* `Shape Object MSDN page`_

.. _Shape Object MSDN page:
   http://msdn.microsoft.com/en-us/library/office/ff744177(v=office.14).aspx

* `MsoShapeType Enumeration`_

.. _MsoShapeType Enumeration:
   http://msdn.microsoft.com/en-us/library/office/aa432678(v=office.14).aspx
