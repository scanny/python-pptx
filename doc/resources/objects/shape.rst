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

Naming
------

http://www.pptools.com/starterset/FAQ00036.htm

Duplicate names are not allowed
PowerPoint won't allow more than one shape on a slide to have the same name. If you already have a shape named XXX, you can't give another shape that same name.

When this happens, you'll see a message that explains why the name can't be changed, and the name will be reset to the original name in the Object Properties dialog box.

Note:

You can have two or more shapes with the same name, as long as they're on different slides.
The same "no duplicates" rule applies to slides within a presentation; no two slides can have the same name.

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
