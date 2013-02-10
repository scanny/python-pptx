============
Placeholders
============

:Updated:  2013-01-01
:Versions: python-pptx 0.1a
:Author:   Steve Canny
:Status:   *DRAFT*

.. :Contributors:

Summary
=======

* Placeholders are a subset of shapes, not a separate class. There is however
  a `PlaceholderFormat object`_ in the MS API accessed via
  ``Shape.PlaceholderFormat`` that provides access to the *name*, *type*, and
  *contained_type* attributes of the placeholder.

* Placeholders are accessed through a slide-type object's shape collection,
  e.g. ``sld.shapes.placeholders``. The contents of the placeholders
  collection is a subset of the shapes in the shape collection.

* The title placeholder, if it exists, always appears first in the
  placeholders collection.

* Placeholders can only be top-level shapes, they cannot be nested in a group
  shape.


Description
===========

From the `Shape Object MSDN page`_:

   Each Shape object in the Placeholders collection represents a placeholder
   for text, a chart, a table, an organizational chart, or some other type of
   object. If the slide has a title, the title is the first placeholder in the
   collection.


Behaviors
=========

Automatic Naming
----------------

* Most placeholders are automatically named '{root_name} Placeholder {num}'
  where *root_name* is something like ``Chart`` and num is a positive integer.
  A typical example is ``Table Placeholder 3``.

* A placeholder with vertical orientation, i.e. ``<p:ph orient="vert">``, is
  prefixed with ``'Vertical '``, e.g. ``Vertical Text Placeholder 2``.

* The word ``'Placeholder'`` is omitted when the type is 'title', 'ctrTitle',
  or 'subTitle'.


On slide creation
-----------------

When a new slide is added to a presentation, empty placeholders are added to it based on the placeholders present in the specified slide layout.

**Working hypothesis.** Only content placeholders are created automatically.
Special placeholders must be created individually with additional calls.
Special placeholders include header, footer, page numbers, date, and slide
number; and perhaps others. In this respect, it seems like special
placeholders on a slide layout are simply default settings for those
placeholders in case they are added later.

During slide life-cycle
-----------------------

* Placeholders can be deleted from slides and can be restored by calling
  methods like AddTitle on the shape collection. AddTitle raises if there's
  already a title placeholder, so need to call .hasTitle() beforehand.

Experiment(s)
-------------

**Experiment:** Add a title-slide layout slide, leave the placeholder text
blank, and see what XML comes up. Look into the slide layout to see how the
two correspond.

**RESULT:** Only the title and subtitle placeholder shapes are created.
Footer, page number, date, etc. shapes were not created. There are menu
options to add those to the slide.


Specifications
==============

* (?) Where can find documentation on what the different placeholder types
  mean?

* The placeholder with idx="0" is the title placeholder. "0" is the default
  for the *idx* attribute on <p:ph>, so one with no idx attribute is the title
  placeholder.

* The document order of placeholders signifies their z-order and has no
  bearing on their index order. :attr:`ShapeCollection.placeholders` contains
  the placeholders in order of *idx* value, which means the title placeholder
  always appears first in the sequence, if it is present.


Related Specifications
======================

* The `Placeholders Object documentation`_ on MSDN is a useful source.

* :doc:`../../schema/ct_placeholder`


Resources
=========

* `PlaceholderFormat Object`_ documentation on MSDN

.. _PlaceholderFormat Object:
   http://msdn.microsoft.com/en-us/library/office/ff745007(v=office.14).aspx

* `Placeholders Object documentation`_ on MSDN

.. _Placeholders Object documentation:
   http://msdn.microsoft.com/en-us/library/office/ff746338(v=office.14).aspx

* `PlaceholderFormat.Type Property`_ possible values

.. _PlaceholderFormat.Type Property:
   http://msdn.microsoft.com/en-us/library/office/ff745930(v=office.14).aspx

* `Shape Object MSDN page`_

.. _Shape Object MSDN page:
   http://msdn.microsoft.com/en-us/library/office/ff744177(v=office.14).aspx

* `MS Shapes.Placeholders API`_ 

.. _MS Shapes.Placeholders API:
   http://msdn.microsoft.com/en-us/library/office/ff744297(v=office.14).aspx

