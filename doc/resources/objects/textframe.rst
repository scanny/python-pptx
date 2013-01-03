=========
TextFrame
=========

Summary
=======

... corresponds to the ``<p:txBody>`` element, see
:doc:`../schema/ct_textbody.rst`


Description
===========

A shape may have a text frame that can contain display text. A shape is
created with or without a text frame, but one cannot be added after the shape
is created.


Behaviors
=========

* word-wrapping is txBody/bodyPr @wrap="none" or "square"

* Only shape-type shapes (``<p:sp>``) can have a txBody element.


Child Objects
=============

* Paragraph
* ListStyle


Resources
=========

* `TextFrame2 Members`_ interop page on MSDN

.. _TextFrame2 Members:
   http://msdn.microsoft.com/en-us/library/microsoft.office.interop
         .powerpoint.textframe2_members(v=office.14).aspx

