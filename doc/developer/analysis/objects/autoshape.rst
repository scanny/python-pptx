##########
Auto Shape
##########


Summary
=======

*AutoShape* is the name the MS Office API uses for a shape with preset
geometry, those referred to in the PowerPoint UI simply as *shape*. Typical
examples are a rectangle, a circle, or a star shape. An auto shape has a type
(preset geometry), an outline line style, a fill, and an effect, in addition to
a position and size. The outline, fill, and effect can each be None.

Each auto shape is contained in an ``<p:sp>`` element.


Auto shape type
===============

There are 188 pre-defined auto shape types, each corresponding to a distinct
preset geometry. Some types have one or more *adjustment value*, which
corresponds to the position of an adjustment handle (small yellow diamond) on
the shape in the UI that changes an aspect of the shape geometry, for example
the corner radius of a rounded rectangle.

* `MsoAutoShapeType Enumeration`_ page on MSDN

.. _`MsoAutoShapeType Enumeration`:
   http://msdn.microsoft.com/en-us/library/office/aa432469(v=office.14).aspx


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

