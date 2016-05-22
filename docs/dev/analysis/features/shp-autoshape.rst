.. _autoshape:

Auto Shape
==========


Summary
-------

*AutoShape* is the name the MS Office API uses for a shape with preset
geometry, those referred to in the PowerPoint UI simply as *shape*. Typical
examples are a rectangle, a circle, or a star shape. An auto shape has a type
(preset geometry), an outline line style, a fill, and an effect, in addition to
a position and size. The outline, fill, and effect can each be None.

Each auto shape is contained in an ``<p:sp>`` element.


Auto shape type
---------------

There are 187 pre-defined auto shape types, each corresponding to a distinct
preset geometry. Some types have one or more *adjustment value*, which
corresponds to the position of an adjustment handle (small yellow diamond) on
the shape in the UI that changes an aspect of the shape geometry, for example
the corner radius of a rounded rectangle.

* `MsoAutoShapeType Enumeration`_ page on MSDN


Adjustment values
-----------------

Many auto shapes may be adjusted to change their shape beyond just their width,
height, and rotation. In the PowerPoint application, this is accomplished by
dragging small yellow diamond-shaped handles that appear on the shape when it
is selected.

In the XML, the position of these adjustment handles is reflected in the
contents of the ``<a:avLst>`` element::

    <a:prstGeom prst="roundRect">
      <a:avLst>
        <a:gd name="adj" fmla="val 30346"/>
      </a:avLst>
    </a:prstGeom>

The |avLst| element can be empty, even if the shape has available adjustments.

guide
   a guide is essentially a variable within DrawingML. Its value is the result
   of a simple expression like ``"*/ 100000 w ss"``. Its value may then be used
   in other expressions or directly to specify a point, by referring to it by
   name.

``<a:gd>`` stands for *guide* ...

Where are the custom geometries defined?
   The geometries for autoshapes are defined in ``presetShapeDefinitions.xml``,
   located in:

   ``pptx/ref/ISO-IEC-29500-1/schemas/dml-geometries/OfficeOpenXML-DrawingMLGeometries/``.

What is the range of the formula values?
   In general, the nominal adjustment value range for the preset shapes is 0 to
   100000. However, the actual value stored can be outside that range,
   depending on the formulas involved. For example, in the case of the chevron
   shape, the range is from 0 to 100000 * width / short-side. When the shape is
   taller than it is wide, width is the short side and the range is 0 to
   100000. However, when the shape is wider than it is tall, the range can be
   much larger. Required values are probably best discovered via
   experimentation or calculation based on the preset formula.

The |gd| element does not appear initially when the shape is inserted. However,
if the adjustment handle is moved and then moved back to the default value, the
|gd| element is not removed. The default value becomes the effective value for
the guide if the |gd| element is not present.

The adjustment value is preserved over scaling of the object. Its effective
value can change however, depending on the formula. In the case of the chevron
shape, the effective value is a function of ``ss``, the "short side", and the
adjustment handle moves on screen when the chevron is made wider than it is
tall.

A certain level of documentation for auto shapes could be generated from the
XML shape definition file. The preset name and the number of adjustment values
it has (guide definitions) and the adjustment value name can readily be
determined by inspection of the XML.

What coordinate system are the formula values expressed in?
   There appear to be two coordinate systems at work. The 0 to 100000 bit the
   adjustment handles are expressed in appears to be the (arbitrary) shape
   coordinate system. The built-in values like ``w``, ``ss``, etc., appear to
   be in the slide coordinate system, the one the ``xfrm`` element traffics in.

   Probably worth experimenting with some custom shapes to find out (edit and
   rezip with bash one-liner).

useful web resources
   What useful web resources are out there on DrawingML?

   http://officeopenxml.com/drwSp-custGeom.php


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
---------

* `DrawingML Shapes`_ on officeopenxml.com

* `Shape Object MSDN page`_

* `MsoShapeType Enumeration`_

.. _DrawingML Shapes:
   http://officeopenxml.com/drwShape.php

.. _Shape Object MSDN page:
   http://msdn.microsoft.com/en-us/library/office/ff744177(v=office.14).aspx

.. _MsoShapeType Enumeration:
   http://msdn.microsoft.com/en-us/library/office/aa432678(v=office.14).aspx

.. _`MsoAutoShapeType Enumeration`:
   http://msdn.microsoft.com/en-us/library/office/aa432469(v=office.14).aspx

.. |avLst| replace:: ``<a:avLst>``

.. |gd| replace:: ``<a:avLst>``
