.. _MathEquation:


Math equations (OMML / ``m:oMath``)
===================================

PowerPoint 2010 and later can embed Office Math (OMML) equations directly in a
slide shape's text body. The equation element is ``m:oMath`` (with optional
paragraph wrapper ``m:oMathPara``) in the namespace
``http://schemas.openxmlformats.org/officeDocument/2006/math`` -- already
registered in ``pptx.oxml.ns`` under the ``m`` prefix.

Because an equation requires the ``a14`` (Office 2010 drawing) extension
namespace, PowerPoint always writes the hosting shape wrapped in an
``mc:AlternateContent`` element. The ``mc:Choice`` subtree contains the
modern, equation-bearing shape; the ``mc:Fallback`` subtree contains a pre-2010
rendering (typically a plain text shape or picture of the rendered equation)
for legacy consumers.

With Foundation F3 in place (see :ref:`AlternateContent`), the shape inside
``mc:Choice`` is surfaced transparently through ``slide.shapes`` iteration, so
detection of an equation simplifies to an XPath descendant scan on the
shape's ``p:sp`` element.


MVP scope (issue #126)
----------------------

The first iteration provides **read-only** access to the raw OMML:

``BaseShape.has_math_equation``
    Returns ``True`` when the shape's XML subtree contains at least one
    ``m:oMath`` descendant element.

``BaseShape.math_equation_xml``
    Returns the serialized XML of the first ``m:oMath`` descendant, or
    ``None`` when there is no equation.

Implementation is a thin XPath query (``.//m:oMath``) on ``self._element``.
No new custom element class is introduced because, at MVP scope, callers do
not need a structured proxy -- they hand the OMML string off to an external
renderer (pandoc, Microsoft's ``omml.xsl``, etc.).


Explicitly deferred
-------------------

- **Equation authoring.** No ``add_equation(...)`` API; no setter for
  ``math_equation_xml``.
- **LaTeX / MathML conversion.** python-pptx never translates between OMML
  and other math formats.
- **Structured equation object.** No typed ``Equation`` class exposing
  fractions, roots, runs, etc.
- **Multiple-equation indexing.** ``math_equation_xml`` returns only the
  *first* ``m:oMath`` descendant even when the shape contains several.
  Callers who need every equation can walk ``shape.element.xpath(".//m:oMath")``
  directly.

These are tracked under issue #126 for a future iteration.


Specimen XML
------------

.. highlight:: xml

A typical equation-bearing shape in a slide (simplified)::

  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
               Requires="a14">
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="99" name="Math-Equation"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:oMathPara>
              <m:oMath>
                <m:r><m:t>x</m:t></m:r>
                <m:r><m:t>=</m:t></m:r>
                <m:r><m:t>1</m:t></m:r>
              </m:oMath>
            </m:oMathPara>
          </a:p>
        </p:txBody>
      </p:sp>
    </mc:Choice>
    <mc:Fallback>
      <p:sp>...legacy rendering...</p:sp>
    </mc:Fallback>
  </mc:AlternateContent>

Foundation F3 walks the ``mc:Choice`` subtree and surfaces the inner ``p:sp``
as a regular shape in ``slide.shapes``. ``has_math_equation`` returns ``True``
for that shape and ``math_equation_xml`` returns the serialized ``m:oMath``
fragment.
