
Working with math equations
===========================

PowerPoint 2010 introduced native support for Office Math (OMML) equations in
slide text. When a slide contains an equation, PowerPoint wraps the owning
shape in a Markup-Compatibility fallback block (`mc:AlternateContent`) so that
older consumers can still render *something* for the same region of the slide.
The richest rendering -- the `mc:Choice` subtree -- contains an `m:oMath`
element that carries the equation markup.

python-pptx surfaces these wrapped shapes through normal `slide.shapes`
iteration (see the Markup-Compatibility traversal added as Foundation F3) and
provides two read-only properties on every shape for discovering and
extracting the raw OMML:

``Shape.has_math_equation``
    ``True`` when the shape's XML subtree contains at least one ``m:oMath``
    descendant, ``False`` otherwise.

``Shape.math_equation_xml``
    A Unicode string containing the raw OMML (the full ``<m:oMath>`` subtree)
    of the first equation in the shape, or ``None`` if the shape contains no
    equation. The serialization includes every namespace declaration lxml
    emits, so the returned fragment is well-formed on its own.

Example::

    from pptx import Presentation

    prs = Presentation("my-equations.pptx")
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_math_equation:
                print(shape.math_equation_xml)

The resulting OMML can be fed into an external library such as
`pandoc <https://pandoc.org/>`_ or a custom XSLT pipeline to convert to
LaTeX, MathML, or an image.


Scope -- deferred functionality
-------------------------------

The initial MVP for `#126
<https://github.com/scanny/python-pptx/issues/126>`_ is **read-only**. The
following are explicitly deferred to a later iteration:

- **Programmatic equation creation.** There is no supported API for inserting
  or replacing an equation in a shape. Callers who need to modify OMML must
  edit the XML element directly via ``shape.element`` and save the
  presentation; this is effectively an unsupported escape hatch.
- **LaTeX / MathML conversion.** python-pptx does not translate between OMML
  and LaTeX (or MathML). Use an external tool such as pandoc or the
  `omml.xsl` XSLT shipped by Microsoft Office.
- **Structured equation proxy.** There is no typed ``Equation`` class exposing
  run/fraction/root sub-elements. Callers work directly with the serialized
  XML string.

Round-trip fidelity for equation-bearing shapes is provided by Foundation F3:
the surrounding ``mc:AlternateContent`` (including the ``mc:Fallback``
subtree) is preserved verbatim on save, so a read-only tool that only inspects
equations will not corrupt a presentation.
