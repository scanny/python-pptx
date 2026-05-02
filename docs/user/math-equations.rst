
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


Inserting an equation
---------------------

`#528 <https://github.com/scanny/python-pptx/issues/528>`_ adds the write
counterpart to the read surface above.

``_Paragraph.add_math_equation(omml_xml)``
    Append an OMML equation to the paragraph. ``omml_xml`` is a Unicode
    string whose root element is ``m:oMath`` or ``m:oMathPara`` (and
    declares the ``m`` math namespace on the root). The method wraps the
    fragment in the ``mc:AlternateContent`` / ``mc:Choice[Requires="a14"]``
    / ``a14:m`` scaffolding PowerPoint emits for an inline equation, and
    pairs it with an ``mc:Fallback`` run that carries a plain-text
    rendering (the concatenated ``m:t`` contents of the OMML) so
    pre-2010 consumers see *something* readable. Existing runs,
    line-breaks, and fields already in the paragraph are preserved.

Example -- insert ``E=mc²`` after a label run::

    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    paragraph = shape.text_frame.paragraphs[0]
    paragraph.text = "Einstein: "
    paragraph.add_math_equation(
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        '<m:r><m:t>E=mc^2</m:t></m:r>'
        '</m:oMath>'
    )
    prs.save("einstein.pptx")

Producing the OMML itself is **not** in scope for python-pptx. The
simplest approach is to write the equation as MathML and transform it
with Microsoft's `MML2OMML.XSL` (shipped with Microsoft Office) or a
third-party library. The caller-supplied OMML string is written verbatim
into the ``a14:m`` wrapper.


Scope -- deferred functionality
-------------------------------

The following are explicitly deferred:

- **LaTeX / MathML conversion.** python-pptx does not translate between OMML
  and LaTeX (or MathML). Use an external tool such as pandoc or the
  `omml.xsl` XSLT shipped by Microsoft Office.
- **Structured equation proxy.** There is no typed ``Equation`` class exposing
  run/fraction/root sub-elements. Callers work directly with the serialized
  XML string.
- **Replacing / editing an existing equation.** ``add_math_equation`` appends
  a new equation to a paragraph. To replace an existing equation, clear the
  paragraph (or the relevant ``mc:AlternateContent`` child) via the raw XML
  element first.

Round-trip fidelity for equation-bearing shapes is provided by Foundation F3:
the surrounding ``mc:AlternateContent`` (including the ``mc:Fallback``
subtree) is preserved verbatim on save, so a read-only tool that only inspects
equations will not corrupt a presentation.
