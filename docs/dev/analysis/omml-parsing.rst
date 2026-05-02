.. _OMMLParsing:


OMML parsing (issue #892 design sketch)
=======================================

This page sketches the design space for **structured parsing** of Office Math
(OMML) equations. It is a planning / deferral document, not a specification of
shipped behavior.


Issue history
-------------

- **#126** — *read math equation via F3*. Landed in Wave 2 as the MVP for
  equation support. It exposes ``Shape.has_math_equation`` (bool) and
  ``Shape.math_equation_xml`` (serialized ``m:oMath`` subtree or ``None``).
  Detection relies on a simple ``.//m:oMath`` XPath descendant scan of the
  shape's element; the equation shape itself is surfaced by Foundation F3
  traversing ``mc:AlternateContent`` transparently. See
  :ref:`MathEquation` for the analysis, and the
  user-facing guide at ``docs/user/math-equations.rst``.

- **#892** — *Support for parsing Equations*. The original report asked
  python-pptx to return, for a given slide / shape, "the equations that are
  in it" so the caller can render or otherwise process them. The find-
  equation-in-slide portion of that request is **resolved by #126's API**
  (see the regression test
  ``DescribeBaseShape.it_surfaces_an_OMML_equation_on_a_real_slide_regression_892``
  in ``tests/shapes/test_base.py``). What remains open — and what this page
  plans for — is *structured* parsing of OMML: returning a typed Python
  object graph the caller can inspect and manipulate without touching XML.


Why a follow-up is (or isn't) needed
------------------------------------

Real-world callers who reach for "equation parsing" usually want one of:

1. **Discover** the equations in a presentation so they can be rendered
   out-of-band (to LaTeX / MathML / PNG). *Resolved by #126.* The raw OMML
   string is exactly the input format that pandoc, Microsoft's ``omml.xsl``,
   and other converters already accept.

2. **Extract a human-readable form** of each equation (LaTeX, plain Unicode,
   MathML). This is *conversion*, not parsing; it is explicitly out of scope
   for python-pptx per the #126 deferral list. Callers should pipe
   ``math_equation_xml`` into an external converter.

3. **Inspect or mutate** the structure of an equation in place — e.g. "change
   the exponent of the third ``m:sSup`` to 3". This is the only bucket that
   genuinely requires structured parsing inside python-pptx.

Bucket (3) has not been the subject of a concrete user request on the tracker
to date; issue #892's two linked comments describe bucket (1) and bucket (2).
That motivates leaving structured parsing deferred until a real need surfaces.


Scope of a future structured-parsing feature
--------------------------------------------

If a future iteration elects to implement structured OMML parsing, the
natural shape of the work is:

**A.** Introduce custom element classes for the OMML grammar under
``pptx.oxml.math``, paralleling how DrawingML is modeled under
``pptx.oxml.dml``. At minimum the top-level and most common math primitives:

================  ====================================================
Element           Role
================  ====================================================
``m:oMath``       the root equation container
``m:oMathPara``   paragraph-level wrapper around one or more ``m:oMath``
``m:r``           math run (carries ``m:rPr`` + ``m:t`` text)
``m:f``           fraction (``m:num`` / ``m:den``)
``m:sSup``        superscript (``m:e`` / ``m:sup``)
``m:sSub``        subscript  (``m:e`` / ``m:sub``)
``m:rad``         radical / root (``m:deg`` / ``m:e``)
``m:nary``        n-ary operator (sum / product / integral)
``m:d``           delimiter group (``m:e``)
``m:m``           matrix (``m:mr`` / ``m:mcs``)
``m:func``        function apply (``m:fName`` / ``m:e``)
================  ====================================================

The full OMML grammar is defined in the ISO/IEC 29500-1 "Math" schema. A
complete round-trip implementation would cover all ~40 elements; a
pragmatic first pass would register the subset above and fall through to
``BaseOxmlElement`` for the long tail.

**B.** Expose a ``Shape.math_equations`` read-property returning a tuple of
``Equation`` proxy objects (one per ``m:oMath`` descendant, not just the
first). Each ``Equation`` wraps its ``CT_OMath`` element and offers a
narrowly-typed API, e.g.::

    class Equation:
        @property
        def xml(self) -> str: ...          # equivalent of today's math_equation_xml
        @property
        def runs(self) -> tuple[Run, ...]: ...
        def walk(self) -> Iterator[MathNode]: ...  # depth-first structural walk

The current ``Shape.math_equation_xml`` property continues to work as the
"escape hatch" for callers who prefer raw XML.

**C.** Decide on a conversion story. Two viable positions:

- **Stay out of conversion.** Keep python-pptx focused on the OOXML object
  model; direct callers toward pandoc or ``omml.xsl`` for LaTeX / MathML.
  This matches the current #126 deferral and is the lowest-maintenance
  outcome.

- **Ship a minimal OMML→MathML converter.** MathML is an XML sibling format
  and the transformation is largely mechanical (Microsoft's ``omml.xsl``
  open-sources a reference implementation). This would add substantial
  surface area — LaTeX in particular has many edge cases — and is not
  recommended without a dedicated maintainer.


Non-goals
---------

The following are out of scope even for a future structured-parsing
iteration:

- **Math rendering**. python-pptx never rasterizes equations.
- **Syntactic validation**. Parsing is lenient; a malformed OMML subtree
  surfaces as generic ``BaseOxmlElement`` children rather than raising.
- **Authoring convenience API**. ``add_equation("x = 1")`` or similar
  constructors are a separate feature (``#126`` follow-up "writing OMML").
- **Support for legacy Equation Editor 3.0 (MathType) blobs** embedded as
  OLE objects in pre-2010 decks. Those are OLE streams, not OMML, and would
  be tracked as a separate issue.


Interop with Foundation F3
--------------------------

Equation-bearing shapes are always wrapped by PowerPoint in
``mc:AlternateContent`` (because they require the ``a14`` namespace). The
``mc:Choice`` subtree carries the modern ``p:sp`` with an ``m:oMath`` child;
the ``mc:Fallback`` subtree carries a 2007-compatible rendering (usually a
static picture of the equation).

Foundation F3 (see :ref:`AlternateContent`) surfaces the ``mc:Choice`` shape
transparently through ``slide.shapes`` iteration and preserves the full
``mc:AlternateContent`` subtree on round-trip save. Any future structured-
parsing feature therefore does **not** need to re-implement alternate-content
traversal — it can assume the equation-bearing ``p:sp`` is what it receives
from the shape tree.


Testing strategy
----------------

A structured-parsing follow-up would reuse the existing
``features/steps/test_files/shp-math-equation.pptx`` fixture (a Word-authored
equation rendered in a PowerPoint shape) plus new fixtures covering:

- nested fractions and radicals;
- ``m:oMathPara`` with multiple ``m:oMath`` siblings (already partially
  covered by ``and_it_returns_only_the_first_equation_when_multiple_are_present``);
- an equation saved by LibreOffice Impress, to guard against MS-specific
  assumptions in the parser.


Effort estimate
---------------

- **B (proxy object + ~12 element classes)**: ~1–2 weeks for a single
  contributor, plus ongoing maintenance as OMML evolves in new Office
  versions.
- **C (OMML→MathML)**: ~1 week to port / wrap ``omml.xsl``; OMML→LaTeX is
  several weeks of edge-case work and almost certainly belongs in a
  dedicated downstream library, not in python-pptx.

Until a motivated contributor brings a concrete use case that the raw-XML
escape hatch cannot serve, structured OMML parsing remains **deferred**.
