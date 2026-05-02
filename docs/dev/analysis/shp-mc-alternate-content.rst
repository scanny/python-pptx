.. _AlternateContent:


Markup-Compatibility AlternateContent
=====================================

`mc:AlternateContent` is the Markup-Compatibility (`mc:`) wrapper defined in
ISO/IEC 29500-3. Office 2010+ uses it to ship extended-namespace content
alongside a pre-existing-schema ("legacy") rendering, so that older consumers
that only understand the original ECMA-376 schema can still display something
reasonable.

In a shape-tree context, PowerPoint emits `mc:AlternateContent` when a shape
requires a feature that depends on a non-base namespace -- for example an
equation (`m:`), a 2014+ chart (`cx:`), a modern comment (`p188:`), or a
2010-era effect on an otherwise plain shape.

Shape-tree iterator semantics
-----------------------------

`CT_GroupShape.iter_shape_elms` transparently walks into the **first**
`mc:Choice` child of an `mc:AlternateContent` element. The shapes inside that
Choice are yielded inline with any unwrapped siblings, preserving XML document
order.

`mc:Fallback` is never yielded; it is left in place on the parent element so
that the wrapper survives unmodified across a load-and-save round trip. This
is important because Fallback content is sometimes the only thing a
non-PowerPoint consumer can render, and stripping it would degrade fidelity.

When the first `mc:Choice` itself contains another `mc:AlternateContent`
(unusual, but schema-legal), the helper recursively descends into that nested
first-Choice as well.

Only direct children of a `mc:Choice` are classified against the caller-
supplied `shape_tags` tuple (e.g. `p:sp`, `p:grpSp`, `p:graphicFrame`,
`p:cxnSp`, `p:pic`, `p:contentPart`). Other intermediate elements are ignored
but left on the tree.


XML Specimens
-------------

.. highlight:: xml

A minimal AlternateContent-wrapped shape::

  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
               Requires="a14">
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="99" name="Modern-Choice"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:sp>
    </mc:Choice>
    <mc:Fallback>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="100" name="Fallback-Shape"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:sp>
    </mc:Fallback>
  </mc:AlternateContent>


Schema Reference (ISO/IEC 29500-3)
----------------------------------

`mc:AlternateContent` is defined in Part 3, "Markup Compatibility and
Extensibility". The relevant grammar:

- `mc:AlternateContent` has 1..n `mc:Choice` children and an optional final
  `mc:Fallback` child.
- Each `mc:Choice` carries a required `Requires` attribute listing the
  extension namespaces it uses (space-separated prefix list).
- `mc:Fallback` content must use only elements from the "known" namespaces
  declared at the package level; it is what a conformant legacy consumer
  should render.

python-pptx's policy is to surface the first `mc:Choice` (the richest) and
preserve `mc:Fallback` as-is. It does not currently negotiate `Requires` --
i.e. it does not skip a Choice when its required namespaces are unknown. In
practice this works because PowerPoint writes the richest Choice first and
consumers that follow link it, and because the library keeps any unrecognized
elements in the parse tree unchanged.
