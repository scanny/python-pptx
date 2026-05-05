"""lxml custom element classes for the `a:` shape-style family.

The ``<p:style>`` child of a ``<p:sp>``, ``<p:cxnSp>``, or ``<p:pic>`` is typed
``a:CT_ShapeStyle`` and contains four mandatory references into the theme's
format-scheme:

- ``<a:lnRef idx="…">`` — index into ``a:fmtScheme/a:lnStyleLst``.
- ``<a:fillRef idx="…">`` — index into ``a:fmtScheme/a:fillStyleLst``.
- ``<a:effectRef idx="…">`` — index into ``a:fmtScheme/a:effectStyleLst``.
- ``<a:fontRef idx="major|minor|none">`` — key into ``a:fontScheme``.

Each reference may carry a color-choice child (``a:schemeClr`` / ``a:srgbClr``
/ …) that specifies the color the referenced style is tinted with; for
theme-style authoring the library emits ``<a:schemeClr val="accent1"/>``
(matching what PowerPoint writes for "Shape Styles" gallery picks).

Grammar reference: ECMA-376 Part 1 §20.1.4.1.29 ``CT_ShapeStyle``,
§20.1.4.1.17 ``CT_StyleMatrixReference``, §20.1.4.1.18 ``CT_FontReference``,
§20.1.4.1.5 ``ST_StyleMatrixColumnIndex`` / ``ST_FontCollectionIndex`` (see
``spec/ISO-IEC-29500-1/schemas/xsd/dml-main.xsd``).
"""

from __future__ import annotations

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne, RequiredAttribute


class _StyleRefBase(BaseOxmlElement):
    """Shared helpers for `a:lnRef`/`a:fillRef`/`a:effectRef`/`a:fontRef`.

    Exposes a ``scheme_color`` read/write accessor for the nested
    ``<a:schemeClr val="…"/>`` child — the "Shape Styles" gallery sets this
    alongside the ``idx`` to select which theme color the referenced style is
    tinted with. Reads return the ``val`` attribute (e.g. ``"accent1"``,
    ``"bg1"``, ``"tx2"``) or |None| if no ``a:schemeClr`` child is present
    (the spec allows other color-choice children such as ``a:srgbClr`` but
    PowerPoint's gallery presets only emit ``a:schemeClr``).
    """

    @property
    def scheme_color(self) -> str | None:
        """Value of the nested ``a:schemeClr/@val``, or |None| when absent."""
        schemeClr = self.find(qn("a:schemeClr"))
        if schemeClr is None:
            return None
        val = schemeClr.get("val")
        return val if val is None else str(val)

    @scheme_color.setter
    def scheme_color(self, value: str | None) -> None:
        from lxml import etree

        # -- remove any existing color-choice child (schemeClr, srgbClr, …) --
        for child_tag in (
            "a:schemeClr",
            "a:srgbClr",
            "a:sysClr",
            "a:prstClr",
            "a:hslClr",
            "a:scrgbClr",
        ):
            for existing in self.findall(qn(child_tag)):
                self.remove(existing)
        if value is None:
            return
        schemeClr = etree.SubElement(self, qn("a:schemeClr"))
        schemeClr.set("val", value)


class CT_StyleMatrixReference(_StyleRefBase):
    """`a:lnRef`, `a:fillRef`, or `a:effectRef` element.

    References a line-, fill-, or effect-style within the theme's
    ``a:fmtScheme``. The ``idx`` attribute is an ``ST_StyleMatrixColumnIndex``
    (``xsd:unsignedInt``) where ``0`` means "no style from the matrix" and
    ``1..N`` indexes into the matching style list.
    """

    idx: int = RequiredAttribute("idx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_FontReference(_StyleRefBase):
    """`a:fontRef` element.

    Selects ``major`` / ``minor`` / ``none`` from the theme's font-scheme.
    The ``idx`` attribute is ``ST_FontCollectionIndex`` — a restricted string
    — not an integer.
    """

    idx: str = RequiredAttribute("idx", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ShapeStyle(BaseOxmlElement):
    """`p:style` element on a `p:sp`, `p:cxnSp`, or `p:pic`.

    Four mandatory ordered children — ``a:lnRef``, ``a:fillRef``,
    ``a:effectRef``, ``a:fontRef`` — referencing the theme's format- and
    color-schemes.
    """

    lnRef: CT_StyleMatrixReference = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:lnRef"
    )
    fillRef: CT_StyleMatrixReference = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:fillRef"
    )
    effectRef: CT_StyleMatrixReference = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:effectRef"
    )
    fontRef: CT_FontReference = OneAndOnlyOne("a:fontRef")  # pyright: ignore[reportAssignmentType]
