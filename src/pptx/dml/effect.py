"""DrawingML visual-effect objects such as shadow, glow, reflection, and soft-edge.

An `a:effectLst` container may appear on any DrawingML container that has shape-properties
semantics: `p:spPr`, `p:grpSpPr`, `p:bgPr` (slide background), `a:tcPr` (table cell),
`a:rPr` (text run properties), and a handful of others. This module defines the shared
Python-layer descriptor family for reading and writing these effects uniformly.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.dml.color import ColorFormat
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.dml.effect import (
        CT_EffectList,
        CT_GlowEffect,
        CT_OuterShadowEffect,
        CT_ReflectionEffect,
        CT_SoftEdgesEffect,
    )
    from pptx.util import Length


class EffectFormat(object):
    """Provides uniform access to the `a:effectLst` container on a shape-like element.

    An `EffectFormat` surfaces each of the independent effect sub-families
    (`.shadow`, `.glow`, `.reflection`, `.soft_edge`) using a shared writer for
    `a:effectLst` so any shape, chart, or text container exposes effects uniformly.
    """

    def __init__(self, spPr):
        # -- `spPr` may be any element type that has an `effectLst` child in its schema,
        # -- e.g. `p:spPr`, `p:grpSpPr`, `p:bgPr`, `a:tcPr`, `a:rPr`.
        self._element = spPr

    @property
    def inherit(self) -> bool:
        """True when the container inherits its effects from the style hierarchy.

        Equivalent to the absence of an `a:effectLst` child. Writable; assigning True
        removes any explicit `a:effectLst` (restoring inheritance for *all* effects on the
        element). Assigning False materializes an empty `a:effectLst` (suppressing inherited
        effects).
        """
        return self._element.effectLst is None

    @inherit.setter
    def inherit(self, value: bool | None) -> None:
        if bool(value):
            self._element._remove_effectLst()
        else:
            self._element.get_or_add_effectLst()

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """|ShadowFormat| object providing access to shadow-effect settings."""
        return ShadowFormat(self._element)

    @lazyproperty
    def glow(self) -> GlowFormat:
        """|GlowFormat| object providing access to glow-effect settings."""
        return GlowFormat(self._element)

    @lazyproperty
    def reflection(self) -> ReflectionFormat:
        """|ReflectionFormat| object providing access to reflection-effect settings."""
        return ReflectionFormat(self._element)

    @lazyproperty
    def soft_edge(self) -> SoftEdgeFormat:
        """|SoftEdgeFormat| object providing access to soft-edge-effect settings."""
        return SoftEdgeFormat(self._element)


class _EffectChildFormat(object):
    """Base class for single-effect accessors (glow / reflection / soft-edge / shadow).

    Holds a reference to the parent spPr-like element and provides a uniform hook to
    locate or create the effect child under `a:effectLst`.
    """

    #: Subclass-level override: local tag of the effect element, e.g. `"glow"`.
    _child_tag: str = ""

    def __init__(self, spPr):
        self._element = spPr

    # -- helpers shared by all sub-types --------------------------------------

    @property
    def _effectLst(self) -> CT_EffectList | None:
        """The enclosing `a:effectLst` element, or None if absent."""
        return self._element.effectLst

    @property
    def _child(self):
        """The effect child element (e.g. `a:glow`) if present, else None."""
        effectLst = self._effectLst
        if effectLst is None:
            return None
        return getattr(effectLst, self._child_tag)

    def _get_or_add_child(self):
        """Return the effect child element, creating effectLst and child as needed."""
        effectLst = self._element.get_or_add_effectLst()
        return getattr(effectLst, "get_or_add_" + self._child_tag)()

    def _remove_child(self) -> None:
        """Remove the effect child if present. Leaves the `a:effectLst` in place."""
        effectLst = self._effectLst
        if effectLst is None:
            return
        getattr(effectLst, "_remove_" + self._child_tag)()

    # -- public API shared across sub-types -----------------------------------

    @property
    def inherit(self) -> bool:
        """True when this specific effect is absent and inherits from the style hierarchy.

        An effect inherits when either the parent `a:effectLst` is absent *or* the specific
        effect child (e.g. `a:glow`) is absent within an explicit `a:effectLst`. Note that
        materializing `a:effectLst` without populating it suppresses *all* inherited
        effects, so individual-child inheritance semantics only hold when `a:effectLst` is
        also absent; this property reports that union for simplicity.
        """
        return self._child is None

    @inherit.setter
    def inherit(self, value: bool | None) -> None:
        if bool(value):
            self._remove_child()
        else:
            self._get_or_add_child()


def _default_srgb(tag: str) -> str:
    """Return an `a:srgbClr` default-black color-choice XML fragment for `tag`."""
    from pptx.oxml.ns import nsdecls

    return '<a:srgbClr %s val="000000"/>' % nsdecls("a")


class ShadowFormat(_EffectChildFormat):
    """Provides access to shadow-effect settings on a DrawingML shape-like element.

    The `ShadowFormat` currently manages an `a:outerShdw` child. Backward-compatible:
    `ShadowFormat(spPr).inherit` behaves as in the pre-F2 skeleton, reporting True when
    the container has no `a:effectLst` at all.
    """

    _child_tag = "outerShdw"

    @property
    def inherit(self) -> bool:
        """True if shape inherits shadow settings.

        Read/write. An explicitly-defined `a:effectLst` on the shape (regardless of
        contents) causes this property to return False. A shape with no explicitly-defined
        `a:effectLst` inherits its shadow settings from the style hierarchy (and so
        returns True).

        Assigning True causes any explicitly-defined effects to be removed and inheritance
        is restored. Note this has the side-effect of removing **all** explicitly-defined
        effects, such as glow and reflection, and restoring inheritance for all effects on
        the shape. Assigning False causes the inheritance link to be broken and **no**
        effects to appear on the shape; this also zeroes any sibling
        `p:style/a:effectRef/@idx` so theme-inherited shadow is fully suppressed (issue
        #446). Note that the original `effectRef/@idx` value is not preserved, so a later
        ``inherit = True`` cannot restore it.
        """
        return self._effectLst is None

    @inherit.setter
    def inherit(self, value: bool | None) -> None:
        if bool(value):
            self._element._remove_effectLst()
        else:
            self._element.get_or_add_effectLst()
            # -- also zero any sibling `p:style/a:effectRef/@idx` so that a theme-level
            # -- shadow (inherited via effectRef) does not leak through the empty
            # -- `a:effectLst`. See issue #446.
            self._zero_sibling_effectRef_idx()

    def _zero_sibling_effectRef_idx(self) -> None:
        """Set `../p:style/a:effectRef/@idx` to ``"0"`` when present.

        The `p:style/a:effectRef` element is a sibling of `p:spPr` on `p:sp` and
        `p:cxnSp` shape elements. A `p:grpSpPr` has no sibling `p:style`, so this is a
        no-op there.
        """
        parent = self._element.getparent()
        if parent is None:
            return
        # -- xpath rather than a typed descriptor so we don't need to add `p:style` to
        # -- every shape schema; `p:style` is optional and its shape is simple.
        effectRefs = parent.xpath("./p:style/a:effectRef")
        for effectRef in effectRefs:
            effectRef.set("idx", "0")

    # -- outer-shadow specific read/write properties --------------------------

    @property
    def blur_radius(self) -> Length | None:
        """Length of the outer-shadow blur radius, or None when no `a:outerShdw` is defined.

        Corresponds to `a:outerShdw@blurRad`. Assigning a |Length| value creates the
        `a:outerShdw` (and `a:effectLst` if necessary) and sets `blurRad`.
        """
        outer = cast("CT_OuterShadowEffect | None", self._child)
        if outer is None:
            return None
        return outer.blurRad

    @blur_radius.setter
    def blur_radius(self, value: Length | int | None) -> None:
        if value is None:
            outer = cast("CT_OuterShadowEffect | None", self._child)
            if outer is None:
                return
            outer.blurRad = Emu(0)
            return
        outer = cast("CT_OuterShadowEffect", self._get_or_add_child())
        outer.blurRad = Emu(value)

    @property
    def distance(self) -> Length | None:
        """Offset distance of the outer shadow, or None when no `a:outerShdw` is defined.

        Corresponds to `a:outerShdw@dist`.
        """
        outer = cast("CT_OuterShadowEffect | None", self._child)
        if outer is None:
            return None
        return outer.dist

    @distance.setter
    def distance(self, value: Length | int | None) -> None:
        if value is None:
            outer = cast("CT_OuterShadowEffect | None", self._child)
            if outer is None:
                return
            outer.dist = Emu(0)
            return
        outer = cast("CT_OuterShadowEffect", self._get_or_add_child())
        outer.dist = Emu(value)

    @property
    def direction(self) -> float | None:
        """Direction (in degrees) of the outer shadow, or None when no shadow is defined.

        Corresponds to `a:outerShdw@dir`. Range is 0 - 359.999... degrees.
        """
        outer = cast("CT_OuterShadowEffect | None", self._child)
        if outer is None:
            return None
        return outer.dir

    @direction.setter
    def direction(self, value: float | None) -> None:
        if value is None:
            outer = cast("CT_OuterShadowEffect | None", self._child)
            if outer is None:
                return
            outer.dir = 0.0
            return
        outer = cast("CT_OuterShadowEffect", self._get_or_add_child())
        outer.dir = value

    @lazyproperty
    def color(self) -> ColorFormat:
        """|ColorFormat| object providing access to the shadow color.

        The `a:outerShdw` (and `a:effectLst`) are created on first access to satisfy the
        required color-choice child group; the initial color is black
        (`<a:srgbClr val="000000"/>`).
        """
        outer = cast("CT_OuterShadowEffect", self._get_or_add_child())
        if outer.eg_colorChoice is None:
            from pptx.oxml import parse_xml

            outer.append(parse_xml(_default_srgb("a:srgbClr")))
        return ColorFormat.from_colorchoice_parent(outer)


class GlowFormat(_EffectChildFormat):
    """Provides access to glow-effect settings (`a:glow`) on a shape-like element."""

    _child_tag = "glow"

    @property
    def size(self) -> Length | None:
        """Length of the glow radius, or None when no `a:glow` is defined.

        Corresponds to `a:glow@rad`.
        """
        glow = cast("CT_GlowEffect | None", self._child)
        if glow is None:
            return None
        return glow.rad

    @size.setter
    def size(self, value: Length | int | None) -> None:
        if value is None:
            glow = cast("CT_GlowEffect | None", self._child)
            if glow is None:
                return
            glow.rad = Emu(0)
            return
        glow = cast("CT_GlowEffect", self._get_or_add_child())
        glow.rad = Emu(value)

    @lazyproperty
    def color(self) -> ColorFormat:
        """|ColorFormat| object providing access to the glow color.

        The `a:glow` (and `a:effectLst`) are created on first access to satisfy the
        required color-choice child group; initial color is black.
        """
        glow = cast("CT_GlowEffect", self._get_or_add_child())
        if glow.eg_colorChoice is None:
            from pptx.oxml import parse_xml

            glow.append(parse_xml(_default_srgb("a:srgbClr")))
        return ColorFormat.from_colorchoice_parent(glow)


class ReflectionFormat(_EffectChildFormat):
    """Provides access to reflection-effect settings (`a:reflection`)."""

    _child_tag = "reflection"

    @property
    def blur_radius(self) -> Length | None:
        """Length of the reflection blur radius, or None when no reflection is defined."""
        refl = cast("CT_ReflectionEffect | None", self._child)
        if refl is None:
            return None
        return refl.blurRad

    @blur_radius.setter
    def blur_radius(self, value: Length | int | None) -> None:
        if value is None:
            refl = cast("CT_ReflectionEffect | None", self._child)
            if refl is None:
                return
            refl.blurRad = Emu(0)
            return
        refl = cast("CT_ReflectionEffect", self._get_or_add_child())
        refl.blurRad = Emu(value)

    @property
    def distance(self) -> Length | None:
        """Offset distance of the reflection, or None when no reflection is defined."""
        refl = cast("CT_ReflectionEffect | None", self._child)
        if refl is None:
            return None
        return refl.dist

    @distance.setter
    def distance(self, value: Length | int | None) -> None:
        if value is None:
            refl = cast("CT_ReflectionEffect | None", self._child)
            if refl is None:
                return
            refl.dist = Emu(0)
            return
        refl = cast("CT_ReflectionEffect", self._get_or_add_child())
        refl.dist = Emu(value)


class SoftEdgeFormat(_EffectChildFormat):
    """Provides access to soft-edge-effect settings (`a:softEdge`)."""

    _child_tag = "softEdge"

    @property
    def size(self) -> Length | None:
        """Length of the soft-edge radius, or None when no soft-edge is defined.

        Corresponds to `a:softEdge@rad`.
        """
        soft = cast("CT_SoftEdgesEffect | None", self._child)
        if soft is None:
            return None
        return soft.rad

    @size.setter
    def size(self, value: Length | int | None) -> None:
        # -- `rad` is required on `a:softEdge`; callers should remove the effect by
        # -- assigning `soft_edge.inherit = True` rather than by passing None.
        if value is None:
            raise ValueError(
                "soft-edge size is required; use soft_edge.inherit = True to remove the effect"
            )
        soft = cast("CT_SoftEdgesEffect", self._get_or_add_child())
        soft.rad = Emu(value)
