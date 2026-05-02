"""lxml custom element classes for DrawingML effect-related XML elements.

Grammar reference: ECMA-376 Part 1 §20.1.8 `CT_EffectList`, `CT_BlurEffect`, `CT_GlowEffect`,
`CT_InnerShadowEffect`, `CT_OuterShadowEffect`, `CT_PresetShadowEffect`, `CT_ReflectionEffect`,
`CT_SoftEdgesEffect` (see `spec/ISO-IEC-29500-1/schemas/xsd/dml-main.xsd`).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_Percentage,
    ST_PositiveCoordinate,
    ST_PositiveFixedAngle,
    ST_PositiveFixedPercentage,
    ST_PresetShadowVal,
    XsdBoolean,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)

if TYPE_CHECKING:
    from pptx.util import Length


class CT_EffectList(BaseOxmlElement):
    """`a:effectLst` custom element class.

    Sequence container for shadow / glow / reflection / soft-edge (plus the rarely used blur and
    fill-overlay) effects, as defined by ECMA-376 `CT_EffectList`.
    """

    get_or_add_glow: Callable[[], "CT_GlowEffect"]
    get_or_add_innerShdw: Callable[[], "CT_InnerShadowEffect"]
    get_or_add_outerShdw: Callable[[], "CT_OuterShadowEffect"]
    get_or_add_reflection: Callable[[], "CT_ReflectionEffect"]
    get_or_add_softEdge: Callable[[], "CT_SoftEdgesEffect"]
    _remove_glow: Callable[[], None]
    _remove_innerShdw: Callable[[], None]
    _remove_outerShdw: Callable[[], None]
    _remove_reflection: Callable[[], None]
    _remove_softEdge: Callable[[], None]

    _tag_seq = (
        "a:blur",
        "a:fillOverlay",
        "a:glow",
        "a:innerShdw",
        "a:outerShdw",
        "a:prstShdw",
        "a:reflection",
        "a:softEdge",
    )
    blur: "CT_BlurEffect | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:blur", successors=_tag_seq[1:]
    )
    glow: "CT_GlowEffect | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:glow", successors=_tag_seq[3:]
    )
    innerShdw: "CT_InnerShadowEffect | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:innerShdw", successors=_tag_seq[4:]
    )
    outerShdw: "CT_OuterShadowEffect | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:outerShdw", successors=_tag_seq[5:]
    )
    prstShdw: "CT_PresetShadowEffect | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:prstShdw", successors=_tag_seq[6:]
    )
    reflection: "CT_ReflectionEffect | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:reflection", successors=_tag_seq[7:]
    )
    softEdge: "CT_SoftEdgesEffect | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:softEdge", successors=_tag_seq[8:]
    )
    del _tag_seq

    @classmethod
    def new_effectLst(cls) -> "CT_EffectList":
        """Return newly-created "loose" empty `a:effectLst` element."""
        from typing import cast

        from pptx.oxml import parse_xml

        return cast("CT_EffectList", parse_xml("<a:effectLst %s/>" % nsdecls("a")))


class CT_BlurEffect(BaseOxmlElement):
    """`a:blur` element."""

    rad: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rad", ST_PositiveCoordinate, default=0
    )
    grow: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "grow", XsdBoolean, default=True
    )


class CT_GlowEffect(BaseOxmlElement):
    """`a:glow` element."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    rad: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rad", ST_PositiveCoordinate, default=0
    )


class CT_InnerShadowEffect(BaseOxmlElement):
    """`a:innerShdw` element."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    blurRad: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "blurRad", ST_PositiveCoordinate, default=0
    )
    dist: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate, default=0
    )
    dir: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_PositiveFixedAngle, default=0.0
    )


class CT_OuterShadowEffect(BaseOxmlElement):
    """`a:outerShdw` element."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    blurRad: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "blurRad", ST_PositiveCoordinate, default=0
    )
    dist: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate, default=0
    )
    dir: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_PositiveFixedAngle, default=0.0
    )
    rotWithShape: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rotWithShape", XsdBoolean, default=True
    )


class CT_PresetShadowEffect(BaseOxmlElement):
    """`a:prstShdw` element."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    prst: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "prst", ST_PresetShadowVal
    )
    dist: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate, default=0
    )
    dir: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_PositiveFixedAngle, default=0.0
    )


class CT_ReflectionEffect(BaseOxmlElement):
    """`a:reflection` element."""

    blurRad: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "blurRad", ST_PositiveCoordinate, default=0
    )
    stA: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "stA", ST_PositiveFixedPercentage, default=1.0
    )
    stPos: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "stPos", ST_PositiveFixedPercentage, default=0.0
    )
    endA: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "endA", ST_PositiveFixedPercentage, default=0.0
    )
    endPos: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "endPos", ST_PositiveFixedPercentage, default=1.0
    )
    dist: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate, default=0
    )
    dir: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_PositiveFixedAngle, default=0.0
    )
    fadeDir: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "fadeDir", ST_PositiveFixedAngle, default=90.0
    )
    sx: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sx", ST_Percentage, default=1.0
    )
    sy: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sy", ST_Percentage, default=1.0
    )
    rotWithShape: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rotWithShape", XsdBoolean, default=True
    )


class CT_SoftEdgesEffect(BaseOxmlElement):
    """`a:softEdge` element."""

    rad: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "rad", ST_PositiveCoordinate
    )
