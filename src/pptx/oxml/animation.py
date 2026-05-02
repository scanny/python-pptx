"""Custom element classes for shape animation behaviors.

Issue #102 MVP. Extends :mod:`pptx.oxml.timing` (the Foundation F8 shell)
with typed classes for the small number of ``p:tnLst`` behavior children
needed to author entrance / emphasis / exit effects on shapes:

* ``CT_TLAnimateBehavior`` (``p:anim``) — animated-value behaviors
  (round-tripped; no MVP authoring helper).
* ``CT_TLAnimateEffectBehavior`` (``p:animEffect``) — the fade / fly-in
  / pulse / fade-out presets the MVP emits.
* ``CT_TLSetBehavior`` (``p:set``) — the "appear" preset (sets the
  shape's ``style.visibility`` to ``visible`` at t=0).
* ``CT_TLCommonBehaviorData`` (``p:cBhvr``) — common behavior parent
  carrying the ``p:cTn`` + ``p:tgtEl`` children every behavior uses.
* ``CT_TLTimeTargetElement`` / ``CT_TLShapeTargetElement``
  (``p:tgtEl`` / ``p:spTgt``) — the shape-id binding between a behavior
  and the shape it drives.
* ``CT_TLTimeCondition`` / ``CT_TLTimeConditionList``
  (``p:cond`` / ``p:stCondLst``) — the start-condition plumbing used to
  represent triggers (``onClick`` / ``onNext`` / ``onPrev``) and delays.

The MVP only surfaces the attributes that the high-level
:class:`pptx.animation.AnimationEffect` proxy reads or writes; the rest
round-trip through lxml by default. Complex behavior types
(``p:animMotion`` motion paths, ``p:animClr`` color emphasis,
``p:animRot`` / ``p:animScale``, ``p:cmd`` verb invocations) are
deliberately omitted and reserved for downstream issues #264 / #1106.

See ``docs/dev/analysis/f8-animations-transitions.rst`` for the
downstream-items map.
"""

from __future__ import annotations

from pptx.oxml.simpletypes import BaseSimpleType, XsdString, XsdUnsignedInt
from pptx.oxml.timing import ST_TLTime
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class ST_TLTriggerEvent(BaseSimpleType):
    """Simple type for ``p:cond/@evt``.

    Mirrors ``ST_TLTriggerEvent`` in ``pml.xsd``. Accepts the eleven
    tokens in the schema (``onClick`` / ``onNext`` / ``onPrev`` / ... );
    python-pptx surfaces only ``onClick``, ``onPrev``, and ``onNext`` at
    the high-level API layer in the MVP. The rest round-trip as-is.
    """

    _VALID = frozenset(
        (
            "onBegin",
            "onEnd",
            "begin",
            "end",
            "onClick",
            "onDblClick",
            "onMouseOver",
            "onMouseOut",
            "onNext",
            "onPrev",
            "onStopAudio",
        )
    )

    @classmethod
    def convert_from_xml(cls, str_value: str) -> str:
        return str_value

    @classmethod
    def convert_to_xml(cls, value) -> str:
        return value

    @classmethod
    def validate(cls, value):
        if not isinstance(value, str) or value not in cls._VALID:
            raise ValueError(
                "ST_TLTriggerEvent must be one of %s, got %r" % (sorted(cls._VALID), value)
            )


class CT_TLShapeTargetElement(BaseOxmlElement):
    """`p:spTgt` element — selects the target shape for a behavior.

    Carries a required ``spid`` attribute holding the target shape's
    ``cNvPr/@id``. The MVP only surfaces ``spid``; the optional
    ``bg`` / ``subSp`` / ``oleChartEl`` / ``txEl`` / ``graphicEl``
    choice children are preserved via lxml round-tripping.
    """

    spid: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "spid", XsdUnsignedInt
    )


class CT_TLTimeTargetElement(BaseOxmlElement):
    """`p:tgtEl` element — wraps the target selector.

    The schema choice (``sldTgt`` / ``sndTgt`` / ``spTgt`` / ``inkTgt``)
    is not modeled as a :class:`ZeroOrOneChoice`; only ``spTgt`` is
    surfaced as a descriptor because that is the target every shape
    animation uses. Other choice children round-trip through lxml.
    """

    _tag_seq = ("p:sldTgt", "p:sndTgt", "p:spTgt", "p:inkTgt")
    spTgt: CT_TLShapeTargetElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:spTgt", successors=_tag_seq[3:]
    )
    del _tag_seq


class CT_TLTimeCondition(BaseOxmlElement):
    """`p:cond` element — a single start/end time condition.

    Used primarily by the MVP to express the "start on click" / "start
    after previous" / "start with delay" conditions on an effect's
    ``p:stCondLst``. Surfaces ``evt`` (``onClick`` / ``onPrev`` / ...) and
    ``delay`` (milliseconds or ``indefinite``) — the nested ``tn`` /
    ``rtn`` / ``tgtEl`` choice children round-trip through lxml.
    """

    evt = OptionalAttribute("evt", ST_TLTriggerEvent)
    delay = OptionalAttribute("delay", ST_TLTime)


class CT_TLTimeConditionList(BaseOxmlElement):
    """`p:stCondLst`, `p:endCondLst`, `p:prevCondLst`, or `p:nextCondLst`.

    A sequence of ``p:cond`` children. The MVP only reads/writes a
    single ``cond`` (the first one) for effect triggers; multiple
    conditions round-trip transparently.
    """


class CT_TLCommonBehaviorData(BaseOxmlElement):
    """`p:cBhvr` element — common behavior wrapper.

    Parent of ``p:cTn`` (the common time-node data carrying ``id``,
    ``dur``, etc.) and ``p:tgtEl`` (the target-shape selector). Every
    shape-bound behavior (``p:set``, ``p:anim``, ``p:animEffect``)
    contains exactly one ``p:cBhvr``.
    """

    _tag_seq = ("p:cTn", "p:tgtEl", "p:attrNameLst")
    # -- cTn is typed via the CT_TLCommonTimeNodeData registration in --
    # -- pptx.oxml.timing; OneAndOnlyOne gives a non-None typed accessor. --
    cTn = OneAndOnlyOne("p:cTn")
    tgtEl: CT_TLTimeTargetElement = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:tgtEl"
    )
    del _tag_seq

    @property
    def spid(self) -> int | None:
        """Target shape's `cNvPr/@id`, or `None` when no ``p:spTgt`` child.

        Convenience read-through to ``p:tgtEl/p:spTgt/@spid``. Returns
        ``None`` if the behavior targets something other than a shape
        (e.g. ``p:sldTgt``), which the MVP does not author itself but
        round-trips.
        """
        spTgt = self.tgtEl.spTgt
        if spTgt is None:
            return None
        return spTgt.spid


class CT_TLSetBehavior(BaseOxmlElement):
    """`p:set` element — "set-value" behavior.

    Used by the "appear" entrance preset to force the shape's
    ``style.visibility`` to ``visible`` at t=0. Carries one ``p:cBhvr``
    and an optional ``p:to`` child giving the value to set.
    """

    _tag_seq = ("p:cBhvr", "p:to")
    cBhvr: CT_TLCommonBehaviorData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cBhvr"
    )
    to = ZeroOrOne("p:to", successors=())
    del _tag_seq


class CT_TLAnimateBehavior(BaseOxmlElement):
    """`p:anim` element — generic animate behavior.

    Animates a numeric or string attribute from one value to another
    across the behavior's duration. The MVP does **not** emit ``p:anim``
    directly (fade / fly-in / pulse are all expressed via
    ``p:animEffect``), but the class is present so decks authored by
    PowerPoint can round-trip through python-pptx unchanged and future
    work can emit ``p:anim`` directly.
    """

    _tag_seq = ("p:cBhvr", "p:tavLst")
    cBhvr: CT_TLCommonBehaviorData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cBhvr"
    )
    tavLst = ZeroOrOne("p:tavLst", successors=())
    del _tag_seq

    calcmode = OptionalAttribute("calcmode", XsdString)
    valueType = OptionalAttribute("valueType", XsdString)
    by_attr = OptionalAttribute("by", XsdString)
    from_attr = OptionalAttribute("from", XsdString)
    to_attr = OptionalAttribute("to", XsdString)


class CT_TLAnimateEffectBehavior(BaseOxmlElement):
    """`p:animEffect` element — filter-based animation behavior.

    The workhorse for the MVP: fade-in, fade-out, fly-in, and pulse are
    all ``p:animEffect`` behaviors that differ only in ``@transition``
    (``in`` / ``out`` / ``none``) and ``@filter`` (e.g. ``fade`` or
    ``slide(fromBottom)``).
    """

    _tag_seq = ("p:cBhvr", "p:progress")
    cBhvr: CT_TLCommonBehaviorData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cBhvr"
    )
    progress = ZeroOrOne("p:progress", successors=())
    del _tag_seq

    transition = OptionalAttribute("transition", XsdString)
    filter_attr = OptionalAttribute("filter", XsdString)
    prLst = OptionalAttribute("prLst", XsdString)
