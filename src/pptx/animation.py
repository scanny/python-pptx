"""High-level shape-animation API.

Issue #102 MVP. Provides :class:`AnimationEffect`, a small read/write
proxy over the animation XML a slide's ``p:timing/p:tnLst`` tree binds
to a single shape, plus the helpers that
:meth:`pptx.shapes.base.BaseShape.animation` /
:meth:`pptx.shapes.base.BaseShape.set_animation` use to walk and mutate
that tree.

Scope
-----

The MVP recognizes and emits five animation presets — :attr:`APPEAR`,
:attr:`FADE_IN`, :attr:`FLY_IN`, :attr:`PULSE`, and :attr:`FADE_OUT`.
Two triggers are supported: :attr:`ON_CLICK` (the default) and
:attr:`AFTER_PREVIOUS`. A non-negative millisecond ``delay`` can be
specified and is written as ``p:cTn/p:stCondLst/p:cond/@delay``.

Deferred (out of MVP scope)
---------------------------

The following are **not** implemented and are reserved for downstream
items. Attempting to author them raises :class:`NotImplementedError`;
reading a shape whose animation XML uses them returns
``AnimationEffect`` with :attr:`MSO_ANIMATION_TYPE.NONE` so callers can
detect the unrecognized case without an exception.

* Motion-path effects (``p:animMotion``) — downstream #264.
* ``p14:morph`` (MORPH transition) — downstream #942.
* Triggers other than ``onClick`` / ``afterEffect`` — ``withEffect`` /
  ``onMouseOver`` / ``onMouseOut`` / ``onDblClick`` / interactive
  sequences — downstream #264.
* Color-emphasis / rotation / scale / rotation-path effects
  (``p:animClr``, ``p:animRot``, ``p:animScale``) — downstream #1106.
* Build-list (``p:bldLst``) authoring — downstream #256.
* Modifying an existing effect's type in-place (use
  :meth:`BaseShape.set_animation` which replaces the existing effect).

Authored XML layout
-------------------

Calling ``shape.set_animation(MSO_ANIMATION_TYPE.FADE_IN)`` for the
first time on a slide produces (new text added is marked with ``+``)::

    +<p:timing>
    +  <p:tnLst>
    +    <p:par>                     <!-- tmRoot -->
    +      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
    +        <p:childTnLst>
    +          <p:seq concurrent="1" nextAc="seek">
    +            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
    +              <p:childTnLst>
    +                <p:par>         <!-- clickEffect -->
    +                  <p:cTn id="3" presetID="10" presetClass="entr"
    +                         presetSubtype="0" fill="hold" nodeType="clickEffect">
    +                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    +                    <p:childTnLst>
    +                      <p:par>   <!-- withGroup parent -->
    +                        <p:cTn id="4" fill="hold">
    +                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    +                          <p:childTnLst>
    +                            <p:par>
    +                              <p:cTn id="5" presetID="10" presetClass="entr"
    +                                     presetSubtype="0" fill="hold">
    +                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    +                                <p:childTnLst>
    +                                  <p:animEffect transition="in" filter="fade">
    +                                    <p:cBhvr>
    +                                      <p:cTn id="6" dur="500"/>
    +                                      <p:tgtEl><p:spTgt spid="N"/></p:tgtEl>
    +                                    </p:cBhvr>
    +                                  </p:animEffect>
    +                                </p:childTnLst>
    +                              </p:cTn>
    +                            </p:par>
    +                          </p:childTnLst>
    +                        </p:cTn>
    +                      </p:par>
    +                    </p:childTnLst>
    +                  </p:cTn>
    +                  <p:prevCondLst>...</p:prevCondLst>
    +                  <p:nextCondLst>...</p:nextCondLst>
    +                </p:par>
    +              </p:childTnLst>
    +            </p:cTn>
    +          </p:seq>
    +        </p:childTnLst>
    +      </p:cTn>
    +    </p:par>
    +  </p:tnLst>
    +</p:timing>

This is a deliberately shallow subset of the full structure PowerPoint
emits. It is sufficient for PowerPoint to correctly play the five MVP
presets. Complex effects (motion paths, grouped builds, iterated
animations) need a deeper tree that downstream issues will author.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE
from pptx.oxml.ns import nsdecls
from pptx.shared import ElementProxy

if TYPE_CHECKING:
    from pptx.oxml.animation import (
        CT_TLAnimateEffectBehavior,
        CT_TLCommonBehaviorData,
        CT_TLSetBehavior,
    )
    from pptx.oxml.shapes.shared import BaseShapeElement
    from pptx.oxml.slide import CT_Slide
    from pptx.oxml.timing import CT_TLCommonTimeNodeData


# -- Preset map: (type, nodeType) -> (presetID, presetClass, presetSubtype) --
# -- presetSubtype is 0 for most presets; FLY_IN uses 4 ("from bottom"). --
_PRESET_MAP: dict[MSO_ANIMATION_TYPE, tuple[int, str, int]] = {
    MSO_ANIMATION_TYPE.APPEAR: (1, "entr", 0),
    MSO_ANIMATION_TYPE.FADE_IN: (10, "entr", 0),
    MSO_ANIMATION_TYPE.FLY_IN: (2, "entr", 4),
    MSO_ANIMATION_TYPE.PULSE: (1, "emph", 0),
    MSO_ANIMATION_TYPE.FADE_OUT: (10, "exit", 0),
}


def _preset_to_type(
    preset_id: int | None, preset_class: str | None
) -> MSO_ANIMATION_TYPE:
    """Reverse-lookup of ``(presetID, presetClass)`` to an MVP preset.

    Returns :attr:`MSO_ANIMATION_TYPE.NONE` when no MVP preset matches —
    the effect XML round-trips, but the high-level API treats it as
    unrecognized.
    """
    if preset_id is None or preset_class is None:
        return MSO_ANIMATION_TYPE.NONE
    for t, (pid, pcls, _psub) in _PRESET_MAP.items():
        if pid == preset_id and pcls == preset_class:
            return t
    return MSO_ANIMATION_TYPE.NONE


class AnimationEffect(ElementProxy):
    """Proxy for a single animation effect bound to a shape.

    Returned from :meth:`pptx.shapes.base.BaseShape.animation` when the
    slide's ``p:timing`` tree contains an effect targeting the shape
    (matched by ``p:spTgt/@spid``).

    The proxy wraps the ``p:par`` element that represents the effect's
    "click effect" container (``nodeType="clickEffect"`` or
    ``nodeType="afterEffect"``) — the node PowerPoint uses as the unit
    of "one animation step" in the main sequence. Reading the
    properties walks down to the inner behavior element
    (``p:set`` / ``p:animEffect`` / ``p:anim``) to discover the preset.

    Writing animations is done via
    :meth:`pptx.shapes.base.BaseShape.set_animation`, which either
    replaces the existing effect or appends a new one. The proxy itself
    is read-only in the MVP; this keeps the semantics of "authored by
    set_animation, introspected by .animation" unambiguous and avoids
    shipping half-implemented mutators.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, par_elm):  # par_elm: CT_TLTimeNodeParallel
        super().__init__(par_elm)
        self._par = par_elm

    @property
    def shape_id(self) -> int | None:
        """Target shape's ``cNvPr/@id``, or ``None`` when unresolvable.

        Walks the effect's descendants for the first ``p:spTgt/@spid``
        and returns it. In practice every MVP-authored effect has
        exactly one ``p:spTgt`` descendant; effects with multiple
        targets (not authored by python-pptx) return the first.

        .. versionadded:: 2026.05.0
        """
        spids = self._par.xpath(".//p:spTgt/@spid")
        if not spids:
            return None
        return int(spids[0])

    @property
    def type(self) -> MSO_ANIMATION_TYPE:
        """The :class:`MSO_ANIMATION_TYPE` preset this effect represents.

        Returns :attr:`MSO_ANIMATION_TYPE.NONE` when the effect's
        ``presetID`` / ``presetClass`` attributes are absent or do not
        match one of the five MVP presets (e.g. a PowerPoint-authored
        motion-path effect).

        .. versionadded:: 2026.05.0
        """
        # -- presetID / presetClass live on the outer clickEffect's p:cTn --
        cTn = self._par.cTn
        preset_id = cTn.presetID
        preset_class = cTn.presetClass
        preset = _preset_to_type(preset_id, preset_class)
        if preset is not MSO_ANIMATION_TYPE.NONE:
            return preset
        # -- some decks carry the preset only on the innermost p:cTn; walk down --
        for inner_cTn in self._par.xpath(".//p:cTn"):
            pid = inner_cTn.get("presetID")
            pcls = inner_cTn.get("presetClass")
            if pid is None or pcls is None:
                continue
            preset = _preset_to_type(int(pid), pcls)
            if preset is not MSO_ANIMATION_TYPE.NONE:
                return preset
        return MSO_ANIMATION_TYPE.NONE

    @property
    def trigger(self) -> MSO_ANIMATION_TRIGGER:
        """:class:`MSO_ANIMATION_TRIGGER` member for this effect's trigger.

        Inferred from the effect's ``p:cTn/@nodeType``:

        * ``clickEffect`` -> :attr:`ON_CLICK`
        * ``afterEffect`` -> :attr:`AFTER_PREVIOUS`
        * anything else -> :attr:`ON_CLICK` (the safe default — the
          effect will still play when the slide is advanced).

        .. versionadded:: 2026.05.0
        """
        node_type = self._par.cTn.nodeType
        if node_type == "afterEffect":
            return MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS
        return MSO_ANIMATION_TRIGGER.ON_CLICK

    @property
    def delay(self) -> int:
        """Effect start delay in milliseconds.

        Reads the first ``p:cond/@delay`` under the effect's
        ``p:cTn/p:stCondLst``. Returns ``0`` when the attribute is
        missing or absent. The ``"indefinite"`` token (used by the
        main-sequence root cTn, not by individual effects) is returned
        as ``0`` in the MVP — callers wanting the raw token should
        inspect the XML directly via :attr:`pptx.slide.Slide.timing_xml`.

        .. versionadded:: 2026.05.0
        """
        stCondLst = self._par.cTn.stCondLst
        if stCondLst is None:
            return 0
        cond_elms = stCondLst.findall(_qn("p:cond"))
        if not cond_elms:
            return 0
        delay_str = cond_elms[0].get("delay")
        if delay_str is None or delay_str == "indefinite":
            return 0
        try:
            return int(delay_str)
        except ValueError:  # pragma: no cover - defensive
            return 0


# -- helpers ----------------------------------------------------------------


def _qn(nsptag: str) -> str:
    """Return Clark-notation qualified name for `nsptag` (e.g. ``p:cond``)."""
    from pptx.oxml.ns import qn

    return qn(nsptag)


def _find_effect_par_for_spid(sld, spid: int):  # sld: CT_Slide
    """Return the "click-effect" ``p:par`` targeting shape `spid`, or ``None``.

    Walks the slide's ``p:timing`` subtree looking for a ``p:par`` whose
    ``p:cTn/@nodeType`` is ``clickEffect`` or ``afterEffect`` and whose
    descendants include a ``p:spTgt`` with ``@spid`` matching. Only the
    first match is returned (the MVP authors one effect per shape).
    """
    timing = sld.timing
    if timing is None:
        return None
    xpath_expr = (
        ".//p:par[p:cTn[@nodeType='clickEffect' or @nodeType='afterEffect']]"
        "[.//p:spTgt[@spid='%d']]" % spid
    )
    matches = timing.xpath(xpath_expr)
    if not matches:
        return None
    return matches[0]


def _remove_effect_par(sld, par_elm) -> None:  # sld: CT_Slide
    """Remove the given effect ``p:par`` element from the main sequence."""
    parent = par_elm.getparent()
    if parent is not None:
        parent.remove(par_elm)


def _ensure_main_sequence(sld) -> "CT_TLCommonTimeNodeData":
    """Return the ``p:cTn`` of the slide's main-sequence ``p:seq``.

    Creates the full tmRoot / mainSeq skeleton lazily if the slide has
    no ``p:timing`` element, or if an existing ``p:timing`` tree has no
    ``mainSeq`` (e.g. a slide with only video-playback timing — that
    case only adds a ``mainSeq`` and leaves the video-playback tree
    alone).
    """
    from pptx.oxml import parse_xml

    timing = sld.get_or_add_timing()
    # -- Find an existing mainSeq p:seq, if any --
    mainSeq_seqs = timing.xpath(".//p:seq[p:cTn[@nodeType='mainSeq']]")
    if mainSeq_seqs:
        return mainSeq_seqs[0].cTn

    # -- Find the tmRoot p:par (it's the one child with nodeType="tmRoot") --
    tmRoot_pars = timing.xpath("./p:tnLst/p:par[p:cTn[@nodeType='tmRoot']]")
    if tmRoot_pars:
        tmRoot_par = tmRoot_pars[0]
    else:
        # -- Add a fresh tmRoot --
        tnLst = timing.get_or_add_tnLst()
        tmRoot_par_xml = (
            "<p:par %s>\n"
            '  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">\n'
            "    <p:childTnLst/>\n"
            "  </p:cTn>\n"
            "</p:par>" % nsdecls("p")
        )
        tmRoot_par = parse_xml(tmRoot_par_xml)
        tnLst.insert(0, tmRoot_par)

    # -- Append a mainSeq p:seq under tmRoot/cTn/childTnLst --
    tmRoot_cTn = tmRoot_par.cTn
    childTnLst = tmRoot_cTn.get_or_add_childTnLst()
    next_id = _next_cTn_id(sld)
    mainSeq_seq_xml = (
        '<p:seq %s concurrent="1" nextAc="seek">\n'
        '  <p:cTn id="%d" dur="indefinite" nodeType="mainSeq">\n'
        "    <p:childTnLst/>\n"
        "  </p:cTn>\n"
        "  <p:prevCondLst>\n"
        '    <p:cond evt="onPrev" delay="0">\n'
        "      <p:tgtEl><p:sldTgt/></p:tgtEl>\n"
        "    </p:cond>\n"
        "  </p:prevCondLst>\n"
        "  <p:nextCondLst>\n"
        '    <p:cond evt="onNext" delay="0">\n'
        "      <p:tgtEl><p:sldTgt/></p:tgtEl>\n"
        "    </p:cond>\n"
        "  </p:nextCondLst>\n"
        "</p:seq>" % (nsdecls("p"), next_id)
    )
    mainSeq_seq = parse_xml(mainSeq_seq_xml)
    childTnLst.append(mainSeq_seq)
    return mainSeq_seq.cTn


def _next_cTn_id(sld) -> int:
    """Return 1 + max existing ``p:cTn/@id`` under the slide's ``p:timing``.

    Mirrors :meth:`CT_TimeNodeList._next_cTn_id` but operates on the
    slide root so the allocated id is globally unique within the
    presentation-timing tree (not just within one ``p:tnLst``).
    """
    ids = sld.xpath("./p:timing//p:cTn/@id")
    int_ids = [int(s) for s in ids]
    return max(int_ids) + 1 if int_ids else 1


def _build_click_effect_par_xml(
    preset_id: int,
    preset_class: str,
    preset_subtype: int,
    node_type: str,
    delay: int,
    id_click: int,
    id_with: int,
    id_inner: int,
    id_behavior_cTn: int,
    behavior_xml: str,
) -> str:
    """Return the XML for a single click-effect ``p:par`` element.

    The ``p:par`` is the outer "one animation step" node for the main
    sequence. The inner ``behavior_xml`` fragment provides the actual
    behavior (``p:set``, ``p:animEffect``, etc.) and is inserted as a
    sibling under the innermost ``childTnLst``.
    """
    return (
        "<p:par %s>\n"
        '  <p:cTn id="%d" presetID="%d" presetClass="%s" presetSubtype="%d"'
        ' fill="hold" nodeType="%s">\n'
        "    <p:stCondLst>\n"
        '      <p:cond delay="%d"/>\n'
        "    </p:stCondLst>\n"
        "    <p:childTnLst>\n"
        "      <p:par>\n"
        '        <p:cTn id="%d" fill="hold">\n'
        "          <p:stCondLst>\n"
        '            <p:cond delay="0"/>\n'
        "          </p:stCondLst>\n"
        "          <p:childTnLst>\n"
        "            <p:par>\n"
        '              <p:cTn id="%d" presetID="%d" presetClass="%s"'
        ' presetSubtype="%d" fill="hold" grpId="0" nodeType="withEffect">\n'
        "                <p:stCondLst>\n"
        '                  <p:cond delay="0"/>\n'
        "                </p:stCondLst>\n"
        "                <p:childTnLst>\n"
        "%s"
        "                </p:childTnLst>\n"
        "              </p:cTn>\n"
        "            </p:par>\n"
        "          </p:childTnLst>\n"
        "        </p:cTn>\n"
        "      </p:par>\n"
        "    </p:childTnLst>\n"
        "  </p:cTn>\n"
        "</p:par>"
        % (
            nsdecls("p"),
            id_click,
            preset_id,
            preset_class,
            preset_subtype,
            node_type,
            delay,
            id_with,
            id_inner,
            preset_id,
            preset_class,
            preset_subtype,
            behavior_xml,
        )
    )


def _build_animEffect_xml(
    cTn_id: int,
    spid: int,
    transition: str,
    filter_str: str,
    dur: int,
) -> str:
    """Return the ``p:animEffect`` behavior XML fragment."""
    return (
        '                  <p:animEffect transition="%s" filter="%s">\n'
        "                    <p:cBhvr>\n"
        '                      <p:cTn id="%d" dur="%d"/>\n'
        "                      <p:tgtEl>\n"
        '                        <p:spTgt spid="%d"/>\n'
        "                      </p:tgtEl>\n"
        "                    </p:cBhvr>\n"
        "                  </p:animEffect>\n"
        % (transition, filter_str, cTn_id, dur, spid)
    )


def _build_set_xml(cTn_id: int, spid: int) -> str:
    """Return the ``p:set`` behavior XML fragment ("appear" preset)."""
    return (
        '                  <p:set>\n'
        "                    <p:cBhvr>\n"
        '                      <p:cTn id="%d" dur="1" fill="hold">\n'
        '                        <p:stCondLst>\n'
        '                          <p:cond delay="0"/>\n'
        '                        </p:stCondLst>\n'
        '                      </p:cTn>\n'
        "                      <p:tgtEl>\n"
        '                        <p:spTgt spid="%d"/>\n'
        "                      </p:tgtEl>\n"
        "                      <p:attrNameLst>\n"
        "                        <p:attrName>style.visibility</p:attrName>\n"
        "                      </p:attrNameLst>\n"
        "                    </p:cBhvr>\n"
        "                    <p:to>\n"
        '                      <p:strVal val="visible"/>\n'
        "                    </p:to>\n"
        "                  </p:set>\n"
        % (cTn_id, spid)
    )


def _set_animation(
    sld,
    spid: int,
    effect_type: MSO_ANIMATION_TYPE,
    trigger: MSO_ANIMATION_TRIGGER,
    delay: int,
) -> "AnimationEffect":
    """Remove any prior effect on `spid` and append a new one.

    Called from :meth:`pptx.shapes.base.BaseShape.set_animation`. See
    that method's docstring for the public contract; this helper
    performs the XML surgery.
    """
    from pptx.oxml import parse_xml

    if effect_type is MSO_ANIMATION_TYPE.NONE:
        raise ValueError(
            "effect_type must be one of the MVP presets (APPEAR / FADE_IN / "
            "FLY_IN / PULSE / FADE_OUT), not MSO_ANIMATION_TYPE.NONE"
        )
    if effect_type not in _PRESET_MAP:
        raise NotImplementedError(
            "MSO_ANIMATION_TYPE.%s is not supported in the MVP preset set. "
            "See docs/dev/analysis/f8-animations-transitions.rst for the "
            "downstream items that will add more presets." % effect_type.name
        )
    MSO_ANIMATION_TRIGGER.validate(trigger)
    if not isinstance(delay, int) or delay < 0:
        raise ValueError(
            "delay must be a non-negative int (milliseconds), got %r" % (delay,)
        )

    # -- Remove any existing effect targeting this shape --
    existing = _find_effect_par_for_spid(sld, spid)
    if existing is not None:
        _remove_effect_par(sld, existing)

    mainSeq_cTn = _ensure_main_sequence(sld)
    mainSeq_childTnLst = mainSeq_cTn.get_or_add_childTnLst()

    preset_id, preset_class, preset_subtype = _PRESET_MAP[effect_type]
    # -- xml_value is always non-None for MVP trigger members; cast for type-checker --
    node_type = trigger.xml_value or "clickEffect"

    # -- Allocate four fresh cTn ids (outer click, with group, inner, behavior) --
    id_click = _next_cTn_id(sld)
    id_with = id_click + 1
    id_inner = id_click + 2
    id_behavior_cTn = id_click + 3

    # -- Build the behavior XML appropriate for the preset --
    if effect_type is MSO_ANIMATION_TYPE.APPEAR:
        behavior_xml = _build_set_xml(id_behavior_cTn, spid)
    elif effect_type is MSO_ANIMATION_TYPE.FADE_IN:
        behavior_xml = _build_animEffect_xml(
            id_behavior_cTn, spid, "in", "fade", 500
        )
    elif effect_type is MSO_ANIMATION_TYPE.FLY_IN:
        behavior_xml = _build_animEffect_xml(
            id_behavior_cTn, spid, "in", "slide(fromBottom)", 500
        )
    elif effect_type is MSO_ANIMATION_TYPE.PULSE:
        behavior_xml = _build_animEffect_xml(
            id_behavior_cTn, spid, "none", "fade", 500
        )
    elif effect_type is MSO_ANIMATION_TYPE.FADE_OUT:
        behavior_xml = _build_animEffect_xml(
            id_behavior_cTn, spid, "out", "fade", 500
        )
    else:  # pragma: no cover - guarded by _PRESET_MAP check above
        raise NotImplementedError(effect_type)

    par_xml = _build_click_effect_par_xml(
        preset_id,
        preset_class,
        preset_subtype,
        node_type,
        delay,
        id_click,
        id_with,
        id_inner,
        id_behavior_cTn,
        behavior_xml,
    )
    par_elm = parse_xml(par_xml)
    mainSeq_childTnLst.append(par_elm)
    return AnimationEffect(par_elm)
