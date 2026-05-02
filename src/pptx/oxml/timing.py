"""Custom element classes for the slide-timing / animation XML subtree.

Foundation F8 MVP. These classes exist to round-trip the ``p:timing``
subtree when a slide is loaded and saved — they do **not** implement a
high-level animation authoring API. That work is deferred to downstream
items (#102, #264, #861, #1004, #1106, and the media-timing items
#376, #501, #811, #954). See ``docs/dev/analysis/f8-animations-transitions.rst``
for the full map of deferrals and extension points.

Key design notes:

* `CT_SlideTiming`, `CT_TimeNodeList`, `CT_TLTimeNodeParallel`,
  `CT_TLTimeNodeSequence`, and `CT_TLCommonTimeNodeData` are declared
  here even though the element-class *registrations* for
  ``p:timing`` / ``p:tnLst`` / ``p:video`` remain in ``pptx.oxml.slide``
  (for backward compatibility with the existing media-timing plumbing
  in ``CT_Slide``). Callers that want the richer class should still
  import via the stable ``pptx.oxml.timing`` name.
* No ``Choice`` descriptor is declared for every child element of
  ``p:tnLst`` (there are twelve behavior types: ``p:par``, ``p:seq``,
  ``p:excl``, ``p:anim``, ``p:animClr``, ``p:animEffect``,
  ``p:animMotion``, ``p:animRot``, ``p:animScale``, ``p:cmd``,
  ``p:set``, ``p:audio``, ``p:video``). A full typed API for each
  is out of MVP scope; the base `CT_TimeNodeList` preserves whatever
  children are present by virtue of ``BaseOxmlElement`` round-tripping.
  Downstream items can add per-behavior classes incrementally without
  touching this file.
* ``CT_TLCommonTimeNodeData`` exposes just ``id``, ``nodeType``, and
  ``dur`` for MVP — enough for tests and for the timing-xml passthrough
  use case. The remaining 20 attributes specified on the type in
  pml.xsd are preserved by lxml automatically but are not surfaced as
  Python descriptors yet.
* Issue #861 extends ``CT_TLCommonTimeNodeData`` with ``stCondLst`` /
  ``endCondLst`` child descriptors and a convenience ``delay`` property
  that reads/writes the ``p:cond/@delay`` attribute on the first
  ``p:stCondLst/p:cond`` child. PowerPoint canonically emits a single
  ``p:cond`` per ``p:stCondLst`` for click / with-previous / after-previous
  triggers, so the convenience targets that common case. New
  ``CT_TLTimeCondition`` and ``CT_TLTimeConditionList`` classes surface
  the schema types so downstream work on full animation trees (#102,
  #264, #1106) can build on them.
"""

from __future__ import annotations

from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import (
    BaseSimpleType,
    XsdInt,
    XsdString,
    XsdToken,
    XsdUnsignedInt,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OneOrMore,
    OptionalAttribute,
    ZeroOrOne,
)


class ST_TLTime(BaseSimpleType):
    """ST_TLTime simple type from ``pml.xsd``.

    Union of ``xsd:unsignedInt`` and the token ``"indefinite"``. Used for
    ``@dur``, ``@repeatDur``, ``@delay`` and friends on the timing
    elements.

    The Python value is either an ``int`` (milliseconds) or the string
    ``"indefinite"``. Downstream item #861 (read/write access to
    animation delays) will consume this — MVP just needs round-trip.
    """

    @classmethod
    def convert_from_xml(cls, str_value: str):
        if str_value == "indefinite":
            return "indefinite"
        return int(str_value)

    @classmethod
    def convert_to_xml(cls, value):
        if value == "indefinite":
            return "indefinite"
        return str(int(value))

    @classmethod
    def validate(cls, value):
        if value == "indefinite":
            return
        if not isinstance(value, int) or value < 0:
            raise ValueError(
                "ST_TLTime must be a non-negative int (milliseconds) or the string "
                "'indefinite', got %r" % (value,)
            )


class CT_SlideTiming(BaseOxmlElement):
    """`p:timing` element, specifying animations and timed behaviors.

    Root of the slide-timing subtree. Children:

    * `p:tnLst` (0..1) — the time-node list (usually one ``p:par`` with
      ``nodeType="tmRoot"``).
    * `p:bldLst` (0..1) — build list for per-shape build metadata.
      Preserved on round-trip but not surfaced.
    * `p:extLst` (0..1) — extension list; ignored by MVP.
    """

    _tag_seq = ("p:tnLst", "p:bldLst", "p:extLst")
    tnLst = ZeroOrOne("p:tnLst", successors=_tag_seq[1:])
    bldLst = ZeroOrOne("p:bldLst", successors=_tag_seq[2:])
    del _tag_seq


class CT_TimeNodeList(BaseOxmlElement):
    """`p:tnLst` or `p:childTnList` element.

    Holds an ordered list of time-node children. The schema says there
    must be at least one child, but python-pptx tolerates empty lists
    so that synthetic / test fixtures can construct one incrementally.

    The existing :meth:`add_video` helper for media playback is preserved
    (see :mod:`pptx.oxml.slide` for where it's invoked from).
    """

    def add_video(self, shape_id):
        """Add a new `p:video` child element for a movie having *shape_id*.

        Existing behavior, preserved verbatim from the pre-F8
        implementation in ``pptx.oxml.slide``. Media-timing issue #811
        (start-time for add_movie) builds on this.
        """
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        video_xml = (
            "<p:video %s>\n"
            '  <p:cMediaNode vol="80000">\n'
            '    <p:cTn id="%d" fill="hold" display="0">\n'
            "      <p:stCondLst>\n"
            '        <p:cond delay="indefinite"/>\n'
            "      </p:stCondLst>\n"
            "    </p:cTn>\n"
            "    <p:tgtEl>\n"
            '      <p:spTgt spid="%d"/>\n'
            "    </p:tgtEl>\n"
            "  </p:cMediaNode>\n"
            "</p:video>\n" % (nsdecls("p"), self._next_cTn_id, shape_id)
        )
        video = parse_xml(video_xml)
        self.append(video)

    @property
    def _next_cTn_id(self):
        """Next available unique ID (int) for a new `p:cTn` element.

        Collects `@id` from every `p:cTn` reachable from the containing
        `p:sld` document root, including `p:cTn` elements nested inside
        an ``mc:AlternateContent``/``mc:Choice`` (or ``mc:Fallback``)
        wrapper. A plain ``/p:sld/p:timing//p:cTn/@id`` xpath misses
        the wrapped variant and would mint a duplicate id (issue #954).
        """
        cTn_id_strs = self.xpath("/p:sld//p:cTn/@id")
        ids = [int(id_str) for id_str in cTn_id_strs]
        return max(ids) + 1 if ids else 1


class CT_TLCommonTimeNodeData(BaseOxmlElement):
    """`p:cTn` element — common time-node attributes.

    This type carries 24 attributes per pml.xsd. MVP surfaces a growing
    subset: ``id`` (referenced by tests and required by cTn-id
    allocation), ``nodeType`` (``tmRoot`` / ``mainSeq`` / ``clickEffect``
    / ... needed to route animations), ``dur`` (ST_TLTime, used by #861
    / #376), and — added by issue #102 — ``presetID`` / ``presetClass``
    / ``presetSubtype`` so entrance/emphasis/exit presets authored by
    :meth:`pptx.shapes.base.BaseShape.set_animation` have a typed home.
    The remaining 18 attributes round-trip transparently through lxml.

    The ``stCondLst`` / ``childTnLst`` children are also now typed as
    ``ZeroOrOne`` descriptors. This lets #102 read the "delay" on an
    effect's start-condition (``cTn/stCondLst/cond/@delay``) and walk
    the child time-node list that holds each effect's behaviors.
    Other children (``endCondLst``, ``endSync``, ``iterate``,
    ``subTnLst``) continue to round-trip via lxml.

    Issue #861 additionally surfaces the ``p:stCondLst`` and
    ``p:endCondLst`` children plus a convenience :attr:`delay` property
    that reads/writes the start-delay (in ms, or ``"indefinite"``) on the
    first ``p:stCondLst/p:cond`` child. PowerPoint canonically emits a
    single ``p:cond`` inside ``p:stCondLst`` so this convenience covers
    the overwhelming majority of real-world animation trees.

    Issue #256 adds the three preset-selection attributes
    (``presetID`` / ``presetClass`` / ``presetSubtype``) so callers can
    introspect the entrance / exit / emphasis effect carried by an
    effect-level ``p:cTn``. These are integer attributes in pml.xsd and
    absent on non-effect time nodes (e.g. ``tmRoot`` / ``mainSeq``).

    Downstream extension points (see the F8 analysis doc):

    * #861: surface ``endCondLst`` and value-list semantics for
      delay read/write on more than one condition.
    * #102 / #1106: add per-effect presetID / presetClass / presetSubtype
      descriptors so entrance/exit presets can be authored.
    * #264: expose full ``p:cond`` list semantics (multiple conditions,
      ``@evt``, and the ``tgtEl`` / ``tn`` / ``rtn`` choice child) for
      click vs with-previous vs after-previous triggers.
    * #954: wrap in ``mc:AlternateContent`` for PowerPoint 2010+
      extensibility.

    ``stCondLst`` is surfaced here as a ``ZeroOrOne`` descriptor for
    issue #811 (start-time for ``add_movie``). The schema allows
    ``stCondLst``, ``endCondLst``, ``endSync``, and ``iterate`` as
    children of ``p:cTn``; only ``stCondLst`` is typed — the remaining
    three round-trip through lxml.
    """

    _tag_seq = (
        "p:stCondLst",
        "p:endCondLst",
        "p:endSync",
        "p:iterate",
        "p:childTnLst",
        "p:subTnLst",
    )
    stCondLst = ZeroOrOne("p:stCondLst", successors=_tag_seq[1:])
    endCondLst = ZeroOrOne("p:endCondLst", successors=_tag_seq[2:])
    childTnLst = ZeroOrOne("p:childTnLst", successors=_tag_seq[5:])
    del _tag_seq

    id: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdUnsignedInt
    )
    nodeType: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "nodeType", XsdString
    )
    dur = OptionalAttribute("dur", ST_TLTime)
    presetID: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "presetID", XsdInt
    )
    presetClass: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "presetClass", XsdString
    )
    presetSubtype: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "presetSubtype", XsdInt
    )

    @property
    def delay(self):
        """Start-delay of this time node, or ``None`` when unspecified.

        Reads the ``@delay`` attribute of the first ``p:cond`` child of
        ``p:stCondLst``. The value is an ``int`` (milliseconds) or the
        sentinel string ``"indefinite"`` when the animation waits for an
        external trigger (e.g. a click).

        Returns ``None`` when no ``p:stCondLst`` is present, when it has
        no ``p:cond`` child, or when the first ``p:cond`` has no
        ``@delay`` attribute.
        """
        stCondLst = self.stCondLst
        if stCondLst is None:
            return None
        conds = stCondLst.findall(qn("p:cond"))
        if not conds:
            return None
        return conds[0].delay

    @delay.setter
    def delay(self, value):
        """Write the start-delay for this time node.

        Setting ``None`` clears the ``@delay`` attribute on the first
        ``p:cond`` child of ``p:stCondLst`` (but leaves the ``p:cond``
        and ``p:stCondLst`` in place — they may carry ``@evt`` or other
        attributes a caller wants to preserve). If no ``p:stCondLst`` /
        ``p:cond`` exists, setting ``None`` is a no-op.

        Setting an ``int`` (milliseconds, must be non-negative) or the
        string ``"indefinite"`` writes that value, adding a
        ``p:stCondLst`` and a blank ``p:cond`` child as needed.
        """
        if value is None:
            stCondLst = self.stCondLst
            if stCondLst is None:
                return
            conds = stCondLst.findall(qn("p:cond"))
            if not conds:
                return
            conds[0].delay = None
            return
        ST_TLTime.validate(value)
        stCondLst = self.get_or_add_stCondLst()
        conds = stCondLst.findall(qn("p:cond"))
        if not conds:
            cond = stCondLst.add_cond()
        else:
            cond = conds[0]
        cond.delay = value


class CT_TLTimeNodeParallel(BaseOxmlElement):
    """`p:par` element — parallel time node.

    One required ``p:cTn`` child. Usually the root time node is a
    parallel node with ``nodeType="tmRoot"``; the main animation
    sequence is then a child ``p:seq`` of that root ``p:par``.
    """

    cTn: CT_TLCommonTimeNodeData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cTn"
    )


class CT_TLTimeNodeSequence(BaseOxmlElement):
    """`p:seq` element — sequence time node.

    In the canonical slide animation tree the "main sequence"
    (``p:seq nodeType="mainSeq"``) holds click-triggered entrance /
    emphasis / exit effects. This MVP surfaces only ``cTn`` plus
    the two attributes a downstream animation-sequence API will need
    first: ``concurrent`` and ``nextAc``.
    """

    _tag_seq = ("p:cTn", "p:prevCondLst", "p:nextCondLst")
    cTn: CT_TLCommonTimeNodeData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cTn"
    )
    prevCondLst = ZeroOrOne("p:prevCondLst", successors=_tag_seq[2:])
    nextCondLst = ZeroOrOne("p:nextCondLst", successors=())
    del _tag_seq


def iter_main_sequence_effects(timing):
    """Generate each effect-level ``p:par`` in the main animation sequence.

    `timing` is a ``p:timing`` element (``CT_SlideTiming`` instance) or
    ``None``. Yields every ``p:par`` whose parent chain includes a
    ``p:seq[@nodeType='mainSeq']`` and whose ``p:cTn`` child carries a
    ``@presetClass`` attribute (i.e. represents an actual entrance /
    exit / emphasis / motion-path effect rather than a grouping node).

    Returns an empty generator when `timing` is ``None`` or when the
    slide has no main animation sequence. Order matches document order
    inside the main sequence, which is the order PowerPoint uses when
    stepping through click-triggered effects.

    This helper is the backbone of ``Slide.animation_sequence`` and is
    deliberately kept free of Python-side proxy classes so downstream
    animation-authoring work can reuse it.
    """
    if timing is None:
        return
    # -- PowerPoint canonically places the main sequence at
    # --   p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq
    # -- where the inner `p:cTn` carries @nodeType="mainSeq". The
    # -- @nodeType attribute lives on `p:cTn`, NOT on `p:seq` directly
    # -- (per pml.xsd — CT_TLTimeNodeSequence has no nodeType attr).
    # -- Descending from the whole timing subtree tolerates the
    # -- occasional synthetic fixture that omits the wrapping
    # -- `p:par`/`p:cTn/childTnLst` shell.
    main_seqs = timing.xpath(".//p:seq[p:cTn/@nodeType='mainSeq']")
    if not main_seqs:
        return
    # -- Effect-level p:par nodes carry a `p:cTn` with @presetClass. The
    # -- main sequence has intermediate `p:par` wrappers (the click-group
    # -- par, the "step" par, and an inner par per effect) — every
    # -- wrapper has a p:cTn but only the innermost one carries
    # -- @presetClass. Using that attribute as the discriminator is
    # -- robust to the four-level vs five-level tree variants PowerPoint
    # -- emits for click-vs-with-previous effects. `@presetClass` is
    # -- unqualified (in the default namespace) so the local attribute
    # -- name is used directly.
    for main_seq in main_seqs:
        for par in main_seq.iter(qn("p:par")):
            cTn = par.find(qn("p:cTn"))
            if cTn is None:
                continue
            if "presetClass" in cTn.attrib:
                yield par


def first_spTgt_spid(par):
    """Return ``@spid`` (int) of first ``p:spTgt`` descendant, or ``None``.

    Walks every ``p:tgtEl/p:spTgt`` descendant of `par` and returns the
    first non-``None`` ``@spid`` value. The effect's behaviors (``p:set``,
    ``p:anim``, ``p:animEffect`` …) each carry their own ``p:tgtEl``; for
    most PowerPoint-authored effects every behavior targets the same
    shape so returning the first is correct. Used by issue #256's
    ``AnimationEffect.shape_id`` accessor.
    """
    for spTgt in par.iter(qn("p:spTgt")):
        if spTgt.spid is not None:
            return spTgt.spid
    return None


# -- CT_TLTimeCondition and CT_TLTimeConditionList are defined in
# -- ``pptx.oxml.animation`` (canonical home shipped by #102). Re-export
# -- them here so ``from pptx.oxml.timing import CT_TLTimeCondition``
# -- continues to work for callers that predate the consolidation
# -- (issue #861's tests import from this module). The #861 ``add_cond``
# -- helper was folded into ``animation.CT_TLTimeConditionList``.
# -- #256 originally defined a `CT_TLShapeTargetElement` here; that class
# -- is also now canonical in `pptx.oxml.animation` (introduced by #102)
# -- so it's re-exported below.
from pptx.oxml.animation import (  # noqa: E402,F401
    CT_TLShapeTargetElement,
    CT_TLTimeCondition,
    CT_TLTimeConditionList,
)
