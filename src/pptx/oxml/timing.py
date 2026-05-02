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
"""

from __future__ import annotations

from pptx.oxml.simpletypes import BaseSimpleType, XsdInt, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
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
        """Next available unique ID (int) for a new `p:cTn` element."""
        cTn_id_strs = self.xpath("/p:sld/p:timing//p:cTn/@id")
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

    Downstream extension points (see the F8 analysis doc):

    * #861: surface ``endCondLst`` and value-list semantics for
      delay read/write on more than one condition.
    * #954: wrap in ``mc:AlternateContent`` for PowerPoint 2010+
      extensibility.
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
    childTnLst = ZeroOrOne("p:childTnLst", successors=_tag_seq[5:])
    del _tag_seq

    id: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdUnsignedInt
    )
    nodeType: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "nodeType", XsdString
    )
    dur = OptionalAttribute("dur", ST_TLTime)
    presetID = OptionalAttribute("presetID", XsdInt)
    presetClass = OptionalAttribute("presetClass", XsdString)
    presetSubtype = OptionalAttribute("presetSubtype", XsdInt)


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
