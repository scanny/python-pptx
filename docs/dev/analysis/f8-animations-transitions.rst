
Animations / transitions XML layer — Foundation F8 (MVP)
========================================================

Foundation item *F8* — the minimum-viable XML scaffolding for the
slide-transition and slide-timing subtrees. Without F8, neither the
``p:transition`` element nor the ``p:timing`` subtree has a real proxy
class in python-pptx; downstream animation / transition feature issues
(#102, #942, #256, #954, #1106, #400, #811, #501, #376, #861, #264,
#1004 — twelve items total) need a shared foundation to layer onto
incrementally.

This analysis documents what the MVP ships, what it intentionally
defers, and the specific extension points each downstream item should
target. The companion XSD reference is ``spec/ISO-IEC-29500-1/schemas/
xsd/pml.xsd`` (complex types ``CT_SlideTransition``,
``CT_SlideTiming``, ``CT_TLCommonTimeNodeData``, ``CT_TimeNodeList``
and their friends).


Scope summary
-------------

**In scope (MVP).**

1. Element classes for every ``p:transition`` variant child plus the
   ``p14:morph`` extension — so any slide that PowerPoint emitted with
   a transition round-trips through python-pptx without being dropped.
2. Element classes for ``p:timing`` / ``p:tnLst`` / ``p:par`` /
   ``p:seq`` / ``p:cTn`` and the `ST_TLTime` simple type — providing
   typed Python access to the common timing-tree nodes. Per-behaviour
   descriptors (``p:anim``, ``p:animClr``, ``p:animMotion``, etc.) are
   *not* authored — they round-trip as raw lxml elements.
3. A ``Slide.transition`` proxy (:class:`pptx.slide.Transition`)
   exposing four read/write properties: ``type`` (a
   :class:`PP_TRANSITION_TYPE` enum member), ``duration`` (ms), and
   the two advance-control booleans/ints ``advance_on_click`` and
   ``advance_after_time``.
4. A ``Slide.has_animations`` read-only boolean and a ``Slide.timing_xml``
   read-only string so downstream code can introspect or reach into
   the timing subtree without touching ``slide._element``.

**Out of scope (deferred — see the downstream-items matrix below).**

* Structured entrance / exit / emphasis / motion-path effect API
  (``Slide.animations.add_entrance_effect(...)``). That is downstream
  #102, #264, #861, #1106.
* Per-variant direction / orientation accessors on
  ``CT_SlideTransition`` children (e.g. ``Fade.through_black``,
  ``Cover.direction``). That is downstream #1004.
* Office 2010 MORPH transition authoring including the
  ``mc:AlternateContent`` wrapper that lets pre-2010 viewers fall back
  to a no-op. That is downstream #942 (with F3 help).
* Media-timing integration (start time for added videos, gif-frame
  animation). Downstream #376, #501, #811.
* ``mc:AlternateContent``-aware timing traversal (don't append a
  second ``p:timing`` when one is already wrapped). Downstream #954
  (with F3 help).


What the MVP ships
------------------

Foundation F8 adds the following concrete pieces; file paths are
absolute within the ``src/pptx`` tree.


1. New / enhanced oxml element classes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

``pptx.oxml.timing`` (new file)
  ``CT_SlideTiming`` — typed ``p:timing`` element with ``tnLst`` and
  ``bldLst`` descriptors (the previous implementation lived in
  ``pptx.oxml.slide`` and only exposed ``tnLst``).

  ``CT_TimeNodeList`` — typed ``p:tnLst`` / ``p:childTnLst``. Retains
  the pre-existing ``add_video()`` / ``_next_cTn_id`` helpers so
  ``SlideShapes.add_movie`` still works unchanged.

  ``CT_TLCommonTimeNodeData`` — typed ``p:cTn``. Surfaces three
  attributes (``id``, ``nodeType``, ``dur``); the remaining 21
  attributes described by pml.xsd round-trip via lxml default
  attribute handling.

  ``CT_TLTimeNodeParallel`` — typed ``p:par``.

  ``CT_TLTimeNodeSequence`` — typed ``p:seq`` with the two
  attributes downstream sequence work will need first
  (``concurrent``, ``nextAc``) — plus ``prevCondLst`` / ``nextCondLst``
  children.

  ``ST_TLTime`` — the ST_TLTime simple type (union of
  ``xsd:unsignedInt`` and the token ``"indefinite"``).

``pptx.oxml.slide`` (enhanced)
  ``CT_Slide.transition`` is now a real ``ZeroOrOne`` descriptor
  (previously only listed in ``_tag_seq``). ``CT_Slide.timing`` is
  typed as ``CT_SlideTiming | None``.

  ``CT_SlideTransition`` — typed ``p:transition``. Exposes the three
  transition attributes (``spd``, ``advClick``, ``advTm``) and the
  ``p14:dur`` extension-attribute accessor. The variant-child choice
  is handled by ``variant_tag`` / ``set_variant()`` helpers rather
  than a :class:`ZeroOrOneChoice` descriptor (see *Implementation
  notes* below).

  ``CT_TransitionVariant`` — registered for every plain-``p:``
  variant tag (``p:fade``, ``p:push``, ``p:wipe``, ``p:cover``,
  ``p:wheel``, ``p:zoom``, and the 15 others enumerated in pml.xsd).
  An empty wrapper whose job is to provide a typed class so the
  element's existence is preserved verbatim on round-trip.

  ``CT_TransitionMorph`` — registered for ``p14:morph``. Carries the
  ``@option`` attribute (``byObject`` / ``byWord`` / ``byChar``).


2. New enum module
~~~~~~~~~~~~~~~~~~

``pptx.enum.transition`` (new file)
  ``PP_TRANSITION_TYPE`` — 23 members covering the 21 ``p:`` transition
  variants plus ``MORPH`` (``p14:morph``) and a sentinel ``NONE``.
  Each member's ``xml_value`` is the local tag name.

  ``PP_TRANSITION_SPEED`` — mirrors ST_TransitionSpeed (``slow`` /
  ``med`` / ``fast``). Not wired into the MVP surface (the proxy
  reads/writes the more-precise ``p14:dur`` instead), but shipped
  now so downstream #1004 has a place to go.


3. Namespace registration
~~~~~~~~~~~~~~~~~~~~~~~~~

``pptx.oxml.ns._nsmap`` gains ``p14`` and ``p15`` entries. The
``p14`` namespace (``http://schemas.microsoft.com/office/powerpoint/
2010/main``) is required for MORPH and the ``p14:dur`` duration
attribute. ``p15`` is added at the same time for ``p15:threadedComment``
and similar 2012-era extension elements that downstream items may
touch; registering it here saves a trivial merge conflict later.


4. New Slide API
~~~~~~~~~~~~~~~~

``pptx.slide.Slide.transition`` — returns a :class:`.Transition`
proxy. Same lazy-property semantics as ``Slide.shapes`` et al.

``pptx.slide.Slide.has_animations`` — ``bool``. ``True`` iff the
slide's XML has a non-empty ``p:timing/p:tnLst`` subtree.

``pptx.slide.Slide.timing_xml`` — ``str | None``. Serialized XML of
the ``p:timing`` subtree, or ``None`` when absent. Read-only.

``pptx.slide.Transition`` — the proxy class.


5. Tests & behave coverage
~~~~~~~~~~~~~~~~~~~~~~~~~~

* ``tests/oxml/test_timing.py`` — 26 tests covering every new class
  plus the ``ST_TLTime`` simple type.
* ``tests/oxml/test_slide.py`` — adds 21 tests covering
  ``CT_Slide.transition``, ``CT_SlideTransition`` (attributes +
  variant helpers + ``p14:dur``), ``CT_TransitionVariant``, and
  ``CT_TransitionMorph``.
* ``tests/test_slide.py`` — adds 26 tests covering ``Slide.transition``,
  ``Slide.has_animations``, ``Slide.timing_xml``, and every public
  property of ``Transition``.
* ``features/sld-transition.feature`` — round-trips a fade transition
  with explicit ``duration`` / ``advance_on_click`` /
  ``advance_after_time`` settings through a save+reopen cycle.


Implementation notes & quirks
-----------------------------

**Why the transition-variant choice is not a ZeroOrOneChoice.**
The existing :class:`pptx.oxml.xmlchemy.Choice` descriptor groups
tags in a single namespace — it has no knowledge of the cross-namespace
case represented by ``p:fade`` / ``p:push`` / ... coexisting with
``p14:morph``. Rather than extend :class:`Choice` (which is used
throughout the codebase and whose semantics I don't want to perturb
in a foundation change), F8 adds ``variant_tag`` / ``set_variant()``
helpers on ``CT_SlideTransition``. They perform the same job as the
descriptor would and give downstream MORPH work (#942) a narrow
surface to extend without touching `xmlchemy.py`.

**p14:dur vs. @spd.** ISO-29500 only specifies ``@spd`` (``slow`` /
``med`` / ``fast``). Microsoft added a ``p14:dur`` attribute on
``p:transition`` in Office 2010 for millisecond-precision duration.
Real-world files use ``p14:dur`` — every PowerPoint version since
2010 emits it alongside ``@spd``. The MVP therefore reads/writes
``p14:dur`` for ``Transition.duration``; a future downstream can
add an explicit ``Transition.speed`` accessor for ``@spd`` if needed.
When ``p14:dur`` is absent, ``duration`` returns ``None`` — callers
should not assume a default because PowerPoint itself derives one
only at render time.

**lxml namespace serialisation.** When ``p14:dur`` is added to a
``p:transition`` element that was authored *without* a ``p14``
namespace declaration in scope, lxml serialises the attribute with
an auto-generated prefix (usually ``ns0:dur``). The URI is still
``http://schemas.microsoft.com/office/powerpoint/2010/main`` and the
XML is semantically identical — round-trip through ``Presentation.save``
+ ``Presentation(...)`` preserves the value (see the behave
scenario). The cosmetic difference could be eliminated by declaring
``p14`` on ``p:sld`` at package build time; that's a candidate for a
follow-up but is not part of F8's MVP.

**``CT_TimeNodeList._next_cTn_id`` empty-list case.** The pre-F8
implementation raised ``ValueError: max() arg is an empty sequence``
when called on a ``p:tnLst`` that was not attached to a ``p:sld``
parent (because the ``/p:sld//p:cTn/@id`` XPath returned nothing).
F8 returns 1 in that case — harmless for all current callers,
useful for unit tests that construct fragment elements.


Downstream-items matrix
-----------------------

The twelve items that depend on F8 (``Blocked by foundation: F8`` in
``lab/gaps-plan.md``). Each row lists the extension points F8 leaves
in place for the downstream work.

``#942`` — MORPH transition *(delivered)*
    *Wave 5.*
    Delivered on top of F8 (branch
    ``feat/issue-942-morph-transition``).
    ``Transition.morph_option`` reads/writes ``CT_TransitionMorph.option``
    (``byObject`` / ``byWord`` / ``byChar``, validated by the new
    :class:`ST_TransitionMorphOption` simple type).
    ``Slide.transition.type = PP_TRANSITION_TYPE.MORPH`` writes the
    ``p14:morph`` variant wrapped in an ``mc:AlternateContent`` with
    an ``mc:Choice Requires="p14"`` (holding the real
    ``p:transition/p14:morph``) and an ``mc:Fallback`` (holding a
    plain ``p:transition/p:fade`` that mirrors ``@spd`` / ``@advClick``
    / ``@advTm`` so older viewers apply the same advance controls).
    Switching the type away from MORPH unwraps the
    ``mc:AlternateContent`` automatically. The implementation lives on
    ``CT_Slide`` as ``transition_effective`` /
    ``get_or_add_transition_effective`` /
    ``wrap_transition_in_alt_content`` /
    ``unwrap_transition_from_alt_content`` /
    ``remove_transition_effective``; #954 (``add_movie`` vs mc-wrapped
    timing) will likely want to mirror this pattern for the
    ``p:timing`` slot.

``#1004`` — transition duration / speed
    *Wave 8. Effort XL.*
    Already largely delivered: ``Transition.duration`` reads/writes
    ``p14:dur``; the F8 MVP covers the common authoring case. This
    item's remaining scope is exposing ``Transition.speed`` as
    :class:`PP_TRANSITION_SPEED` (``slow`` / ``med`` / ``fast``) and
    surfacing per-variant flags — ``Fade.through_black``,
    ``Cover.direction``, ``Split.orientation``, etc. The typed
    subclass hierarchy can be built incrementally: add a
    ``CT_FadeTransition(CT_TransitionVariant)`` subclass with the
    extra attribute and re-register ``p:fade`` to point at it.

``#256`` — programmatic slide-timing authoring
    *Wave 4. Effort XL.* **Partially delivered (Wave 5, read-only MVP).**
    ``Slide.animation_sequence`` returns a read-only tuple of
    ``AnimationEffectView`` describing each effect in the slide's main
    sequence (``p:seq`` whose ``p:cTn/@nodeType == "mainSeq"``). Each
    ``AnimationEffectView`` exposes ``shape_id`` (``p:spTgt/@spid``),
    ``preset_class`` / ``preset_id`` / ``preset_subtype``, and
    ``delay`` (first ``p:cond/@delay``). The supporting oxml layer
    adds ``CT_TLShapeTargetElement`` (for ``p:spTgt``), the three
    preset attributes on ``CT_TLCommonTimeNodeData``, and the
    module-level helpers :func:`~pptx.oxml.timing.iter_main_sequence_effects`
    and :func:`~pptx.oxml.timing.first_spTgt_spid`.
    Authoring (add / remove / reorder / retarget effects) is still
    deferred to #102, #264, #1106 — they layer mutation methods
    atop the read surface without further oxml changes. (The class
    was originally shipped as ``AnimationEffect`` and renamed to
    ``AnimationEffectView`` to avoid collision with the authoring
    proxy :class:`pptx.animation.AnimationEffect` delivered by #102;
    ``pptx.slide.AnimationEffect`` remains a deprecated alias for one
    release.)

``#102`` — animations on shapes
    *Wave 3. Effort M.*
    Add ``Slide.shapes[i].animation`` or
    ``Slide.animations.add_entrance_effect(shape, preset)`` backed
    by new ``CT_TLAnimateEffectBehavior`` / ``CT_TLAnimateBehavior``
    subclasses. F8 leaves ``CT_TimeNodeList`` extensible via the
    ``p:tnLst`` registration — add the new behaviour-type classes
    in a new ``pptx/oxml/animation.py`` module and register each
    with ``register_element_cls``. Use ``CT_TLCommonTimeNodeData``
    for the ``p:cTn`` children.

``#1106`` — entrance/exit animation API
    *Wave 9. Effort L.*
    Near-duplicate of #102 with a broader scope (exit / emphasis /
    motion-path included). Same extension surface. F3 is also
    required to wrap legacy-incompatible XML in
    ``mc:AlternateContent``.

``#264`` — shape-animation control
    *Wave 7. Effort XL.* Read-half landed in Wave 5 as
    ``Slide.iter_shape_animations()`` — yields a ``ShapeAnimation``
    proxy per behaviour element that exposes ``shape_id``,
    ``effect_type``, ``delay_ms``, ``duration_ms``, and the underlying
    ``element``. That directly serves the original ask on the issue
    ("modify the delay of shape animations") for inspection + manual
    XML editing. Remaining Wave-7 scope: the full authoring lift —
    add a ``Sequence`` proxy on ``Slide.animations`` that composes
    ``CT_TLTimeNodeSequence`` and exposes click-vs-with-previous-vs-
    after-previous triggers, plus writable ``delay_ms`` /
    ``duration_ms`` setters on ``ShapeAnimation`` (aligning with
    #861). F8's ``prevCondLst`` / ``nextCondLst`` descriptors on
    ``CT_TLTimeNodeSequence`` are the leverage point.

``#861`` — animation delay read/write
    *Wave 5. Effort M.* Delivered (XML-layer scope).
    ``CT_TLCommonTimeNodeData.stCondLst`` / ``.endCondLst`` are now
    ``ZeroOrOne`` descriptors, and a convenience
    ``CT_TLCommonTimeNodeData.delay`` property reads/writes the
    ``@delay`` attribute of the first ``p:stCondLst/p:cond`` child
    as an ``int`` milliseconds or the sentinel string
    ``"indefinite"``. New ``CT_TLTimeCondition`` (``p:cond``) and
    ``CT_TLTimeConditionList`` (``p:stCondLst`` / ``p:endCondLst``)
    element classes are registered so the condition tree is typed.
    A user-facing ``AnimationEffectView.delay`` / ``.duration`` wrapper is
    deferred to whatever animation-proxy API lands with #102 / #1106;
    until then, delays are read/written via the ``p:cTn.delay``
    descriptor on each timing node.

``#400`` — animation control (legacy umbrella)
    *Wave 10. Effort XL.*
    Umbrella for #102 + #1106 + #264 + #861. F8 provides the
    foundation; the downstream work is aggregating the pieces into
    a coherent ``Slide.animations`` API.

``#376`` — auto-play by seconds
    *Wave 5. Effort M.*
    Already largely delivered: ``Transition.advance_after_time``
    covers the common case (set ``@advTm``). Remaining scope is
    bulk-apply (``for slide in prs.slides: slide.transition
    .advance_after_time = 3000``) and auto-play UI. No additional
    schema work needed.

``#811`` — start time for ``add_movie``
    *Wave 13. Effort M.*
    Touch ``SlideShapes.add_movie`` to accept an optional
    ``start_time`` kwarg and write it as ``p:cTn/p:stCondLst/p:cond``
    inside the video's timing node. Use F8's
    ``CT_TLCommonTimeNodeData`` + ``CT_TimeNodeList.add_video`` —
    both preserved.

``#954`` — ``add_movie`` corrupts AlternateContent-wrapped timing
    *Wave 8. Effort L.*
    Bug fix: the pre-F8 ``_add_video_timing`` helper assumes
    ``p:timing`` is a direct child of ``p:sld``. When timing is
    wrapped in ``mc:AlternateContent`` (common on newer PowerPoint
    output) the helper appends a duplicate. Requires F3 (mc
    traversal) plus the F8 ``CT_SlideTiming`` class. The fix
    migrates ``_add_video_timing`` to resolve the effective timing
    element via the mc descriptor F3 provides, then call
    ``CT_SlideTiming.get_or_add_tnLst()``.

``#501`` — animated GIFs only showing first frame
    *Wave 6. Effort S.*
    Emit the ``p:timing`` subtree's ``nvPicPr/nvPr/p:custDataLst``
    or equivalent hint that marks the picture as a cycling blip.
    Tangentially touches F8 — the media-timing node allocation
    uses ``CT_TimeNodeList._next_cTn_id`` unchanged from pre-F8.


Non-goals (explicitly rejected for MVP)
---------------------------------------

* **``mc:AlternateContent`` traversal.** That is Foundation F3's
  job. Two downstream items (#942 MORPH and #954 add_movie bug) need
  both F3 and F8 — they will land after both foundations are in
  place.
* **Typed per-variant subclasses.** ``CT_TransitionVariant`` is a
  single class registered for every ``p:`` variant. The five
  directional sub-types (``CT_SideDirectionTransition``,
  ``CT_CornerDirectionTransition``, ``CT_EightDirectionTransition``,
  ``CT_OrientationTransition``, ``CT_InOutTransition``) and the
  three "toggle" sub-types (``CT_OptionalBlackTransition``,
  ``CT_SplitTransition``, ``CT_WheelTransition``) from pml.xsd are
  *not* represented — they round-trip via lxml's default attribute
  preservation. Downstream #1004 will subclass ``CT_TransitionVariant``
  as needed.
* **Build-list (``p:bldLst``) authoring.** ``CT_SlideTiming.bldLst``
  is a ``ZeroOrOne`` descriptor (so the element round-trips) but no
  typed class is registered for ``p:bldLst`` or its children
  (``p:bldP``, ``p:bldDgm``, ``p:bldOleChart``, ``p:bldGraphic``).
  Downstream paragraph-build authoring work can add them.
* **Animation behaviour types.** ``CT_TLAnimateBehavior``,
  ``CT_TLAnimateColorBehavior``, ``CT_TLAnimateEffectBehavior``,
  ``CT_TLAnimateMotionBehavior``, ``CT_TLAnimateRotationBehavior``,
  ``CT_TLAnimateScaleBehavior``, ``CT_TLCommandBehavior``,
  ``CT_TLSetBehavior``, ``CT_TLMediaNodeAudio`` — none of these have
  typed classes. They round-trip as raw lxml elements. Downstream
  items #102, #264, #1106 will add each one as it's needed.


Cross-references
----------------

* Issue list: the 12 downstream items referenced above are listed
  in ``lab/gaps-plan.md`` by the strings ``Blocked by foundation:
  F8`` and ``Blocked by foundation: F3,F8``.
* pml.xsd: ``spec/ISO-IEC-29500-1/schemas/xsd/pml.xsd``, specifically
  the complex types ``CT_SlideTransition`` (lines 83–114),
  ``CT_SlideTiming`` (lines 704–710), ``CT_TimeNodeList`` (lines
  229–245), and ``CT_TLCommonTimeNodeData`` (lines 297–329).
* p14 / MORPH: lives outside ISO-29500 in Microsoft's
  [MS-PPTX-ECMA] extension. The ``p14`` namespace URI is
  ``http://schemas.microsoft.com/office/powerpoint/2010/main``.
* `HISTORY.rst` entry: "- Foundation: animations/transitions XML
  layer (F8 — MVP)".
