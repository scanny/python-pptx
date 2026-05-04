.. _animation_api:

Animation
=========

.. module:: pptx.animation

Shape-level entrance / exit / emphasis animations are surfaced through
two complementary proxies. The **authoring** proxy
:class:`pptx.animation.AnimationEffect` is returned from
:meth:`pptx.shapes.base.BaseShape.animation` and describes the single
effect (if any) currently bound to that shape. The **introspection**
proxy :class:`pptx.slide.AnimationEffectView` is returned as an element
of :attr:`pptx.slide.Slide.animation_sequence` and describes one step
in the slide's main animation sequence.

Typical read / write usage::

    from pptx import Presentation
    from pptx.enum.animation import MSO_ANIMATION_TYPE, MSO_ANIMATION_TRIGGER

    prs = Presentation("deck.pptx")
    slide = prs.slides[0]
    shape = slide.shapes[0]

    # -- add / replace an entrance animation on a shape --
    shape.set_animation(
        MSO_ANIMATION_TYPE.FADE_IN,
        trigger=MSO_ANIMATION_TRIGGER.ON_CLICK,
        delay=0,
    )

    # -- inspect the effect now bound to `shape` --
    effect = shape.animation
    assert effect.type is MSO_ANIMATION_TYPE.FADE_IN

    # -- walk the whole slide sequence for reporting --
    for step in slide.animation_sequence:
        print(step.shape_id, step.preset_class, step.preset_id, step.delay)

    prs.save("deck.pptx")


|AnimationEffect| objects
-------------------------

The |AnimationEffect| authoring proxy wraps the ``p:par`` XML element
that represents a single "click-effect" / "after-effect" step in the
slide's ``p:timing`` tree. It is returned from
:meth:`pptx.shapes.base.BaseShape.animation` and is not intended to be
constructed directly; effects are authored by
:meth:`pptx.shapes.base.BaseShape.set_animation`.

.. autoclass:: pptx.animation.AnimationEffect()
   :members:
   :member-order: bysource
   :undoc-members:


|AnimationEffectView| objects
-----------------------------

|AnimationEffectView| is a read-only view of a single entrance / exit /
emphasis effect in the slide's main animation sequence. Instances are
obtained from :attr:`pptx.slide.Slide.animation_sequence` and expose
``shape_id``, ``preset_class``, ``preset_id``, ``preset_subtype``, and
``delay`` for regression testing and timing introspection. The class
was renamed from ``AnimationEffect`` to avoid a name collision with the
authoring proxy above; the original name remains available as a
deprecated alias for one release.

.. autoclass:: pptx.slide.AnimationEffectView()
   :members:
   :member-order: bysource
   :undoc-members:
