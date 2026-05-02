.. _PpTransitionSpeed:

``PP_TRANSITION_SPEED``
=======================

Slow / medium / fast selector for the ``p:transition/@spd`` attribute.

PowerPoint renders these as approximate durations in the UI, but the XML
carries only the token value. Per-transition fine-grained duration uses
the ``p14:dur`` attribute (in milliseconds) in the 2010 extension
namespace, accessed via :attr:`~pptx.slide.Transition.duration`.

Example::

    from pptx.enum.transition import PP_TRANSITION_SPEED

    slide.transition.speed = PP_TRANSITION_SPEED.MEDIUM

----

SLOW
    Slow — approximately 1.6 seconds.

MEDIUM
    Medium — approximately 1.0 seconds.

FAST
    Fast — approximately 0.5 seconds (default).
