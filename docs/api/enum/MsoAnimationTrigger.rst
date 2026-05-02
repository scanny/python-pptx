.. _MsoAnimationTrigger:

``MSO_ANIMATION_TRIGGER``
=========================

Identifies what causes an animation effect to play.

The MVP supports two triggers that cover the vast majority of real-world
decks; additional trigger kinds (``WITH_PREVIOUS``, ``ON_MOUSE_OVER``,
etc.) are deferred to a downstream issue.

Example::

    from pptx.enum.animation import MSO_ANIMATION_TRIGGER

    shape.animation.trigger = MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS

----

ON_CLICK
    Effect plays when the user clicks to advance the slide. This is the
    default trigger.

AFTER_PREVIOUS
    Effect plays automatically after the preceding animation in the
    main-sequence finishes.
