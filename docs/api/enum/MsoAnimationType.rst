.. _MsoAnimationType:

``MSO_ANIMATION_TYPE``
======================

Identifies an animation-effect preset bound to a shape.

The MVP preset set covers the most common entrance, emphasis, and exit
effects. The ``NONE`` member is returned only when the animation proxy
wraps an element whose preset attributes are unrecognized.

Example::

    from pptx.enum.animation import MSO_ANIMATION_TYPE

    shape.animation.type = MSO_ANIMATION_TYPE.FADE_IN

----

NONE
    No animation effect is bound to the shape.

APPEAR
    Appear entrance — the shape snaps into view with no visual transition.

FADE_IN
    Fade-in entrance — the shape fades from transparent to opaque.

FLY_IN
    Fly-in entrance — the shape flies in from the bottom of the slide.

PULSE
    Pulse emphasis — the shape briefly changes size/opacity to draw attention.

FADE_OUT
    Fade-out exit — the shape fades from opaque to transparent on exit.
