.. _PpTransitionType:

``PP_TRANSITION_TYPE``
======================

Identifies a slide-transition variant.

Each member corresponds to a child element of ``p:transition`` (e.g.
``p:fade``, ``p:push``, ``p14:morph``). ``PP_TRANSITION_TYPE.NONE`` is
a return-value-only member returned when no transition-variant child is
present on the ``p:transition`` element.

Example::

    from pptx.enum.transition import PP_TRANSITION_TYPE

    slide.transition.type = PP_TRANSITION_TYPE.FADE

----

NONE
    No transition variant child element is present.

BLINDS
    Blinds transition.

CHECKER
    Checker transition.

CIRCLE
    Circle transition.

DISSOLVE
    Dissolve transition.

COMB
    Comb transition.

COVER
    Cover transition.

CUT
    Cut transition.

DIAMOND
    Diamond transition.

FADE
    Fade transition.

NEWSFLASH
    Newsflash transition.

PLUS
    Plus transition.

PULL
    Pull transition.

PUSH
    Push transition.

RANDOM
    Random transition.

RANDOM_BAR
    Random-bar transition.

SPLIT
    Split transition.

STRIPS
    Strips transition.

WEDGE
    Wedge transition.

WHEEL
    Wheel transition.

WIPE
    Wipe transition.

ZOOM
    Zoom transition.

MORPH
    Morph transition (Office 2010 extension namespace ``p14:morph``).
    When selected, PowerPoint writes a ``<p14:morph .../>`` child,
    typically wrapped in ``<mc:AlternateContent>`` so older viewers fall
    back gracefully.
