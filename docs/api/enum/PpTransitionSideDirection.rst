.. _PpTransitionSideDirection:

``PP_TRANSITION_SIDE_DIRECTION``
================================

Left / up / right / down selector for the ``@dir`` attribute on the
``p:wipe`` and ``p:push`` transition variants — the direction that
content slides *from*. The schema default is ``LEFT``.

Example::

    from pptx.enum.transition import PP_TRANSITION_SIDE_DIRECTION

    slide.transition.direction = PP_TRANSITION_SIDE_DIRECTION.RIGHT

----

LEFT
    Left — content enters from / moves toward the left.

UP
    Up — content enters from / moves toward the top.

RIGHT
    Right — content enters from / moves toward the right.

DOWN
    Down — content enters from / moves toward the bottom.
