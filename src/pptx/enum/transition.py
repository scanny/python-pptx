"""Enumerations used for slide transitions and animations.

These enumerations are part of the Foundation F8 MVP. They identify each
`p:transition` variant child element (e.g. ``p:fade``, ``p:push``,
``p14:morph``) with a stable integer token so higher-level code can select a
transition kind without string-typing the XML tag.

Downstream issues that build on this enum set (non-exhaustive):

* `#942 <https://github.com/scanny/python-pptx/issues/942>`_ —
  ``Slide.transition.type = PP_TRANSITION_TYPE.MORPH``.
* `#1004 <https://github.com/scanny/python-pptx/issues/1004>`_ —
  advance/duration read/write (this MVP ships the attribute plumbing).
* `#256 <https://github.com/scanny/python-pptx/issues/256>`_ —
  timing-programmatic support (transitions are one half of it).

See ``docs/dev/analysis/f8-animations-transitions.rst`` for the full map.
"""

from __future__ import annotations

from pptx.enum.base import BaseXmlEnum


class PP_TRANSITION_TYPE(BaseXmlEnum):
    """Identifies a slide-transition variant.

    Each member's ``xml_value`` is the local name of the child element of
    ``p:transition`` that selects this transition kind — for example
    ``"fade"`` corresponds to ``<p:fade/>`` and ``"morph"`` to
    ``<p14:morph .../>`` (the only member that lives outside the ``p:``
    namespace). ``PP_TRANSITION_TYPE.NONE`` is the "return-value only"
    member returned when no transition-variant child is present on a
    ``p:transition`` element.

    MS API Name: *none — this enum is python-pptx-specific.*

    .. versionadded:: 2026.05.0
    """

    NONE = (0, "", "No transition variant child element is present.")
    """No transition-variant child (the `p:transition` element is either
    absent or carries only `spd` / `advClick` / `advTm` attributes)."""

    BLINDS = (1, "blinds", "Blinds transition (`p:blinds`).")
    """Blinds transition."""

    CHECKER = (2, "checker", "Checker transition (`p:checker`).")
    """Checker transition."""

    CIRCLE = (3, "circle", "Circle transition (`p:circle`).")
    """Circle transition."""

    DISSOLVE = (4, "dissolve", "Dissolve transition (`p:dissolve`).")
    """Dissolve transition."""

    COMB = (5, "comb", "Comb transition (`p:comb`).")
    """Comb transition."""

    COVER = (6, "cover", "Cover transition (`p:cover`).")
    """Cover transition."""

    CUT = (7, "cut", "Cut transition (`p:cut`).")
    """Cut transition."""

    DIAMOND = (8, "diamond", "Diamond transition (`p:diamond`).")
    """Diamond transition."""

    FADE = (9, "fade", "Fade transition (`p:fade`).")
    """Fade transition."""

    NEWSFLASH = (10, "newsflash", "Newsflash transition (`p:newsflash`).")
    """Newsflash transition."""

    PLUS = (11, "plus", "Plus transition (`p:plus`).")
    """Plus transition."""

    PULL = (12, "pull", "Pull transition (`p:pull`).")
    """Pull transition."""

    PUSH = (13, "push", "Push transition (`p:push`).")
    """Push transition."""

    RANDOM = (14, "random", "Random transition (`p:random`).")
    """Random transition."""

    RANDOM_BAR = (15, "randomBar", "Random-bar transition (`p:randomBar`).")
    """Random-bar transition."""

    SPLIT = (16, "split", "Split transition (`p:split`).")
    """Split transition."""

    STRIPS = (17, "strips", "Strips transition (`p:strips`).")
    """Strips transition."""

    WEDGE = (18, "wedge", "Wedge transition (`p:wedge`).")
    """Wedge transition."""

    WHEEL = (19, "wheel", "Wheel transition (`p:wheel`).")
    """Wheel transition."""

    WIPE = (20, "wipe", "Wipe transition (`p:wipe`).")
    """Wipe transition."""

    ZOOM = (21, "zoom", "Zoom transition (`p:zoom`).")
    """Zoom transition."""

    MORPH = (
        22,
        "morph",
        "Morph transition (`p14:morph`, Office 2010 extension namespace).",
    )
    """Morph transition.

    Unlike the other members this one lives in the ``p14:`` extension
    namespace. When selected, PowerPoint writes a ``<p14:morph .../>``
    child instead of a ``p:`` child, typically wrapped in
    ``<mc:AlternateContent>`` so viewers that do not understand the
    2010 extension fall back gracefully. See the F8 analysis doc for
    the AlternateContent wrapping plan (downstream #942)."""


class PP_TRANSITION_SPEED(BaseXmlEnum):
    """Slow / medium / fast selector for the ``p:transition/@spd`` attribute.

    This reproduces ST_TransitionSpeed from ``pml.xsd``. PowerPoint
    renders these as approximate durations in the UI (1.6s / 1.0s / 0.5s
    per Microsoft's published defaults), but the XML carries only the
    token value. Per-transition fine-grained duration uses the
    ``p14:dur`` attribute (in milliseconds) in the 2010 extension
    namespace and is what :attr:`.Transition.duration` writes / reads in
    the MVP.

    .. versionadded:: 2026.05.0
    """

    SLOW = (1, "slow", "Slow — approximately 1.6 seconds.")
    """Slow — approximately 1.6 seconds."""

    MEDIUM = (2, "med", "Medium — approximately 1.0 seconds.")
    """Medium — approximately 1.0 seconds."""

    FAST = (3, "fast", "Fast — approximately 0.5 seconds (default).")
    """Fast — approximately 0.5 seconds (default)."""


class PP_TRANSITION_SIDE_DIRECTION(BaseXmlEnum):
    """Left / up / right / down selector for transition-variant ``@dir``.

    This reproduces ``ST_TransitionSideDirectionType`` from ``pml.xsd``.
    It is the value-space for the ``@dir`` attribute on the
    ``p:wipe`` and ``p:push`` transition variants (i.e. the direction
    the content slides *from*). The schema default is ``"l"`` (left).

    .. versionadded:: 2026.05.0
    """

    LEFT = (1, "l", "Left — content enters from / moves toward the left.")
    """Left."""

    UP = (2, "u", "Up — content enters from / moves toward the top.")
    """Up."""

    RIGHT = (3, "r", "Right — content enters from / moves toward the right.")
    """Right."""

    DOWN = (4, "d", "Down — content enters from / moves toward the bottom.")
    """Down."""
