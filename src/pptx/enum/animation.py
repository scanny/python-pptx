"""Enumerations used for shape animations.

Issue #102 MVP. These enumerations classify the common animation effect
presets (entrance / emphasis / exit) that python-pptx supports authoring
via :meth:`pptx.shapes.base.BaseShape.set_animation`, and the per-effect
trigger kinds that are supported by the MVP.

Only the small presets set that covers the 80% use case is in scope:
``APPEAR``, ``FADE_IN``, ``FLY_IN``, ``PULSE``, ``FADE_OUT``. Complex
effects (motion paths, MORPH, color emphasis, etc.) are deferred to
later downstream items — see ``docs/dev/analysis/f8-animations-transitions.rst``.
"""

from __future__ import annotations

from pptx.enum.base import BaseXmlEnum


class MSO_ANIMATION_TYPE(BaseXmlEnum):
    """Identifies an animation-effect preset.

    Each member's ``xml_value`` is the ISO-29500 ``presetClass`` token for
    the effect (``entr`` / ``emph`` / ``exit``) — but since multiple
    members can share a preset class, the enum is primarily driven by the
    (``presetClass``, ``presetID``) pair tracked in
    :data:`pptx.animation._PRESET_MAP`. The ``from_xml``/``to_xml`` helper
    inherited from :class:`BaseXmlEnum` is therefore not used directly for
    presets; callers look up members via
    :meth:`pptx.animation.AnimationEffect._type_from_preset` and friends.

    MS API Name: *none — this enum is python-pptx-specific.*

    .. versionadded:: 2026.05.0
    """

    NONE = (0, "", "No animation effect is bound to the shape.")
    """No animation is bound to the shape. Returned from
    :attr:`pptx.animation.AnimationEffect.type` only when the proxy wraps
    an element whose preset attributes are unrecognized (i.e. an
    entrance/emphasis/exit effect outside the MVP preset set)."""

    APPEAR = (1, "entr", 'Appear entrance (`presetID=1`, `presetClass="entr"`).')
    """Appear — the shape snaps into view with no visual transition.
    Implemented as a ``p:set`` behavior writing ``style.visibility="visible"``."""

    FADE_IN = (2, "entr", 'Fade-in entrance (`presetID=10`, `presetClass="entr"`).')
    """Fade-in — the shape fades from transparent to opaque.
    Implemented as a ``p:animEffect`` with ``transition="in"`` and
    ``filter="fade"``."""

    FLY_IN = (3, "entr", 'Fly-in entrance (`presetID=2`, `presetClass="entr"`).')
    """Fly-in — the shape flies in from the bottom of the slide.
    Implemented as a ``p:animEffect`` with ``transition="in"`` and
    ``filter="slide(fromBottom)"``. Direction is fixed to ``fromBottom``
    in the MVP."""

    PULSE = (4, "emph", 'Pulse emphasis (`presetID=1`, `presetClass="emph"`).')
    """Pulse — the shape briefly changes size/opacity to draw attention.
    Implemented as a ``p:animEffect`` with ``transition="none"`` and
    ``filter="fade"`` as a minimal emphasis."""

    FADE_OUT = (5, "exit", 'Fade-out exit (`presetID=10`, `presetClass="exit"`).')
    """Fade-out — the shape fades from opaque to transparent on exit.
    Implemented as a ``p:animEffect`` with ``transition="out"`` and
    ``filter="fade"``."""


class MSO_ANIMATION_TRIGGER(BaseXmlEnum):
    """Identifies what causes an animation effect to play.

    Each member's ``xml_value`` is the corresponding ``p:cTn/@nodeType``
    value on the effect's containing ``p:par`` (``clickEffect`` /
    ``afterEffect``); the MVP supports two triggers covering roughly
    95% of real-world decks:

    * :attr:`ON_CLICK` — the effect plays when the slide is advanced with
      a mouse click (``nodeType="clickEffect"``). The default.
    * :attr:`AFTER_PREVIOUS` — the effect plays automatically after the
      previous effect completes (``nodeType="afterEffect"``). The
      ``p:cond/@delay="indefinite"`` start condition is replaced with an
      ``afterEffect`` start condition against the preceding sibling.

    The ``WITH_PREVIOUS`` / ``ON_MOUSE_OVER`` / ``ON_MOUSE_OUT`` /
    ``ON_DOUBLE_CLICK`` triggers are out of MVP scope and will be added
    by downstream issue #264.

    .. versionadded:: 2026.05.0
    """

    ON_CLICK = (0, "clickEffect", "Effect plays on mouse click.")
    """Effect plays when the user clicks to advance the slide.

    The surrounding ``p:par`` has ``nodeType="clickEffect"`` and its
    ``p:cTn/p:stCondLst/p:cond`` has ``@delay="0"`` (no delay) and no
    ``p:tn`` trigger."""

    AFTER_PREVIOUS = (1, "afterEffect", "Effect plays after previous effect.")
    """Effect plays automatically after the preceding animation in the
    main-sequence finishes.

    The surrounding ``p:par`` has ``nodeType="afterEffect"``. The
    ``p:cond/@delay`` attribute still controls the post-previous delay
    in milliseconds."""
