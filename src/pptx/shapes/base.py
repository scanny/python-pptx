"""Base shape-related objects such as BaseShape."""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, cast

from lxml import etree

from pptx.action import ActionSetting
from pptx.dml.effect import ShadowFormat
from pptx.oxml.ns import qn
from pptx.shared import ElementProxy
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from typing import Protocol

    from pptx.animation import AnimationEffect
    from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.oxml.shapes.shared import CT_Placeholder
    from pptx.oxml.slide import CT_Slide
    from pptx.parts.slide import BaseSlidePart
    from pptx.types import ProvidesPart
    from pptx.util import Length

    class _ShapesParent(Protocol):
        """Structural type for the shape-collection parent of a shape.

        The concrete parent class (e.g. `SlideShapes`) exposes these three members that
        `BaseShape.duplicate()` needs to append a new shape to the shape tree.
        """

        _spTree: CT_GroupShape

        @property
        def _next_shape_id(self) -> int: ...

        def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape: ...


class BaseShape(object):
    """Base class for shape objects.

    Subclasses include |Shape|, |Picture|, and |GraphicFrame|.
    """

    def __init__(self, shape_elm: ShapeElement, parent: ProvidesPart):
        super().__init__()
        self._element = shape_elm
        self._parent = parent

    def __eq__(self, other: object) -> bool:
        """|True| if this shape object proxies the same element as *other*.

        Equality for proxy objects is defined as referring to the same XML element, whether or not
        they are the same proxy object instance.
        """
        if not isinstance(other, BaseShape):
            return False
        return self._element is other._element

    def __ne__(self, other: object) -> bool:
        if not isinstance(other, BaseShape):
            return True
        return self._element is not other._element

    def bring_forward(self) -> None:
        """Move this shape one position toward the front in its sibling z-order.

        Z-order on a slide is the document order of shape children within the owning
        `p:spTree` (or `p:grpSp` when the shape belongs to a group). The first shape
        in document order is the backmost and the last is the frontmost. If the shape
        is already the frontmost of its siblings this call is a no-op.

        .. versionadded:: 2026.05.0
        """
        siblings = self._zorder_siblings
        idx = siblings.index(self._element)
        if idx == len(siblings) - 1:
            return
        parent = self._element.getparent()
        assert parent is not None
        # -- move this element to the position immediately after the next sibling --
        parent.remove(self._element)
        parent.insert(parent.index(siblings[idx + 1]) + 1, self._element)

    def bring_to_front(self) -> None:
        """Move this shape to the frontmost position in its sibling z-order.

        The shape becomes the last shape child under its owning `p:spTree` or
        `p:grpSp` (but before any trailing `p:extLst`). If the shape is already
        frontmost this call is a no-op.

        .. versionadded:: 2026.05.0
        """
        siblings = self._zorder_siblings
        if siblings[-1] is self._element:
            return
        parent = self._element.getparent()
        assert parent is not None
        parent.remove(self._element)
        # -- insert before a trailing p:extLst if present, otherwise append --
        parent.insert_element_before(self._element, "p:extLst")

    @property
    def alt_text(self) -> str:
        """Accessibility description (alt-text) for this shape.

        Read/write. Maps to the ``descr`` attribute on the shape's ``cNvPr``
        element (``p:cNvPr`` for slide-level shapes such as ``p:sp``,
        ``p:pic``, ``p:cxnSp``, ``p:graphicFrame`` and ``p:grpSp``).
        Returns the empty string when the attribute is absent (PowerPoint's
        default).

        This is the "Description" text shown by PowerPoint's Alt Text pane
        and consumed by screen readers to describe the shape's content to
        users who cannot see it. See :attr:`title` for the companion short
        title. Issue #508.

        Assigning the empty string removes the ``descr`` attribute (matching
        PowerPoint's behavior of treating the default as "no alt text").
        """
        return self._element._nvXxPr.cNvPr.descr  # pyright: ignore[reportPrivateUsage]

    @alt_text.setter
    def alt_text(self, value: str) -> None:
        self._element._nvXxPr.cNvPr.descr = value  # pyright: ignore[reportPrivateUsage]

    @property
    def animation(self) -> AnimationEffect | None:
        """|AnimationEffect| for this shape's animation, or |None| if none.

        Returns an :class:`~pptx.animation.AnimationEffect` proxy for
        the entrance / emphasis / exit effect currently bound to this
        shape by the slide's ``p:timing`` subtree. Returns ``None`` when
        the shape has no animation, when the shape is not on a slide
        (e.g. a shape on a layout or master), or when the slide's
        timing tree is present but contains no effect targeting this
        shape's id.

        Only the first matching effect is returned; the MVP
        :meth:`set_animation` replaces any prior effect so there is
        normally at most one effect per shape.

        The returned proxy is read-only in the MVP. To change an
        animation, call :meth:`set_animation` (which replaces the
        effect with a fresh one) or delete ``shape.animation`` by
        calling ``shape.set_animation(None)``.

        See :mod:`pptx.animation` for the supported preset set and the
        list of deferred (complex) effects.

        .. versionadded:: 2026.05.0
        """
        from pptx.animation import AnimationEffect, _find_effect_par_for_spid

        sld = self._owning_sld
        if sld is None:
            return None
        try:
            spid = self.shape_id
        except Exception:  # pragma: no cover - defensive
            return None
        par_elm = _find_effect_par_for_spid(sld, spid)
        if par_elm is None:
            return None
        return AnimationEffect(par_elm)

    def set_animation(
        self,
        effect_type: MSO_ANIMATION_TYPE | None,
        trigger: MSO_ANIMATION_TRIGGER | str = "onClick",
        delay: int = 0,
    ) -> AnimationEffect | None:
        """Bind an animation preset to this shape.

        Replaces any existing animation on the shape. Passing
        ``effect_type=None`` removes the animation (if any) and returns
        ``None``.

        Arguments:

        * ``effect_type`` — a member of
          :class:`~pptx.enum.animation.MSO_ANIMATION_TYPE` selecting
          one of the MVP presets (:attr:`APPEAR`, :attr:`FADE_IN`,
          :attr:`FLY_IN`, :attr:`PULSE`, :attr:`FADE_OUT`), or
          ``None`` to remove the existing animation. Presets outside
          this set raise :class:`NotImplementedError` — see
          :mod:`pptx.animation` for the deferred-items list.
        * ``trigger`` — either a member of
          :class:`~pptx.enum.animation.MSO_ANIMATION_TRIGGER` or one
          of the strings ``"onClick"`` (default) or ``"onPrev"``
          (an alias for
          :attr:`MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS` accepted for
          parity with the PowerPoint XML token). The string
          ``"afterEffect"`` is also accepted.
        * ``delay`` — non-negative int, milliseconds from the trigger
          firing to when the effect starts. Defaults to 0.

        Raises ``ValueError`` if the shape is not on a slide (it has
        no ``p:timing`` parent to write to), or if ``delay`` is
        negative.

        Returns the new :class:`~pptx.animation.AnimationEffect` proxy
        for the effect, or ``None`` when ``effect_type=None``.

        Out of scope (raises :class:`NotImplementedError`):

        * Motion-path effects.
        * Triggers other than ``onClick`` / ``onPrev`` / ``afterEffect``.
        * MORPH (transition-level, not a shape animation).

        See ``docs/dev/analysis/f8-animations-transitions.rst`` for
        the downstream items that will expand the supported set.

        .. versionadded:: 2026.05.0
        """
        from pptx.animation import (
            _find_effect_par_for_spid,
            _remove_effect_par,
            _set_animation,
        )
        from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE

        sld = self._owning_sld
        if sld is None:
            raise ValueError(
                "shape is not on a slide — animations can only be set on"
                " slide-level shapes (not on masters, layouts, or"
                " group-shape descendants)."
            )
        spid = self.shape_id

        if effect_type is None:
            existing = _find_effect_par_for_spid(sld, spid)
            if existing is not None:
                _remove_effect_par(sld, existing)
            return None

        # -- Normalize trigger arg: accept string aliases and enum members. --
        if isinstance(trigger, str):
            if trigger == "onClick" or trigger == "clickEffect":
                trigger_enum = MSO_ANIMATION_TRIGGER.ON_CLICK
            elif trigger == "onPrev" or trigger == "afterEffect":
                trigger_enum = MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS
            else:
                raise ValueError(
                    "unrecognized trigger %r — use 'onClick', 'onPrev', or a"
                    " member of MSO_ANIMATION_TRIGGER" % (trigger,)
                )
        else:
            trigger_enum = trigger

        if not isinstance(effect_type, MSO_ANIMATION_TYPE):
            raise TypeError(
                "effect_type must be an MSO_ANIMATION_TYPE member or None, got"
                " %s" % type(effect_type).__name__
            )

        return _set_animation(sld, spid, effect_type, trigger_enum, delay)

    @property
    def _owning_sld(self) -> CT_Slide | None:
        """The owning ``p:sld`` element, or ``None`` for non-slide shapes.

        Walks up the shape's ancestor chain until a ``p:sld`` is found.
        Returns ``None`` for shapes living on a slide layout or master
        (they cannot have animations) and for shapes not yet attached
        to a shape-tree.
        """
        from pptx.oxml.ns import qn

        elm = self._element
        sld_tag = qn("p:sld")
        while elm is not None:
            if elm.tag == sld_tag:
                return elm  # type: ignore[return-value]
            elm = elm.getparent()
        return None

    @lazyproperty
    def click_action(self) -> ActionSetting:
        """|ActionSetting| instance providing access to click behaviors.

        Click behaviors are hyperlink-like behaviors including jumping to a hyperlink (web page)
        or to another slide in the presentation. The click action is that defined on the overall
        shape, not a run of text within the shape. An |ActionSetting| object is always returned,
        even when no click behavior is defined on the shape.
        """
        cNvPr = self._element._nvXxPr.cNvPr  # pyright: ignore[reportPrivateUsage]
        return ActionSetting(cNvPr, self)

    def delete(self) -> None:
        """Remove this shape from the slide it appears on.

        The shape's XML element is removed from its parent `p:spTree`. Subclasses may override
        this method to perform additional cleanup, such as dropping a no-longer-referenced image
        part for a picture shape.

        Subsequent use of this shape object is undefined; most operations will raise an exception.

        .. versionadded:: 2026.05.0
        """
        self._element.getparent().remove(self._element)

    def replace_with(self, other_shape: BaseShape) -> None:
        """Swap `other_shape` into this shape's position and delete this shape.

        Copies this shape's position and size (``left``, ``top``, ``width``, ``height``) onto
        `other_shape`, moves `other_shape`'s XML element to this shape's slot in the parent
        shape-tree (preserving z-order), then deletes this shape. This provides a convenient
        single-call replacement flow such as::

            new_pic = slide.shapes.add_picture("new.png", 0, 0)
            old_pic.replace_with(new_pic)

        `other_shape` must already belong to a shape-tree (typically just created by one of the
        ``add_*`` methods) and must not be this same shape. Any of this shape's position/size
        values that are |None| are skipped rather than copied.

        After this call returns, this shape should not be used further — its XML element has
        been removed from the shape-tree (and any subclass-specific cleanup, such as dropping
        a picture's image relationship, has run).

        .. versionadded:: 2026.05.0
        """
        if other_shape is self or other_shape._element is self._element:
            raise ValueError("cannot replace a shape with itself")
        parent = self._element.getparent()
        if parent is None:
            raise ValueError("shape has no parent shape tree; cannot replace_with")
        other_parent = other_shape._element.getparent()
        if other_parent is None:
            raise ValueError(
                "other_shape has no parent shape tree; it must be added to a shape tree"
                " before it can replace another shape"
            )

        # -- copy position and size from self onto other_shape; skip when source is None --
        if self.left is not None:
            other_shape.left = self.left
        if self.top is not None:
            other_shape.top = self.top
        if self.width is not None:
            other_shape.width = self.width
        if self.height is not None:
            other_shape.height = self.height

        # -- move other_shape's element into self's slot for z-order preservation --
        target_idx = parent.index(self._element)
        other_parent.remove(other_shape._element)
        parent.insert(target_idx, other_shape._element)

        # -- dispatch to subclass delete() for part-cleanup (e.g. Picture drops image rel) --
        self.delete()

    def duplicate(self) -> BaseShape:
        """Return a new shape that is a duplicate of this shape.

        The new shape is appended to the end of the same shape tree as this shape (making it
        topmost in z-order) and is an exact copy of this shape's XML, except that it is assigned
        a new unique shape-id and a new unique name.

        Only simple shapes are supported in this implementation: auto-shapes, text-boxes, and
        connectors. Duplicating a picture, chart, table, or media shape raises
        `NotImplementedError` because those shapes own one or more package-relationships (to an
        image, embedded chart or xlsx part, etc.) that must also be copied for the duplicate to be
        valid; that work is deferred to a follow-up.

        Duplicating a group-shape is handled by :meth:`GroupShape.duplicate` (issue #1085),
        which overrides this method to clone the whole subtree via
        :class:`~pptx.opc.package.PartRelationshipCloner` and place the duplicate at the
        original's slide-relative effective rectangle (see :attr:`effective_left`).

        .. versionadded:: 2026.05.0
        """
        sp_tag = qn("p:sp")
        cxnSp_tag = qn("p:cxnSp")
        src_tag = self._element.tag

        if src_tag not in (sp_tag, cxnSp_tag):
            raise NotImplementedError(
                "shape.duplicate() is only implemented for simple shapes (auto-shape, text-box,"
                " connector); duplicating pictures, charts, tables, group-shapes, and media"
                " shapes requires copying package relationships and is not yet supported."
            )

        # -- placeholders duplicate their idx and that breaks the "one shape per idx" invariant --
        if self._element.has_ph_elm:
            raise NotImplementedError(
                "shape.duplicate() does not support placeholder shapes; placeholders are cloned"
                " from the slide layout rather than duplicated on the slide."
            )

        # -- parent is a `_BaseShapes` collection with `_spTree`, `_next_shape_id`, and
        # -- `_shape_factory`; the BaseShape type signature is deliberately narrower
        # -- (`ProvidesPart`), so narrow here via a structural cast.
        parent = cast("_ShapesParent", self._parent)
        spTree = parent._spTree  # pyright: ignore[reportPrivateUsage]

        # -- deep-copy the element so the duplicate is independent of the source --
        new_elm = copy.deepcopy(self._element)

        # -- assign a new, unique shape id and name --
        new_elm._nvXxPr.cNvPr.id = (  # pyright: ignore[reportPrivateUsage]
            parent._next_shape_id  # pyright: ignore[reportPrivateUsage]
        )
        new_elm._nvXxPr.cNvPr.name = _unique_shape_name(  # pyright: ignore[reportPrivateUsage]
            self.name, spTree
        )

        # -- append before any trailing `p:extLst` --
        spTree.insert_element_before(new_elm, "p:extLst")

        return parent._shape_factory(new_elm)  # pyright: ignore[reportPrivateUsage]

    @property
    def element(self) -> ShapeElement:
        """`lxml` element for this shape, e.g. a CT_Shape instance.

        Note that manipulating this element improperly can produce an invalid presentation file.
        Make sure you know what you're doing if you use this to change the underlying XML.
        """
        return self._element

    @property
    def flip_horizontal(self) -> bool:
        """|True| if this shape is flipped horizontally (mirrored left-to-right).

        Read/write. Maps to the ``flipH`` attribute on the shape's ``a:xfrm``
        element. Returns |False| when the shape has no ``a:xfrm`` or when
        ``flipH`` is absent (PowerPoint's default). Assigning a truthy value
        creates an ``a:xfrm`` if one is not already present.
        """
        return bool(self._element.flipH)

    @flip_horizontal.setter
    def flip_horizontal(self, value: bool) -> None:
        self._element.flipH = bool(value)

    @property
    def flip_vertical(self) -> bool:
        """|True| if this shape is flipped vertically (mirrored top-to-bottom).

        Read/write. Maps to the ``flipV`` attribute on the shape's ``a:xfrm``
        element. Returns |False| when the shape has no ``a:xfrm`` or when
        ``flipV`` is absent (PowerPoint's default). Assigning a truthy value
        creates an ``a:xfrm`` if one is not already present.
        """
        return bool(self._element.flipV)

    @flip_vertical.setter
    def flip_vertical(self, value: bool) -> None:
        self._element.flipV = bool(value)

    def flip_horizontally(self) -> None:
        """Toggle this shape's horizontal-flip state.

        Equivalent to ``shape.flip_horizontal = not shape.flip_horizontal``.
        Provided for parity with the PowerPoint UI's ``Flip Horizontal``
        command, which mirrors the shape left-to-right on each invocation.

        .. versionadded:: 2026.05.0
        """
        self._element.flipH = not bool(self._element.flipH)

    def flip_vertically(self) -> None:
        """Toggle this shape's vertical-flip state.

        Equivalent to ``shape.flip_vertical = not shape.flip_vertical``.
        Provided for parity with the PowerPoint UI's ``Flip Vertical``
        command, which mirrors the shape top-to-bottom on each invocation.

        .. versionadded:: 2026.05.0
        """
        self._element.flipV = not bool(self._element.flipV)

    @property
    def has_chart(self) -> bool:
        """|True| if this shape is a graphic frame containing a chart object.

        |False| otherwise. When |True|, the chart object can be accessed using the ``.chart``
        property.
        """
        # This implementation is unconditionally False, the True version is
        # on GraphicFrame subclass.
        return False

    @property
    def has_math_equation(self) -> bool:
        """|True| if this shape contains an Office Math (OMML) equation.

        An equation is detected by the presence of one or more `m:oMath` descendants
        anywhere within the shape's XML subtree. This covers both the `m:oMath` that
        appears directly under an `a14:m` wrapper (PowerPoint 2010+ equation shapes
        surfaced via an `mc:AlternateContent/mc:Choice` element) and the `m:oMath` that
        appears inside an `m:oMathPara` container.

        This is a read-only, MVP-scope property. Writing OMML and converting to/from
        LaTeX are explicitly deferred (see `math_equation_xml`).

        .. versionadded:: 2026.05.0
        """
        return bool(self._element.xpath(".//m:oMath"))

    @property
    def math_equation_xml(self) -> str | None:
        """Serialized `m:oMath` XML for the first equation in this shape, or |None|.

        Returns the raw OMML (Office Math Markup Language) subtree of the first
        `m:oMath` descendant of the shape's XML element, serialized as a Unicode
        string. Returns |None| when the shape contains no equation (see
        `has_math_equation`).

        The returned string is a best-effort read-only snapshot -- it includes whatever
        namespace declarations lxml chooses to emit so the fragment is well-formed on
        its own. Mutating the returned string has no effect on the presentation.

        MVP scope -- read-only access to the raw OMML is provided so callers can feed
        it into an external renderer (e.g. a MathML/LaTeX converter). The following
        are **explicitly deferred** to a follow-up feature iteration of issue #126:

        - writing / inserting equations programmatically
        - converting OMML to or from LaTeX or MathML
        - a structured `Equation` proxy object with typed accessors

        Until then, callers who need to modify an equation must edit the underlying
        XML element directly (via `shape.element`) and re-save the presentation.

        .. versionadded:: 2026.05.0
        """
        oMath_elms = self._element.xpath(".//m:oMath")
        if not oMath_elms:
            return None
        return etree.tostring(oMath_elms[0], encoding="unicode")

    @property
    def has_table(self) -> bool:
        """|True| if this shape is a graphic frame containing a table object.

        |False| otherwise. When |True|, the table object can be accessed using the ``.table``
        property.
        """
        # This implementation is unconditionally False, the True version is
        # on GraphicFrame subclass.
        return False

    @property
    def has_text_frame(self) -> bool:
        """|True| if this shape can contain text."""
        # overridden on Shape to return True. Only <p:sp> has text frame
        return False

    @property
    def height(self) -> Length:
        """Read/write. Integer distance between top and bottom extents of shape in EMUs."""
        return self._element.cy

    @height.setter
    def height(self, value: Length):
        self._element.cy = value

    @property
    def effective_height(self) -> Length | None:
        """Slide-relative height of this shape after group-transform compositing.

        Equivalent to :attr:`height` for a top-level shape, but for a shape nested
        inside one or more :class:`.GroupShape` ancestors, the height is scaled by
        the cumulative ``a:chExt``/``a:ext`` ratio of each enclosing group so it
        reflects the size the shape actually renders at on the slide.

        Returns |None| when the shape has no ``a:ext`` of its own (i.e. its raw
        :attr:`height` would also be |None|).

        Related to issue #925: ``shape.height`` on a child of a group returns the
        raw XML value which is expressed in the enclosing group's child coordinate
        system and can differ from the rendered size when the group has been
        resized in PowerPoint.

        .. versionadded:: 2026.05.0
        """
        return self._effective_geometry[3]

    @property
    def effective_left(self) -> Length | None:
        """Slide-relative left coordinate of this shape after group-transform compositing.

        Equivalent to :attr:`left` for a top-level shape. For a shape nested inside
        one or more :class:`.GroupShape` ancestors the value is recomputed by
        walking each enclosing group and mapping the shape's local coordinate
        through the group's ``a:chOff``/``a:chExt`` -> ``a:off``/``a:ext`` linear
        transform so the returned value is the slide-relative position the shape
        actually renders at.

        Returns |None| when the shape has no ``a:off`` of its own (i.e. its raw
        :attr:`left` would also be |None|).

        .. versionadded:: 2026.05.0
        """
        return self._effective_geometry[0]

    @property
    def effective_top(self) -> Length | None:
        """Slide-relative top coordinate of this shape after group-transform compositing.

        See :attr:`effective_left` for the transform details.

        .. versionadded:: 2026.05.0
        """
        return self._effective_geometry[1]

    @property
    def effective_width(self) -> Length | None:
        """Slide-relative width of this shape after group-transform compositing.

        See :attr:`effective_height` for the transform details.

        .. versionadded:: 2026.05.0
        """
        return self._effective_geometry[2]

    @property
    def is_hidden(self) -> bool:
        """|True| if this shape is marked hidden on the slide.

        Read/write. Maps to the ``hidden`` attribute on the shape's ``cNvPr``
        element (``p:cNvPr`` for slide-level shapes such as ``p:sp``,
        ``p:pic``, ``p:cxnSp``, ``p:graphicFrame`` and ``p:grpSp``).
        Returns |False| when the attribute is absent (PowerPoint's default).

        Assigning a truthy value sets ``hidden="1"`` on the ``cNvPr``
        element, causing PowerPoint to omit the shape when the slide is
        presented or printed while leaving it visible in the editing
        surface. Assigning |False| removes the attribute so the shape is
        shown normally.
        """
        return self._element._nvXxPr.cNvPr.hidden  # pyright: ignore[reportPrivateUsage]

    @is_hidden.setter
    def is_hidden(self, value: bool) -> None:
        self._element._nvXxPr.cNvPr.hidden = bool(  # pyright: ignore[reportPrivateUsage]
            value
        )

    @property
    def is_placeholder(self) -> bool:
        """True if this shape is a placeholder.

        A shape is a placeholder if it has a <p:ph> element.
        """
        return self._element.has_ph_elm

    @property
    def left(self) -> Length:
        """Integer distance of the left edge of this shape from the left edge of the slide.

        Read/write. Expressed in English Metric Units (EMU)
        """
        return self._element.x

    @left.setter
    def left(self, value: Length):
        self._element.x = value

    @property
    def name(self) -> str:
        """Name of this shape, e.g. 'Picture 7'."""
        return self._element.shape_name

    @name.setter
    def name(self, value: str):
        self._element._nvXxPr.cNvPr.name = value  # pyright: ignore[reportPrivateUsage]

    @property
    def part(self) -> BaseSlidePart:
        """The package part containing this shape.

        A |BaseSlidePart| subclass in this case. Access to a slide part should only be required if
        you are extending the behavior of |pp| API objects.
        """
        return cast("BaseSlidePart", self._parent.part)

    @property
    def placeholder_format(self) -> _PlaceholderFormat:
        """Provides access to placeholder-specific properties such as placeholder type.

        Raises |ValueError| on access if the shape is not a placeholder.
        """
        ph = self._element.ph
        if ph is None:
            raise ValueError("shape is not a placeholder")
        return _PlaceholderFormat(ph)

    def send_backward(self) -> None:
        """Move this shape one position toward the back in its sibling z-order.

        The shape moves one slot earlier in the document order of its parent
        `p:spTree` or `p:grpSp`. If the shape is already the backmost of its
        siblings this call is a no-op.

        .. versionadded:: 2026.05.0
        """
        siblings = self._zorder_siblings
        idx = siblings.index(self._element)
        if idx == 0:
            return
        parent = self._element.getparent()
        assert parent is not None
        parent.remove(self._element)
        # -- place immediately before the current previous-sibling shape --
        parent.insert(parent.index(siblings[idx - 1]), self._element)

    def send_to_back(self) -> None:
        """Move this shape to the backmost position in its sibling z-order.

        The shape becomes the first shape child under its owning `p:spTree` or
        `p:grpSp` (but after any leading non-shape elements such as `p:nvGrpSpPr`
        and `p:grpSpPr`). If the shape is already backmost this call is a no-op.

        .. versionadded:: 2026.05.0
        """
        siblings = self._zorder_siblings
        if siblings[0] is self._element:
            return
        parent = self._element.getparent()
        assert parent is not None
        # -- capture target index before removing self; siblings[0] is not self here --
        target_idx = parent.index(siblings[0])
        parent.remove(self._element)
        parent.insert(target_idx, self._element)

    @property
    def rotation(self) -> float:
        """Degrees of clockwise rotation.

        Read/write float. Negative values can be assigned to indicate counter-clockwise rotation,
        e.g. assigning -45.0 will change setting to 315.0.
        """
        return self._element.rot

    @rotation.setter
    def rotation(self, value: float):
        self._element.rot = value

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """|ShadowFormat| object providing access to shadow for this shape.

        A |ShadowFormat| object is always returned, even when no shadow is
        explicitly defined on this shape (i.e. it inherits its shadow
        behavior).
        """
        return ShadowFormat(self._element.spPr)

    @property
    def shape_id(self) -> int:
        """Read-only positive integer identifying this shape.

        The id of a shape is unique among all shapes on a slide.
        """
        return self._element.shape_id

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """A member of MSO_SHAPE_TYPE classifying this shape by type.

        Like ``MSO_SHAPE_TYPE.CHART``. Must be implemented by subclasses.
        """
        raise NotImplementedError(f"{type(self).__name__} does not implement `.shape_type`")

    @property
    def title(self) -> str:
        """Accessibility title for this shape.

        Read/write. Maps to the ``title`` attribute on the shape's ``cNvPr``
        element (``p:cNvPr`` for slide-level shapes such as ``p:sp``,
        ``p:pic``, ``p:cxnSp``, ``p:graphicFrame`` and ``p:grpSp``).
        Returns the empty string when the attribute is absent (PowerPoint's
        default).

        This is the short "Title" text shown by PowerPoint's Alt Text pane
        and read alongside :attr:`alt_text` by screen readers. Issue #508.
        Note that this is distinct from a slide's *title placeholder* and
        from the presentation-level title in the core properties.

        Assigning the empty string removes the ``title`` attribute (matching
        PowerPoint's behavior of treating the default as "no title").
        """
        return self._element._nvXxPr.cNvPr.title  # pyright: ignore[reportPrivateUsage]

    @title.setter
    def title(self, value: str) -> None:
        self._element._nvXxPr.cNvPr.title = value  # pyright: ignore[reportPrivateUsage]

    @property
    def top(self) -> Length:
        """Distance from the top edge of the slide to the top edge of this shape.

        Read/write. Expressed in English Metric Units (EMU)
        """
        return self._element.y

    @top.setter
    def top(self, value: Length):
        self._element.y = value

    @property
    def width(self) -> Length:
        """Distance between left and right extents of this shape.

        Read/write. Expressed in English Metric Units (EMU).
        """
        return self._element.cx

    @width.setter
    def width(self, value: Length):
        self._element.cx = value

    @property
    def zorder_index(self) -> int:
        """Zero-based position of this shape within its parent shape-tree in z-order.

        Index 0 is the backmost shape; the highest index is the frontmost shape. The
        value reflects the document order of shape elements under the owning
        `p:spTree` (for slide-level shapes) or `p:grpSp` (for shapes contained in a
        group). Non-shape siblings such as `p:nvGrpSpPr`, `p:grpSpPr`, and `p:extLst`
        are skipped, so the returned value matches what is observed by iterating
        ``slide.shapes`` (or ``group_shape.shapes``).

        To change a shape's z-order use :meth:`bring_to_front`,
        :meth:`send_to_back`, :meth:`bring_forward`, or :meth:`send_backward`.

        .. versionadded:: 2026.05.0
        """
        return self._zorder_siblings.index(self._element)

    @property
    def _effective_geometry(
        self,
    ) -> tuple[Length | None, Length | None, Length | None, Length | None]:
        """(left, top, width, height) tuple composited through enclosing groups.

        See :attr:`effective_left` for the transform math. Returns whichever of
        the four values are |None| on the raw element unchanged; only the values
        that are present on the shape's own ``a:xfrm`` are composited.
        """
        # -- BaseShapeElement.x/y/cx/cy are declared Length but return None when --
        # -- the element has no a:xfrm; widen to Optional[int] before compositing --
        raw_x = cast("Length | None", self._element.x)
        raw_y = cast("Length | None", self._element.y)
        raw_cx = cast("Length | None", self._element.cx)
        raw_cy = cast("Length | None", self._element.cy)
        x: int | None = None if raw_x is None else int(raw_x)
        y: int | None = None if raw_y is None else int(raw_y)
        cx: int | None = None if raw_cx is None else int(raw_cx)
        cy: int | None = None if raw_cy is None else int(raw_cy)

        # -- walk enclosing p:grpSp ancestors; each contributes its own linear --
        # -- transform on the child coord-system.                              --
        parent = self._element.getparent()
        grpSp_tag = qn("p:grpSp")
        while parent is not None and parent.tag == grpSp_tag:
            grpSp = cast("CT_GroupShape", parent)
            step = _group_xfrm_params(grpSp)
            parent = grpSp.getparent()
            if step is None:
                # -- group lacks a usable transform; treat as identity --
                continue
            off_x, off_y, ext_cx, ext_cy, chOff_x, chOff_y, chExt_cx, chExt_cy = step
            # -- guard against zero-extent group; no meaningful scale available --
            sx = (ext_cx / chExt_cx) if chExt_cx else 1.0
            sy = (ext_cy / chExt_cy) if chExt_cy else 1.0
            if x is not None:
                x = int(round(off_x + (x - chOff_x) * sx))
            if y is not None:
                y = int(round(off_y + (y - chOff_y) * sy))
            if cx is not None:
                cx = int(round(cx * sx))
            if cy is not None:
                cy = int(round(cy * sy))

        return (
            None if x is None else Emu(x),
            None if y is None else Emu(y),
            None if cx is None else Emu(cx),
            None if cy is None else Emu(cy),
        )

    @property
    def _zorder_siblings(self) -> list[ShapeElement]:
        """List of shape-element siblings (including this shape) in document order.

        Siblings are the shape children of this shape's parent (`p:spTree` or
        `p:grpSp`). Non-shape children are excluded. Used to compute z-order and
        implement the z-order mutators.
        """
        parent = self._element.getparent()
        if parent is None:
            raise ValueError("shape has no parent shape tree; z-order is undefined")
        return list(cast("CT_GroupShape", parent).iter_shape_elms())


def _group_xfrm_params(
    grpSp: CT_GroupShape,
) -> tuple[int, int, int, int, int, int, int, int] | None:
    """Return (off_x, off_y, ext_cx, ext_cy, chOff_x, chOff_y, chExt_cx, chExt_cy).

    |None| if the group shape is missing any of the transform elements needed to
    compute a mapping. Values are read via XPath so the helper does not depend on
    typed accessors that are incomplete on ``CT_Transform2D``.
    """
    xfrm = grpSp.xfrm
    if xfrm is None:
        return None
    off_x_lst = cast("list[str]", xfrm.xpath("./a:off/@x"))
    off_y_lst = cast("list[str]", xfrm.xpath("./a:off/@y"))
    ext_cx_lst = cast("list[str]", xfrm.xpath("./a:ext/@cx"))
    ext_cy_lst = cast("list[str]", xfrm.xpath("./a:ext/@cy"))
    chOff_x_lst = cast("list[str]", xfrm.xpath("./a:chOff/@x"))
    chOff_y_lst = cast("list[str]", xfrm.xpath("./a:chOff/@y"))
    chExt_cx_lst = cast("list[str]", xfrm.xpath("./a:chExt/@cx"))
    chExt_cy_lst = cast("list[str]", xfrm.xpath("./a:chExt/@cy"))
    if not (
        off_x_lst
        and off_y_lst
        and ext_cx_lst
        and ext_cy_lst
        and chOff_x_lst
        and chOff_y_lst
        and chExt_cx_lst
        and chExt_cy_lst
    ):
        return None
    return (
        int(off_x_lst[0]),
        int(off_y_lst[0]),
        int(ext_cx_lst[0]),
        int(ext_cy_lst[0]),
        int(chOff_x_lst[0]),
        int(chOff_y_lst[0]),
        int(chExt_cx_lst[0]),
        int(chExt_cy_lst[0]),
    )


def _unique_shape_name(base_name: str, spTree: ShapeElement) -> str:
    """Return a name derived from `base_name` that is unique within `spTree`.

    The existing name is returned with an incrementing integer suffix appended (e.g.
    `"Rectangle 1"` -> `"Rectangle 1 2"`) until a value not already in use by any `p:cNvPr/@name`
    under `spTree` is found.
    """
    existing = set(spTree.xpath(".//p:cNvPr/@name"))
    if base_name not in existing:
        return base_name
    n = 2
    while True:
        candidate = f"{base_name} {n}"
        if candidate not in existing:
            return candidate
        n += 1


class _PlaceholderFormat(ElementProxy):
    """Provides properties specific to placeholders, such as the placeholder type.

    Accessed via the :attr:`~.BaseShape.placeholder_format` property of a placeholder shape,
    """

    def __init__(self, element: CT_Placeholder):
        super().__init__(element)
        self._ph = element

    @property
    def element(self) -> CT_Placeholder:
        """The `p:ph` element proxied by this object."""
        return self._ph

    @property
    def idx(self) -> int:
        """Integer placeholder 'idx' attribute."""
        return self._ph.idx

    @property
    def type(self) -> PP_PLACEHOLDER:
        """Placeholder type.

        A member of the :ref:`PpPlaceholderType` enumeration, e.g. PP_PLACEHOLDER.CHART
        """
        return self._ph.type
