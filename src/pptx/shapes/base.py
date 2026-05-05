"""Base shape-related objects such as BaseShape."""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, Iterator, NamedTuple, cast

from lxml import etree

from pptx.action import ActionSetting
from pptx.dml.effect import ShadowFormat
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, nsuri, qn
from pptx.shared import ElementProxy
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from typing import Protocol

    from pptx.animation import AnimationEffect
    from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.opc.package import Part
    from pptx.oxml.dml.shape_style import CT_ShapeStyle
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.oxml.shapes.shared import CT_NonVisualDrawingProps, CT_Placeholder
    from pptx.oxml.slide import CT_Slide
    from pptx.parts.slide import BaseSlidePart
    from pptx.text.text import TextFrame, TextFrameRect
    from pptx.types import ProvidesPart
    from pptx.util import Length

    class _ShapeWithTextFrame(Protocol):
        """Structural type for a shape that exposes a `text_frame`.

        Used by :attr:`BaseShape.text_frame_rect` to narrow from the base class
        to a subclass that defines ``text_frame`` (e.g. :class:`Shape`).
        """

        @property
        def text_frame(self) -> TextFrame: ...

    class _ShapesParent(Protocol):
        """Structural type for the shape-collection parent of a shape.

        The concrete parent class (e.g. `SlideShapes`) exposes these members that
        `BaseShape.duplicate()` and `BaseShape.clone_onto()` need to append a new
        shape to the shape tree and resolve its owning part.
        """

        _spTree: CT_GroupShape

        @property
        def _next_shape_id(self) -> int: ...

        @property
        def part(self) -> BaseSlidePart: ...

        def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape: ...


class ThemeStyleRefs(NamedTuple):
    """Four-tuple of shape-style references into the theme's format-scheme.

    Returned by :attr:`BaseShape.theme_style_refs` and accepted by its setter.
    The three integer fields index into the slide master theme's
    ``a:fmtScheme`` — ``lnStyleLst``, ``fillStyleLst``, and ``effectStyleLst``
    respectively — and ``font_ref`` is one of ``"major"``, ``"minor"``, or
    ``"none"`` (per ``ST_FontCollectionIndex``) keying into ``a:fontScheme``.

    Index ``0`` means "no style from the matrix" (PowerPoint's UI equivalent
    is "No Fill" / "No Outline"); indices ``1..N`` select the Nth style in
    the theme's corresponding list.

    .. versionadded:: 2026.05.0
    """

    line_ref: int
    fill_ref: int
    effect_ref: int
    font_ref: str


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

    @property
    def custom_props(self) -> _CustomPropsDict:
        """Dict-like mapping of application-defined string metadata on this shape.

        Provides a persisted, round-trip-safe place for callers to attach
        arbitrary string-to-string metadata to an individual shape — analogous
        to custom document properties but scoped to a single shape instead of
        the whole presentation.

        The returned object is a live proxy over the shape's XML; mutations
        are written back to the shape element immediately and survive
        :meth:`Presentation.save` / reload. Keys preserve insertion order, both
        keys and values must be |str|, and the mapping supports the usual
        ``mapping[key]`` / ``key in mapping``, ``del mapping[key]``,
        :meth:`~._CustomPropsDict.get`, :meth:`~._CustomPropsDict.clear`, and
        iteration idioms::

            shape.custom_props["department"] = "marketing"
            shape.custom_props["owner"] = "alice"
            assert shape.custom_props.get("department") == "marketing"
            assert list(shape.custom_props) == ["department", "owner"]
            del shape.custom_props["owner"]
            shape.custom_props.clear()

        The values are stored under the shape's ``p:cNvPr/a:extLst`` using a
        single ``a:ext`` element whose ``uri`` attribute is
        ``{urn:loadfix-pptx:custom-props:v1}``. PowerPoint preserves unknown
        ``a:ext`` entries on save, so the custom properties round-trip
        through PowerPoint in addition to round-tripping through python-pptx.

        .. versionadded:: 2026.05.0
        """
        cNvPr = cast(
            "CT_NonVisualDrawingProps",
            self._element._nvXxPr.cNvPr,  # pyright: ignore[reportPrivateUsage]
        )
        return _CustomPropsDict(cNvPr)

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

    @lazyproperty
    def hover_action(self) -> ActionSetting:
        """|ActionSetting| instance providing access to mouse-hover behaviors.

        Parallel to :attr:`click_action`, but backed by the shape's
        ``a:hlinkMouseOver`` (``a:hlinkHover``) element rather than
        ``a:hlinkClick``. A hover action fires when the slideshow viewer's
        mouse pointer passes over the shape, without a click being required.
        An |ActionSetting| object is always returned, even when no hover
        behavior is defined on the shape.
        """
        cNvPr = self._element._nvXxPr.cNvPr  # pyright: ignore[reportPrivateUsage]
        return ActionSetting(cNvPr, self, hover=True)

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

    def clone_onto(
        self,
        shape_tree: _ShapesParent,
        left: Length | None = None,
        top: Length | None = None,
    ) -> BaseShape:
        """Deep-copy this shape into `shape_tree`, return the new proxy.

        The source shape's full XML subtree (``<p:sp>``, ``<p:pic>``,
        ``<p:graphicFrame>``, ``<p:cxnSp>``, or ``<p:grpSp>``) is deep-copied
        and appended to the end of `shape_tree` (making the clone the topmost
        shape in z-order on that tree). Every external relationship the
        subtree carries — image blobs on ``a:blip``, chart / chartex parts on
        ``c:chart`` or ``cx:chart``, OLE payloads and their icon images,
        SmartArt's four ``dgm:relIds`` rels, ``am3d:model3D`` 3D-model media,
        hyperlink externals, and any other ``r:id`` / ``r:embed`` / ``r:link``
        reference — is re-materialised on `shape_tree`'s owning package (or
        reused when source and target already share a package) and each
        matching attribute on the clone is rewritten to the new rId so the
        result is self-contained and fully valid.

        The ``cNvPr/@id`` of the cloned top-level shape (and every descendant
        ``p:cNvPr`` inside a cloned group) is reassigned a fresh, unique id on
        the target tree so the clone never collides with an existing shape id.
        The clone's top-level name is made unique against the target tree's
        existing ``p:cNvPr/@name`` values, derived from the source name.

        Optional ``left`` / ``top`` override the clone's slide-relative
        position. ``None`` leaves the corresponding coordinate unchanged from
        the source. For a ``p:grpSp`` the override applies to the group's
        outer ``a:xfrm/a:off`` (the child coordinate system is preserved); for
        every other shape kind it applies to the shape's own
        ``p:spPr/a:xfrm/a:off`` (``p:xfrm/a:off`` for a ``p:graphicFrame``).

        `shape_tree` may belong to this same presentation (copying a shape
        from one slide to another — or cloning within the same slide) or to a
        different presentation entirely (cross-presentation copy). In the
        cross-package case every referenced part is materialised in the
        target package; no cross-package references are retained.

        Placeholders raise :class:`NotImplementedError` because a placeholder
        duplicates its ``idx`` which breaks the "one shape per idx" slide
        invariant; clone the underlying content onto a fresh non-placeholder
        shape instead.

        .. versionadded:: 2026.05.0
        """
        from pptx.opc.package import PartRelationshipCloner
        from pptx.oxml.ns import qn as _qn

        # -- Placeholders would duplicate their `idx` on the target slide,  --
        # -- which breaks the "one placeholder per idx" invariant and can   --
        # -- confuse PowerPoint's layout-inheritance machinery. Placeholder --
        # -- content belongs on the slide via the layout, not on the shape  --
        # -- tree directly.                                                 --
        if self._element.has_ph_elm:
            raise NotImplementedError(
                "BaseShape.clone_onto() does not support placeholder shapes; placeholders are"
                " cloned from the slide layout rather than from an existing placeholder."
            )

        src_part = self.part
        tgt_part = shape_tree.part
        spTree = shape_tree._spTree  # pyright: ignore[reportPrivateUsage]

        # -- 1. Deep-copy the XML subtree and re-materialise every rId-bearing  --
        # --    relationship (images, chart parts, OLE parts, etc.) against the --
        # --    target part. `PartRelationshipCloner` walks every `r:id`,       --
        # --    `r:embed`, `r:link` attribute and rewrites them on the clone.   --
        new_elm = cast(
            "ShapeElement",
            PartRelationshipCloner.clone(src_part, tgt_part, self._element),
        )

        # -- 2. Handle SmartArt's `dgm:relIds` separately: its rIds live on     --
        # --    non-standard attribute names (`r:dm` / `r:lo` / `r:qs` / `r:cs`)--
        # --    that `PartRelationshipCloner` does not recognise. Walk each one --
        # --    and rewrite it to a new rId on the target part.                 --
        _clone_dgm_relIds(new_elm, src_part, tgt_part)

        # -- 3. Reassign every `cNvPr/@id` in the cloned subtree. For a simple  --
        # --    shape this is just the top-level `p:cNvPr`; for a group shape   --
        # --    it is every descendant `p:cNvPr` so nested children don't       --
        # --    collide either.                                                 --
        # -- 4. Append to the target `p:spTree` *before* bumping ids so the     --
        # --    `_next_shape_id` readout (which checks `spTree.max_shape_id`)   --
        # --    observes each just-assigned id and yields the next free value.  --
        spTree.insert_element_before(new_elm, "p:extLst")
        for cNvPr in new_elm.xpath(".//p:cNvPr"):
            cNvPr.id = shape_tree._next_shape_id  # pyright: ignore[reportPrivateUsage]

        # -- 5. Give the clone a unique name relative to the target tree. The  --
        # --    source's name is used as the basename; descendant names inside --
        # --    a cloned group are left unchanged (matching `GroupShape.       --
        # --    duplicate` behaviour).                                         --
        new_elm._nvXxPr.cNvPr.name = _unique_shape_name(  # pyright: ignore[reportPrivateUsage]
            self.name, spTree
        )

        # -- 6. Apply optional position override. For a `p:grpSp` this writes  --
        # --    the outer `a:xfrm/a:off` (child `a:chOff`/`a:chExt` preserved);--
        # --    every other shape kind writes its own `a:off` via the standard --
        # --    `.x`/`.y` descriptors.                                         --
        if left is not None or top is not None:
            if new_elm.tag == _qn("p:grpSp"):
                xfrm = cast("CT_GroupShape", new_elm).get_or_add_xfrm()
                if left is not None:
                    xfrm.x = left
                if top is not None:
                    xfrm.y = top
            else:
                if left is not None:
                    new_elm.x = left
                if top is not None:
                    new_elm.y = top

        return shape_tree._shape_factory(new_elm)  # pyright: ignore[reportPrivateUsage]

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
    def text_frame_rect(self) -> TextFrameRect:
        """Slide-relative rectangle PowerPoint allocates for rendering this shape's text.

        Returns a :class:`~pptx.text.text.TextFrameRect` — a 4-tuple
        ``(left, top, width, height)`` of |Length| values in EMU — derived
        from the shape's position and size minus the text frame's four
        insets (``margin_left``, ``margin_top``, ``margin_right``,
        ``margin_bottom``).

        This is the rectangle text will actually render within; it is
        smaller than the shape's bounding box by the sum of the opposing
        insets on each axis. Useful for picking a font size that will not
        overflow the shape (issue #663).

        Raises |ValueError| when the shape does not have a text frame
        (e.g. a connector, or a picture shape with no text body). Check
        :attr:`has_text_frame` first when the caller is unsure.

        Notes:

        * The returned width or height may be zero or negative when the
          shape is smaller than the sum of its insets on that axis; no
          clamping is performed.
        * Text auto-fit (``auto_size`` / ``normAutofit``) and text rotation
          (``TextFrame.rotation``) are not accounted for — the rectangle
          is the authored "inset box", not the rendered glyph extent.
        * For a shape nested inside a group, :attr:`left` / :attr:`top` /
          :attr:`width` / :attr:`height` reflect the raw (pre-composite)
          coordinates; if you need the slide-relative rendered rectangle
          use :attr:`effective_left` / :attr:`effective_top` /
          :attr:`effective_width` / :attr:`effective_height` to compose
          the transform yourself.

        .. versionadded:: 2026.05.0
        """
        from pptx.text.text import TextFrameRect as _TextFrameRect

        if not self.has_text_frame:
            raise ValueError("shape has no text frame")

        # -- mypy/pyright narrowing: `.text_frame` is only defined on
        # -- subclasses that override `has_text_frame` to True; cast here.
        text_frame = cast("_ShapeWithTextFrame", self).text_frame
        margin_l = text_frame.margin_left
        margin_t = text_frame.margin_top
        margin_r = text_frame.margin_right
        margin_b = text_frame.margin_bottom
        return _TextFrameRect(
            left=Emu(int(self.left) + int(margin_l)),
            top=Emu(int(self.top) + int(margin_t)),
            width=Emu(int(self.width) - int(margin_l) - int(margin_r)),
            height=Emu(int(self.height) - int(margin_t) - int(margin_b)),
        )

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
        self._element._nvXxPr.cNvPr.hidden = bool(value)  # pyright: ignore[reportPrivateUsage]

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
    def theme_style_refs(self) -> ThemeStyleRefs | None:
        """Four-tuple of theme-style references on this shape, or |None|.

        Maps to the shape's ``p:style`` child (``a:CT_ShapeStyle``). Returns
        a :class:`ThemeStyleRefs` named-tuple with the ``idx`` of each of the
        four ``*Ref`` children (``a:lnRef``, ``a:fillRef``, ``a:effectRef``,
        ``a:fontRef``) when the shape carries a ``p:style``, or |None| when
        it does not (e.g. placeholders, graphic-frame wrappers, and group
        shapes do not have a ``p:style``).

        The first three fields are zero-based theme-matrix indices; the
        fourth (``font_ref``) is the ``ST_FontCollectionIndex`` token —
        ``"major"``, ``"minor"``, or ``"none"``.

        Assigning a :class:`ThemeStyleRefs` (or any 4-tuple in the same
        order) writes a fresh ``p:style`` subtree — each ``*Ref`` gets its
        ``idx`` set and an ``<a:schemeClr val="accent1"/>`` color-choice
        child (matching what PowerPoint writes when the user applies a
        preset from the Shape Styles gallery). Assigning |None| removes the
        ``p:style`` subtree entirely.

        Setting a ``p:style`` is only supported on simple shapes
        (``p:sp``, ``p:cxnSp``, ``p:pic``). Assigning on a shape whose
        underlying element is a graphic-frame or group-shape raises
        :class:`ValueError`.

        Sample read::

            refs = shape.theme_style_refs
            if refs is not None:
                print(refs.line_ref, refs.fill_ref, refs.effect_ref, refs.font_ref)

        Sample write::

            shape.theme_style_refs = ThemeStyleRefs(1, 2, 0, "minor")

        See issue #447.

        .. versionadded:: 2026.05.0
        """
        style = self._element.find(qn("p:style"))
        if style is None:
            return None
        style = cast("CT_ShapeStyle", style)
        return ThemeStyleRefs(
            line_ref=style.lnRef.idx,
            fill_ref=style.fillRef.idx,
            effect_ref=style.effectRef.idx,
            font_ref=style.fontRef.idx,
        )

    @theme_style_refs.setter
    def theme_style_refs(self, value: ThemeStyleRefs | tuple[int, int, int, str] | None) -> None:
        successors_by_tag = {
            qn("p:sp"): ("p:txBody", "p:extLst"),
            qn("p:cxnSp"): ("p:extLst",),
            qn("p:pic"): ("p:extLst",),
        }
        elm = self._element
        if elm.tag not in successors_by_tag:
            raise ValueError(
                "shape does not support a p:style (theme_style_refs is only available on"
                " auto-shapes, text-boxes, connectors, and pictures)"
            )
        # -- always drop any existing p:style first, whether clearing or rewriting --
        existing = elm.find(qn("p:style"))
        if existing is not None:
            elm.remove(existing)
        if value is None:
            return

        line_ref, fill_ref, effect_ref, font_ref = value
        # -- defend against silently coerced non-int inputs at runtime even --
        # -- though the type signature requires ints                        --
        if not (
            isinstance(line_ref, int)  # pyright: ignore[reportUnnecessaryIsInstance]
            and isinstance(fill_ref, int)  # pyright: ignore[reportUnnecessaryIsInstance]
            and isinstance(effect_ref, int)  # pyright: ignore[reportUnnecessaryIsInstance]
        ):
            raise TypeError("line_ref, fill_ref, and effect_ref must be non-negative int values")
        if line_ref < 0 or fill_ref < 0 or effect_ref < 0:
            raise ValueError("line_ref, fill_ref, and effect_ref must be non-negative")
        if font_ref not in ("major", "minor", "none"):
            raise ValueError(
                "font_ref must be one of 'major', 'minor', or 'none';" f" got {font_ref!r}"
            )

        style_xml = (
            f"<p:style {nsdecls('a', 'p')}>\n"
            f'  <a:lnRef idx="{line_ref}">\n'
            f'    <a:schemeClr val="accent1"/>\n'
            f"  </a:lnRef>\n"
            f'  <a:fillRef idx="{fill_ref}">\n'
            f'    <a:schemeClr val="accent1"/>\n'
            f"  </a:fillRef>\n"
            f'  <a:effectRef idx="{effect_ref}">\n'
            f'    <a:schemeClr val="accent1"/>\n'
            f"  </a:effectRef>\n"
            f'  <a:fontRef idx="{font_ref}">\n'
            f'    <a:schemeClr val="lt1"/>\n'
            f"  </a:fontRef>\n"
            f"</p:style>"
        )
        new_style = cast("etree.ElementBase", parse_xml(style_xml))
        elm.insert_element_before(new_style, *successors_by_tag[elm.tag])

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


def _clone_dgm_relIds(
    clone_elm: ShapeElement,
    src_part: "Part",
    tgt_part: "Part",
) -> None:
    """Rewrite every ``dgm:relIds/@r:{dm,lo,qs,cs}`` on `clone_elm` against `tgt_part`.

    SmartArt diagram frames reference their four XML parts (data, layout,
    quickStyle, colors) via a ``dgm:relIds`` element whose attributes use
    the non-standard ``r:dm`` / ``r:lo`` / ``r:qs`` / ``r:cs`` names rather
    than the conventional ``r:id``. :class:`PartRelationshipCloner` only
    rewrites the three standard rId attributes, so any SmartArt rId on the
    clone still references an rId that lives on `src_part`. This helper
    walks every ``dgm:relIds`` found on `clone_elm` (or its descendants),
    looks each rId up on `src_part`, clones the referenced part into the
    target package (or reuses an existing part if already present), creates
    a matching relationship on `tgt_part`, and rewrites the attribute to
    the newly-allocated rId.

    Does nothing when the clone contains no SmartArt subtree.
    """
    relIds_tag = qn("dgm:relIds")
    dgm_attrs = (qn("r:dm"), qn("r:lo"), qn("r:qs"), qn("r:cs"))
    relIds_elms = clone_elm.xpath(".//dgm:relIds")
    if not relIds_elms:
        return

    for relIds in relIds_elms:
        if relIds.tag != relIds_tag:
            continue  # pragma: no cover -- xpath already filters by tag
        for attr in dgm_attrs:
            old_rId = relIds.get(attr)
            if old_rId is None:
                continue
            try:
                src_rel = src_part.rels[old_rId]
            except KeyError:  # pragma: no cover -- defensive: dangling rId
                continue
            if src_rel.is_external:  # pragma: no cover -- SmartArt rels are internal
                new_rId = tgt_part.relate_to(src_rel.target_ref, src_rel.reltype, is_external=True)
            else:
                tgt_target_part = _get_or_clone_sibling_part(src_rel.target_part, tgt_part)
                new_rId = tgt_part.relate_to(tgt_target_part, src_rel.reltype)
            relIds.set(attr, new_rId)


def _get_or_clone_sibling_part(src_target_part: "Part", tgt_part: "Part") -> "Part":
    """Return a Part mirroring `src_target_part` on `tgt_part`'s package.

    Reuses the source part when the two parts already share a package (same-
    package clone). Otherwise materialises a shallow duplicate of the source
    part with a non-colliding partname in the target package, preserving the
    content-type and blob bytes. Relationships *within* the target part are
    not recursed into — the SmartArt case needs a four-part shallow clone,
    not deep graph cloning.
    """
    from pptx.opc.package import (
        XmlPart,
        _partname_template_for,  # pyright: ignore[reportPrivateUsage]
    )

    tgt_package = tgt_part.package
    if src_target_part.package is tgt_package:
        return src_target_part

    partname_tmpl = _partname_template_for(src_target_part.partname)
    new_partname = tgt_package.next_partname(partname_tmpl)
    blob = src_target_part.blob
    src_cls = type(src_target_part)
    # -- XmlPart subclasses take a parsed `element` in their constructor; use
    # -- `load` to parse the blob into a fresh element tree on the target
    # -- package. Non-XML parts use the standard blob signature.
    if issubclass(src_cls, XmlPart):
        cloned = src_cls.load(new_partname, src_target_part.content_type, tgt_package, blob)
    else:
        cloned = src_cls(new_partname, src_target_part.content_type, tgt_package, blob)
    return cloned


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


class _CustomPropsDict:
    """Live, dict-like view over a shape's custom-props extension entries.

    Stores string-to-string metadata under the shape's ``p:cNvPr/a:extLst`` in
    a single ``a:ext`` whose ``uri`` is the fork's custom-props URI.
    Mutations write back to the XML tree immediately and survive save/reload
    round-trips.

    This proxy is intentionally low-level and is not exposed directly — callers
    obtain an instance via :attr:`BaseShape.custom_props`. Instances are cheap
    to create; each access re-resolves the underlying XML, so a proxy held
    across XML mutations on the shape remains valid.
    """

    # -- URI under which we stash the custom-props container inside the shape's --
    # -- `p:cNvPr/a:extLst`. The `urn:loadfix-pptx:…` scope prevents collision --
    # -- with any Microsoft or third-party extension uri.                       --
    _URI = "{urn:loadfix-pptx:custom-props:v1}"

    def __init__(self, cNvPr: CT_NonVisualDrawingProps):
        self._cNvPr = cNvPr

    def __contains__(self, key: object) -> bool:
        if not isinstance(key, str):
            return False
        return self._find_prop_elm(key) is not None

    def __delitem__(self, key: str) -> None:
        self._check_key(key)
        prop_elm = self._find_prop_elm(key)
        if prop_elm is None:
            raise KeyError(key)
        container = prop_elm.getparent()
        assert container is not None
        container.remove(prop_elm)
        # -- when the container is emptied, remove the wrapping `a:ext` and --
        # -- any now-orphan `a:extLst` so the on-disk XML stays tidy.       --
        self._prune_empty_containers(container)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, _CustomPropsDict):
            return self._cNvPr is other._cNvPr
        if isinstance(other, dict):
            return dict(self.items()) == other
        return NotImplemented

    def __getitem__(self, key: str) -> str:
        self._check_key(key)
        prop_elm = self._find_prop_elm(key)
        if prop_elm is None:
            raise KeyError(key)
        return _prop_value(prop_elm)

    def __iter__(self) -> Iterator[str]:
        return iter(self.keys())

    def __len__(self) -> int:
        container = self._container_elm
        if container is None:
            return 0
        return len(container.findall(qn("lfxcp:prop")))

    def __ne__(self, other: object) -> bool:
        result = self.__eq__(other)
        if result is NotImplemented:
            return NotImplemented
        return not result

    def __repr__(self) -> str:
        return f"{type(self).__name__}({dict(self.items())!r})"

    def __setitem__(self, key: str, value: str) -> None:
        self._check_key(key)
        if not isinstance(value, str):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(f"custom_props value must be str, not {type(value).__name__}")
        container = self._get_or_add_container_elm()
        existing = self._find_prop_in(container, key)
        if existing is not None:
            existing.text = value
            return
        prop_elm = etree.SubElement(container, qn("lfxcp:prop"))
        prop_elm.set("key", key)
        prop_elm.text = value

    def clear(self) -> None:
        """Remove all custom properties from this shape.

        Equivalent to iterating every key and deleting it, but runs in a
        single XML mutation. Removes the wrapping ``a:ext`` and ``a:extLst``
        when they would otherwise be left empty.
        """
        container = self._container_elm
        if container is None:
            return
        ext = container.getparent()
        assert ext is not None
        extLst = ext.getparent()
        assert extLst is not None
        extLst.remove(ext)
        if len(extLst.findall(qn("a:ext"))) == 0:
            parent_of_extLst = extLst.getparent()
            assert parent_of_extLst is not None
            parent_of_extLst.remove(extLst)

    def get(self, key: str, default: str | None = None) -> str | None:
        """Return the value for `key` if present, else `default`.

        Mirrors :meth:`dict.get`. Returns |None| (the default) when `key` is
        absent and no explicit `default` is supplied.
        """
        self._check_key(key)
        prop_elm = self._find_prop_elm(key)
        if prop_elm is None:
            return default
        return _prop_value(prop_elm)

    def items(self) -> list[tuple[str, str]]:
        """List of ``(key, value)`` tuples in insertion order."""
        container = self._container_elm
        if container is None:
            return []
        return [
            (str(p.get("key")), _prop_value(p))
            for p in container.findall(qn("lfxcp:prop"))
            if p.get("key") is not None
        ]

    def keys(self) -> list[str]:
        """List of keys in insertion order."""
        return [k for k, _ in self.items()]

    def values(self) -> list[str]:
        """List of values in insertion order."""
        return [v for _, v in self.items()]

    def _check_key(self, key: object) -> None:
        if not isinstance(key, str):
            raise TypeError(f"custom_props key must be str, not {type(key).__name__}")

    @property
    def _container_elm(self) -> etree._Element | None:
        """The `lfxcp:customProps` container, or |None| if absent."""
        extLst = self._cNvPr.extLst
        if extLst is None:
            return None
        ext = self._find_our_ext(extLst)
        if ext is None:
            return None
        return ext.find(qn("lfxcp:customProps"))

    def _find_our_ext(self, extLst: etree._Element) -> etree._Element | None:
        """Return our `a:ext` child of `extLst` (matching our uri), or |None|."""
        for ext in extLst.findall(qn("a:ext")):
            if ext.get("uri") == self._URI:
                return ext
        return None

    def _find_prop_elm(self, key: str) -> etree._Element | None:
        container = self._container_elm
        if container is None:
            return None
        return self._find_prop_in(container, key)

    @staticmethod
    def _find_prop_in(container: etree._Element, key: str) -> etree._Element | None:
        for prop in container.findall(qn("lfxcp:prop")):
            if prop.get("key") == key:
                return prop
        return None

    def _get_or_add_container_elm(self) -> etree._Element:
        """Return the `lfxcp:customProps` element, creating the chain if needed."""
        extLst = self._cNvPr.get_or_add_extLst()
        ext = self._find_our_ext(extLst)
        if ext is None:
            ext = etree.SubElement(extLst, qn("a:ext"))
            ext.set("uri", self._URI)
        container = ext.find(qn("lfxcp:customProps"))
        if container is None:
            # -- declare the fork namespace directly on the container so the --
            # -- element is self-describing when PowerPoint serializes it.   --
            container = etree.SubElement(
                ext,
                qn("lfxcp:customProps"),
                nsmap={"lfxcp": nsuri("lfxcp")},
            )
        return container

    def _prune_empty_containers(self, container: etree._Element) -> None:
        if len(container.findall(qn("lfxcp:prop"))) > 0:
            return
        ext = container.getparent()
        assert ext is not None
        extLst = ext.getparent()
        assert extLst is not None
        extLst.remove(ext)
        if len(extLst.findall(qn("a:ext"))) == 0:
            parent_of_extLst = extLst.getparent()
            assert parent_of_extLst is not None
            parent_of_extLst.remove(extLst)


def _prop_value(prop_elm: etree._Element) -> str:
    """Return the text content of a `lfxcp:prop` element as a |str|.

    An explicit empty string is preserved (i.e. ``<lfxcp:prop key="k"/>`` and
    ``<lfxcp:prop key="k"></lfxcp:prop>`` both map to ``""``).
    """
    return prop_elm.text or ""


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
