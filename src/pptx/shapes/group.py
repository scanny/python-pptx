"""GroupShape and related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.dml.effect import ShadowFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import PartRelationshipCloner
from pptx.oxml.ns import qn
from pptx.shapes.base import BaseShape, _unique_shape_name
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.action import ActionSetting
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.shapes.base import _ShapesParent
    from pptx.shapes.shapetree import GroupShapes
    from pptx.types import ProvidesPart


class GroupShape(BaseShape):
    """A shape that acts as a container for other shapes."""

    def __init__(self, grpSp: CT_GroupShape, parent: ProvidesPart):
        super().__init__(grpSp, parent)
        self._grpSp = grpSp

    @lazyproperty
    def click_action(self) -> ActionSetting:
        """Unconditionally raises `TypeError`.

        A group shape cannot have a click action or hover action.
        """
        raise TypeError("a group shape cannot have a click action")

    def duplicate(self) -> BaseShape:
        """Return a new :class:`GroupShape` that is a duplicate of this group.

        The new group is appended as a top-level shape on the same shape tree as this
        group (making it the topmost shape in z-order there) and occupies the same
        slide-relative position and size as the original — even when the original is
        itself nested inside one or more enclosing groups, in which case the clone is
        placed at the original's :attr:`~BaseShape.effective_left` /
        :attr:`~BaseShape.effective_top` / :attr:`~BaseShape.effective_width` /
        :attr:`~BaseShape.effective_height` rectangle.

        Every Part relationship (images, media, embedded charts, OLE parts, ...)
        referenced by shapes inside the group is cloned onto the same slide part via
        :class:`~pptx.opc.package.PartRelationshipCloner`, so the duplicate is a
        self-contained, fully-valid subtree.

        Every ``cNvPr/@id`` in the cloned subtree — including the new group itself
        and every shape nested inside — is reassigned a fresh, globally-unique shape
        id to preserve the "one id per slide" invariant. Names are left unchanged on
        inner shapes; the new group itself is given a unique name derived from the
        original's name.

        Fixes issue #1085.

        .. versionadded:: 2026.05.0
        """
        # -- `_parent` is a `_BaseShapes` (concretely `SlideShapes`/`GroupShapes`) --
        # -- exposing `_spTree`, `_next_shape_id`, and `_shape_factory`. The       --
        # -- `BaseShape` type signature is deliberately narrower (`ProvidesPart`), --
        # -- so narrow via a `TYPE_CHECKING`-only structural cast.                 --
        parent = cast("_ShapesParent", self._parent)
        spTree = parent._spTree

        # -- 1. Capture the source's slide-relative rectangle *before* any    --
        # --    mutation happens. When `self` is nested inside other groups,  --
        # --    effective_* composites the transform cascade so the values    --
        # --    are in slide coords; for a top-level group they equal the     --
        # --    raw a:off/a:ext already.                                      --
        eff_x = self.effective_left
        eff_y = self.effective_top
        eff_cx = self.effective_width
        eff_cy = self.effective_height

        # -- 2. Clone the XML subtree, re-materialising every referenced Part  --
        # --    (images, media, ole, ...) on this slide's part and rewriting   --
        # --    every r:id / r:embed / r:link attribute to the new rIds.       --
        new_grpSp = cast(
            "CT_GroupShape",
            PartRelationshipCloner.clone(self.part, self.part, self._element),
        )

        # -- 3. Attach the clone to the shape-tree *before* reassigning ids so --
        # --    `_next_shape_id` (which reads `spTree.max_shape_id + 1`) sees   --
        # --    each just-assigned id and yields the next free value on the    --
        # --    subsequent iteration. Appending before a trailing `p:extLst`   --
        # --    gives the clone topmost z-order on the spTree.                 --
        spTree.insert_element_before(new_grpSp, "p:extLst")

        # -- 4. Reassign every cNvPr/@id in the cloned subtree — including the --
        # --    top-level group's own id and every nested shape — so the       --
        # --    clone doesn't collide with its source. Every @id on the slide  --
        # --    must be unique.                                                --
        for cNvPr in new_grpSp.xpath(".//p:cNvPr"):
            cNvPr.id = parent._next_shape_id

        # -- 5. Give the top-level group a unique name derived from the        --
        # --    original's (inner-shape names are left alone).                 --
        new_grpSp._nvXxPr.cNvPr.name = _unique_shape_name(self.name, spTree)

        # -- 6. Place the clone at the source's slide-relative rectangle. The  --
        # --    clone's child coord system (a:chOff/a:chExt) is preserved from --
        # --    the source so nested children keep their local a:off/a:ext;   --
        # --    the outer off/ext rectangle is what maps that child space     --
        # --    onto the slide. Fixes issue #1085 specifically for the case    --
        # --    where the source is nested inside another group.               --
        if (
            eff_x is not None
            and eff_y is not None
            and eff_cx is not None
            and eff_cy is not None
        ):
            xfrm = new_grpSp.get_or_add_xfrm()
            xfrm.x = eff_x
            xfrm.y = eff_y
            xfrm.cx = eff_cx
            xfrm.cy = eff_cy

        return parent._shape_factory(new_grpSp)

    def ungroup(self) -> list[BaseShape]:
        """Dissolve this group, hoisting each child to the slide-level shape tree.

        Each direct child of this group is moved out of the enclosing ``p:grpSp``
        into the root ``p:spTree`` of the slide (or layout / master / notes-slide
        that owns this shape) at its slide-relative
        :attr:`~BaseShape.effective_left` /
        :attr:`~BaseShape.effective_top` /
        :attr:`~BaseShape.effective_width` /
        :attr:`~BaseShape.effective_height` rectangle. The now-empty group is
        removed and the freed shapes are returned in the same z-order they had
        inside the group (i.e. the first returned shape was backmost within the
        group). Every freed shape becomes a top-level sibling on the slide.

        Nested groups are handled transparently: when this group is itself a
        descendant of one or more enclosing groups, the cascade of group
        transforms (Wave 3, issue #925) is composited into each child's
        ``effective_*`` value before the hoist so the child keeps its rendered
        slide-relative position regardless of nesting depth. A child that is
        itself a :class:`GroupShape` is likewise hoisted whole — its own raw
        ``a:off`` / ``a:ext`` are rewritten to its effective slide-rectangle
        while its internal ``a:chOff`` / ``a:chExt`` are preserved so its
        descendants continue to render at their original slide positions; the
        caller may recurse by calling :meth:`ungroup` on that returned shape to
        flatten further.

        After this call returns this :class:`GroupShape` instance refers to an
        XML element that has been removed from the shape tree; further use of
        this instance is undefined.

        Fixes issue #730.
        """
        # -- 1. Walk the XML upward to the owning root `p:spTree`. Group shapes --
        # -- always live under a `p:spTree` eventually (nested groups have      --
        # -- `p:grpSp` ancestors in between); the cascade of enclosing groups   --
        # -- is exactly what `effective_*` already composites for us.           --
        grpSp = self._grpSp
        spTree_tag = qn("p:spTree")
        ancestor = grpSp.getparent()
        while ancestor is not None and ancestor.tag != spTree_tag:
            ancestor = ancestor.getparent()
        if ancestor is None:
            raise ValueError(
                "group shape is not attached to a shape tree; cannot ungroup a detached group."
            )
        root_spTree = cast("CT_GroupShape", ancestor)

        # -- 2. Resolve the slide-level shape-collection that owns `root_spTree` --
        # -- so we can build correctly-typed proxies and produce fresh unique    --
        # -- shape ids when needed. `self.part` is the owning BaseSlidePart;     --
        # -- dispatching on it yields the matching SlideShapes / LayoutShapes /  --
        # -- MasterShapes / NotesSlideShapes proxy.                              --
        owner_shapes = cast("_ShapesParent", _owner_shapes_for_part(self.part))

        # -- 3. Capture each direct child's effective (slide-relative) rectangle --
        # -- *before* any mutation. The effective_* cascade walks every          --
        # -- enclosing `p:grpSp` including this one, so the returned values are  --
        # -- the slide coordinates we want each child to land at.                --
        child_elms = list(grpSp.iter_shape_elms())
        effective_rects: list[tuple[int | None, int | None, int | None, int | None]] = []
        for child in child_elms:
            proxy = BaseShape(child, self._parent)
            effective_rects.append(
                (
                    None if proxy.effective_left is None else int(proxy.effective_left),
                    None if proxy.effective_top is None else int(proxy.effective_top),
                    None if proxy.effective_width is None else int(proxy.effective_width),
                    None if proxy.effective_height is None else int(proxy.effective_height),
                )
            )

        # -- 4. Move each child element onto `root_spTree` just before the      --
        # -- trailing `p:extLst` (if present), otherwise appended. This makes   --
        # -- every freed shape a top-level sibling on the slide. The insertion  --
        # -- order preserves the children's relative z-order; each freed shape  --
        # -- ends up frontmost among its prior siblings.                        --
        for child in child_elms:
            grpSp.remove(child)
            root_spTree.insert_element_before(child, "p:extLst")

        # -- 5. Rewrite each freed child's raw `a:off` / `a:ext` to its         --
        # -- effective slide-rectangle. Child groups keep their own             --
        # -- `a:chOff` / `a:chExt` unchanged so their descendants render at the --
        # -- same slide positions (since the enclosing-group scale that was     --
        # -- dropped by the ungroup now re-appears inside the child group's own --
        # -- `ext/chExt` ratio).                                                --
        for child, (eff_x, eff_y, eff_cx, eff_cy) in zip(child_elms, effective_rects):
            _apply_effective_rect(child, eff_x, eff_y, eff_cx, eff_cy)

        # -- 6. Remove the now-empty `p:grpSp` from its parent. --
        grpSp_parent = grpSp.getparent()
        if grpSp_parent is not None:
            grpSp_parent.remove(grpSp)

        # -- 7. Return the freed shapes as properly-typed proxies parented by  --
        # -- the slide-level shape collection.                                 --
        return [
            owner_shapes._shape_factory(child)  # pyright: ignore[reportPrivateUsage]
            for child in child_elms
        ]

    @property
    def has_text_frame(self) -> bool:
        """Unconditionally |False|.

        A group shape does not have a textframe and cannot itself contain text. This does not
        impact the ability of shapes contained by the group to each have their own text.
        """
        return False

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """|ShadowFormat| object representing shadow effect for this group.

        A |ShadowFormat| object is always returned, even when no shadow is explicitly defined on
        this group shape (i.e. when the group inherits its shadow behavior).
        """
        return ShadowFormat(self._grpSp.grpSpPr)

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Member of :ref:`MsoShapeType` identifying the type of this shape.

        Unconditionally `MSO_SHAPE_TYPE.GROUP` in this case
        """
        return MSO_SHAPE_TYPE.GROUP

    @lazyproperty
    def shapes(self) -> GroupShapes:
        """|GroupShapes| object for this group.

        The |GroupShapes| object provides access to the group's member shapes and provides methods
        for adding new ones.
        """
        from pptx.shapes.shapetree import GroupShapes

        return GroupShapes(self._element, self)


def _apply_effective_rect(
    shape_elm: ShapeElement,
    eff_x: int | None,
    eff_y: int | None,
    eff_cx: int | None,
    eff_cy: int | None,
) -> None:
    """Rewrite `shape_elm`'s raw ``a:off`` / ``a:ext`` to the given slide rectangle.

    Each value is written only when it is not |None|. Any of the four values
    that are |None| on the source leaves the corresponding attribute on the
    shape's own ``a:xfrm`` unchanged. The shape's own ``a:chOff`` / ``a:chExt``
    (present only on `p:grpSp` elements) are deliberately preserved so the
    child coordinate system of a freed child-group remains consistent with its
    own descendants' raw positions.
    """
    if eff_x is not None:
        shape_elm.x = eff_x
    if eff_y is not None:
        shape_elm.y = eff_y
    if eff_cx is not None:
        shape_elm.cx = eff_cx
    if eff_cy is not None:
        shape_elm.cy = eff_cy


def _owner_shapes_for_part(part: object) -> object:
    """Return the slide-level shape-collection proxy owning `part`'s ``p:spTree``.

    Dispatches on the concrete part class so the correct proxy subclass is
    returned — :class:`SlideShapes` for a slide, :class:`LayoutShapes` for a
    layout, :class:`MasterShapes` for a master, :class:`NotesSlideShapes` for
    a notes slide. This lets :meth:`GroupShape.ungroup` build correctly-typed
    shape proxies for children hoisted onto the root shape tree regardless of
    which kind of slide the group lives on.
    """
    from pptx.parts.slide import (
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    if isinstance(part, SlidePart):
        return part.slide.shapes
    if isinstance(part, SlideLayoutPart):
        return part.slide_layout.shapes
    if isinstance(part, SlideMasterPart):
        return part.slide_master.shapes
    if isinstance(part, NotesSlidePart):
        return part.notes_slide.shapes
    raise TypeError(
        "cannot locate slide-level shape tree for part of type %s" % type(part).__name__
    )
