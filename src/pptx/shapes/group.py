"""GroupShape and related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.dml.effect import ShadowFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import PartRelationshipCloner
from pptx.shapes.base import BaseShape, _unique_shape_name
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.action import ActionSetting
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
