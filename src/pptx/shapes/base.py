"""Base shape-related objects such as BaseShape."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.action import ActionSetting
from pptx.dml.effect import ShadowFormat
from pptx.shared import ElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.oxml.shapes.shared import CT_Placeholder
    from pptx.parts.slide import BaseSlidePart
    from pptx.types import ProvidesPart
    from pptx.util import Length


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
        """
        siblings = self._zorder_siblings
        if siblings[-1] is self._element:
            return
        parent = self._element.getparent()
        assert parent is not None
        parent.remove(self._element)
        # -- insert before a trailing p:extLst if present, otherwise append --
        parent.insert_element_before(self._element, "p:extLst")

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
        """
        self._element.getparent().remove(self._element)

    @property
    def element(self) -> ShapeElement:
        """`lxml` element for this shape, e.g. a CT_Shape instance.

        Note that manipulating this element improperly can produce an invalid presentation file.
        Make sure you know what you're doing if you use this to change the underlying XML.
        """
        return self._element

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
        """
        return self._zorder_siblings.index(self._element)

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
