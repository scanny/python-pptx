"""Base shape-related objects such as BaseShape."""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, cast

from pptx.action import ActionSetting
from pptx.dml.effect import ShadowFormat
from pptx.oxml.ns import qn
from pptx.shared import ElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from typing import Protocol

    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.oxml.shapes.shared import CT_Placeholder
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

    def duplicate(self) -> BaseShape:
        """Return a new shape that is a duplicate of this shape.

        The new shape is appended to the end of the same shape tree as this shape (making it
        topmost in z-order) and is an exact copy of this shape's XML, except that it is assigned
        a new unique shape-id and a new unique name.

        Only simple shapes are supported in this implementation: auto-shapes, text-boxes, and
        connectors. Duplicating a picture, chart, table, group-shape, or media shape raises
        `NotImplementedError` because those shapes own one or more package-relationships (to an
        image, embedded chart or xlsx part, etc.) that must also be copied for the duplicate to be
        valid; that work is deferred to a follow-up.
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
