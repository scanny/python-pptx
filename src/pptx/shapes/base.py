"""Base shape-related objects such as BaseShape."""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, cast

from lxml import etree

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
    def has_math_equation(self) -> bool:
        """|True| if this shape contains an Office Math (OMML) equation.

        An equation is detected by the presence of one or more `m:oMath` descendants
        anywhere within the shape's XML subtree. This covers both the `m:oMath` that
        appears directly under an `a14:m` wrapper (PowerPoint 2010+ equation shapes
        surfaced via an `mc:AlternateContent/mc:Choice` element) and the `m:oMath` that
        appears inside an `m:oMathPara` container.

        This is a read-only, MVP-scope property. Writing OMML and converting to/from
        LaTeX are explicitly deferred (see `math_equation_xml`).
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
