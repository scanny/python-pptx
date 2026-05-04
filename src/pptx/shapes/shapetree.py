"""The shape tree, the structure that holds a slide's shapes."""

from __future__ import annotations

import io
import os
from typing import IO, TYPE_CHECKING, Callable, Iterable, Iterator, cast

from pptx.enum.shapes import PP_PLACEHOLDER, PROG_ID
from pptx.media import SPEAKER_IMAGE_BYTES, Video
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from pptx.oxml.shapes.picture import CT_Picture
from pptx.oxml.simpletypes import ST_Direction
from pptx.parts.chart import ChartPart
from pptx.shapes.autoshape import AutoShapeType, Shape
from pptx.shapes.base import BaseShape
from pptx.shapes.connector import Connector
from pptx.shapes.freeform import FreeformBuilder
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Movie, Picture
from pptx.shapes.placeholder import (
    ChartPlaceholder,
    LayoutPlaceholder,
    MasterPlaceholder,
    NotesSlidePlaceholder,
    PicturePlaceholder,
    PlaceholderGraphicFrame,
    PlaceholderPicture,
    SlidePlaceholder,
    TablePlaceholder,
)
from pptx.shared import ParentedElementProxy
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.chart.chart import Chart
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.connector import CT_Connector
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.parts.image import ImagePart
    from pptx.parts.slide import BaseSlidePart, SlidePart
    from pptx.slide import Slide, SlideLayout
    from pptx.types import ProvidesPart
    from pptx.util import Length

# -- Minimal 1x1 fully-transparent PNG used as the rasterized fallback for an
# -- editable SVG when the caller of ``add_picture_svg`` does not supply one.
# -- PowerPoint 365 picks the ``asvg:svgBlip`` extension and renders the SVG
# -- itself; only older clients (and the .pptx thumbnail preview) fall back to
# -- the raster. Sized 1x1 so it never distorts the surrounding layout when it
# -- does get drawn. See issue #358.
_SVG_PNG_FALLBACK_BYTES = (
    b"\x89PNG\r\n\x1a\n"
    b"\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
    b"\x00\x00\x00\rIDATx\x9cc\xf8\xcf\xc0\x00\x00\x00\x03\x00\x01\xde\xfb\xc4f"
    b"\x00\x00\x00\x00IEND\xaeB`\x82"
)


def _svg_desc(svg_file: str | IO[bytes]) -> str:
    """Return a filename-like description for `svg_file` suitable for ``p:cNvPr/@descr``.

    Falls back to ``"image.svg"`` when the source is a file-like object without
    a readable ``name`` attribute.
    """
    if isinstance(svg_file, str):
        return os.path.basename(svg_file) or "image.svg"
    name = getattr(svg_file, "name", None)
    if isinstance(name, str) and name:
        return os.path.basename(name)
    return "image.svg"


# +-- _BaseShapes
# |   |
# |   +-- _BaseGroupShapes
# |   |   |
# |   |   +-- GroupShapes
# |   |   |
# |   |   +-- SlideShapes
# |   |   |
# |   |   +-- LayoutShapes
# |   |   |
# |   |   +-- MasterShapes
# |   |
# |   +-- NotesSlideShapes
# |   |
# |   +-- BasePlaceholders
# |       |
# |       +-- LayoutPlaceholders
# |       |
# |       +-- MasterPlaceholders
# |           |
# |           +-- NotesSlidePlaceholders
# |
# +-- SlidePlaceholders


class _BaseShapes(ParentedElementProxy):
    """Base class for a shape collection appearing in a slide-type object.

    Subclasses include Slide, SlideLayout, and SlideMaster. Provides common methods.
    """

    def __init__(self, spTree: CT_GroupShape, parent: ProvidesPart):
        super(_BaseShapes, self).__init__(spTree, parent)
        self._spTree = spTree
        self._cached_max_shape_id = None

    def __getitem__(self, idx: int) -> BaseShape:
        """Return shape at `idx` in sequence, e.g. `shapes[2]`."""
        shape_elms = list(self._iter_member_elms())
        try:
            shape_elm = shape_elms[idx]
        except IndexError:
            raise IndexError("shape index out of range")
        return self._shape_factory(shape_elm)

    def __iter__(self) -> Iterator[BaseShape]:
        """Generate a reference to each shape in the collection, in sequence."""
        for shape_elm in self._iter_member_elms():
            yield self._shape_factory(shape_elm)

    def __len__(self) -> int:
        """Return count of shapes in this shape tree.

        A group shape contributes 1 to the total, without regard to the number of shapes contained
        in the group.
        """
        shape_elms = list(self._iter_member_elms())
        return len(shape_elms)

    def clear(self, preserve_placeholders: bool = True) -> None:
        """Remove every top-level shape from this shape tree.

        By default placeholders are preserved because removing them severs the layout
        inheritance that drives the slide's look and feel. Pass ``preserve_placeholders=False``
        to remove every shape including placeholders.

        Only the shapes enumerated by :attr:`__iter__` are considered — nested shapes inside
        a :class:`.GroupShape` are removed together with their enclosing group because the
        group element itself is detached from the shape tree. Non-shape children of
        ``p:spTree`` (for example ``p:nvGrpSpPr`` / ``p:grpSpPr``) are left in place.

        Subclass :meth:`~.BaseShape.delete` overrides run for each removed shape, so any
        per-shape side effects (dropping an image relationship for a picture, dropping an
        embedded chart part for a chart, etc.) execute in the normal way.

        Returns ``None`` to match :meth:`list.clear`.

        .. versionadded:: 2026.05.0
        """
        # -- snapshot first — mutating the tree during iteration would skip siblings --
        for shape in list(self):
            if preserve_placeholders and shape.is_placeholder:
                continue
            shape.delete()

    def descendants(self) -> Iterator[BaseShape]:
        """Generate each shape in this shape tree, including descendants of any group shape.

        This is the Selection-Pane-equivalent traversal: every |BaseShape| that appears
        in the slide's visual stacking — top-level shapes plus each shape nested inside
        any |GroupShape|, to arbitrary depth — is yielded in document (z-order)
        sequence. Group shapes themselves are yielded *before* their children so the
        caller sees the container alongside its contents (matching PowerPoint's
        Selection Pane list). For a slide that contains no groups this iterator yields
        the same sequence as ``iter(shapes)``.

        The iterator walks the XML tree lazily; side-effects on the returned proxies
        (e.g. ``shape.delete()``) during iteration are unsupported and may skip or
        revisit siblings. Materialize with ``list(shapes.descendants())`` when you
        need a stable snapshot.

        .. versionadded:: 2026.05.0
        """
        for shape_elm in self._iter_member_elms():
            shape = self._shape_factory(shape_elm)
            yield shape
            if isinstance(shape, GroupShape):
                yield from shape.shapes.descendants()

    def iter_leaf_shapes(self) -> Iterator[BaseShape]:
        """Generate each non-group descendant shape in document (z-order) sequence.

        This is a leaf-only variant of :meth:`descendants`: it yields every
        |BaseShape| that would be yielded by :meth:`descendants` *except* the
        |GroupShape| containers themselves. The result is the set of shapes
        that actually carry rendered content — autoshapes, pictures,
        connectors, graphic frames (tables, charts, OLE), and so on — flattened
        across every nesting depth. For a slide that contains no groups this
        iterator yields the same sequence as ``iter(shapes)``.

        Useful when you want to touch every drawable element once without
        having to special-case the group containers (e.g. collecting alt-text,
        auditing fills, exporting a shape-by-shape report).

        The iterator walks the XML tree lazily; side-effects on the returned
        proxies (e.g. ``shape.delete()``) during iteration are unsupported and
        may skip or revisit siblings. Materialize with
        ``list(shapes.iter_leaf_shapes())`` when you need a stable snapshot.

        .. versionadded:: 2026.05.0
        """
        return (s for s in self.descendants() if not isinstance(s, GroupShape))

    def clone_placeholder(self, placeholder: LayoutPlaceholder) -> None:
        """Add a new placeholder shape based on `placeholder`.

        The cloned placeholder preserves the `name` of the source layout placeholder when
        that name is not already in use on this shape tree. This allows user-customized
        placeholder names authored on the slide layout (e.g. "Agenda Title") to flow through
        to slides created from that layout. If the layout's placeholder name would collide
        with an existing shape on this tree, a unique name is generated from the standard
        placeholder basename instead (preserving the pre-existing behavior for that case).
        """
        sp = placeholder.element
        ph_type, orient, sz, idx = (sp.ph_type, sp.ph_orient, sp.ph_sz, sp.ph_idx)
        id_ = self._next_shape_id
        # --- preserve layout placeholder's `name` when available and non-colliding ---
        layout_name = sp.shape_name
        existing_names = self._spTree.xpath("//p:cNvPr/@name")
        if layout_name and layout_name not in existing_names:
            name = layout_name
        else:
            name = self._next_ph_name(ph_type, id_, orient)
        self._spTree.add_placeholder(id_, name, ph_type, orient, sz, idx)

    def ph_basename(self, ph_type: PP_PLACEHOLDER) -> str:
        """Return the base name for a placeholder of `ph_type` in this shape collection.

        There is some variance between slide types, for example a notes slide uses a different
        name for the body placeholder, so this method can be overriden by subclasses.
        """
        return {
            PP_PLACEHOLDER.BITMAP: "ClipArt Placeholder",
            PP_PLACEHOLDER.BODY: "Text Placeholder",
            PP_PLACEHOLDER.CENTER_TITLE: "Title",
            PP_PLACEHOLDER.CHART: "Chart Placeholder",
            PP_PLACEHOLDER.DATE: "Date Placeholder",
            PP_PLACEHOLDER.FOOTER: "Footer Placeholder",
            PP_PLACEHOLDER.HEADER: "Header Placeholder",
            PP_PLACEHOLDER.MEDIA_CLIP: "Media Placeholder",
            PP_PLACEHOLDER.OBJECT: "Content Placeholder",
            PP_PLACEHOLDER.ORG_CHART: "SmartArt Placeholder",
            PP_PLACEHOLDER.PICTURE: "Picture Placeholder",
            PP_PLACEHOLDER.SLIDE_NUMBER: "Slide Number Placeholder",
            PP_PLACEHOLDER.SUBTITLE: "Subtitle",
            PP_PLACEHOLDER.TABLE: "Table Placeholder",
            PP_PLACEHOLDER.TITLE: "Title",
        }[ph_type]

    @property
    def turbo_add_enabled(self) -> bool:
        """True if "turbo-add" mode is enabled. Read/Write.

        EXPERIMENTAL: This feature can radically improve performance when adding large numbers
        (hundreds of shapes) to a slide. It works by caching the last shape ID used and
        incrementing that value to assign the next shape id. This avoids repeatedly searching all
        shape ids in the slide each time a new ID is required.

        Performance is not noticeably improved for a slide with a relatively small number of
        shapes, but because the search time rises with the square of the shape count, this option
        can be useful for optimizing generation of a slide composed of many shapes.

        Shape-id collisions can occur (causing a repair error on load) if more than one |Slide|
        object is used to interact with the same slide in the presentation. Note that the |Slides|
        collection creates a new |Slide| object each time a slide is accessed (e.g. `slide =
        prs.slides[0]`, so you must be careful to limit use to a single |Slide| object.
        """
        return self._cached_max_shape_id is not None

    @turbo_add_enabled.setter
    def turbo_add_enabled(self, value: bool):
        enable = bool(value)
        self._cached_max_shape_id = self._spTree.max_shape_id if enable else None

    @staticmethod
    def _is_member_elm(shape_elm: ShapeElement) -> bool:
        """Return true if `shape_elm` represents a member of this collection, False otherwise."""
        return True

    def _iter_member_elms(self) -> Iterator[ShapeElement]:
        """Generate each child of the `p:spTree` element that corresponds to a shape.

        Items appear in XML document order.
        """
        for shape_elm in self._spTree.iter_shape_elms():
            if self._is_member_elm(shape_elm):
                yield shape_elm

    def _next_ph_name(self, ph_type: PP_PLACEHOLDER, id: int, orient: str) -> str:
        """Next unique placeholder name for placeholder shape of type `ph_type`.

        Usually will be standard placeholder root name suffixed with id-1, e.g.
        _next_ph_name(ST_PlaceholderType.TBL, 4, 'horz') ==> 'Table Placeholder 3'. The number is
        incremented as necessary to make the name unique within the collection. If `orient` is
        `'vert'`, the placeholder name is prefixed with `'Vertical '`.
        """
        basename = self.ph_basename(ph_type)

        # prefix rootname with 'Vertical ' if orient is 'vert'
        if orient == ST_Direction.VERT:
            basename = "Vertical %s" % basename

        # increment numpart as necessary to make name unique
        numpart = id - 1
        names = self._spTree.xpath("//p:cNvPr/@name")
        while True:
            name = "%s %d" % (basename, numpart)
            if name not in names:
                break
            numpart += 1

        return name

    @property
    def _next_shape_id(self) -> int:
        """Return a unique shape id suitable for use with a new shape.

        The returned id is 1 greater than the maximum shape id used so far. In practice, the
        minimum id is 2 because the spTree element is always assigned id="1".
        """
        # ---presence of cached-max-shape-id indicates turbo mode is on---
        if self._cached_max_shape_id is not None:
            self._cached_max_shape_id += 1
            return self._cached_max_shape_id

        return self._spTree.max_shape_id + 1

    def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape:
        """Return an instance of the appropriate shape proxy class for `shape_elm`."""
        return BaseShapeFactory(shape_elm, self)


class _BaseGroupShapes(_BaseShapes):
    """Base class for shape-trees that can add shapes."""

    part: BaseSlidePart  # pyright: ignore[reportIncompatibleMethodOverride]
    _element: CT_GroupShape

    def __init__(self, grpSp: CT_GroupShape, parent: ProvidesPart):
        super(_BaseGroupShapes, self).__init__(grpSp, parent)
        self._grpSp = grpSp

    def add_chart(
        self,
        chart_type: XL_CHART_TYPE,
        x: Length,
        y: Length,
        cx: Length,
        cy: Length,
        chart_data: ChartData,
    ) -> Chart:
        """Add a new chart of `chart_type` to the slide.

        The chart is positioned at (`x`, `y`), has size (`cx`, `cy`), and depicts `chart_data`.
        `chart_type` is one of the :ref:`XlChartType` enumeration values. `chart_data` is a
        |ChartData| object populated with the categories and series values for the chart.

        Note that a |GraphicFrame| shape object is returned, not the |Chart| object contained in
        that graphic frame shape. The chart object may be accessed using the :attr:`chart`
        property of the returned |GraphicFrame| object.
        """
        rId = self.part.add_chart_part(chart_type, chart_data)
        graphicFrame = self._add_chart_graphicFrame(rId, x, y, cx, cy)
        self._recalculate_extents()
        return cast("Chart", self._shape_factory(graphicFrame))

    def clone_chart(
        self,
        source_chart: Chart,
        x: Length,
        y: Length,
        cx: Length,
        cy: Length,
    ) -> GraphicFrame:
        """Duplicate `source_chart` onto this shape tree at (`x`, `y`) size (`cx`, `cy`).

        Implements the cross-slide chart-copy primitive (issue #877). The new
        chart is a structurally-independent duplicate of `source_chart`: a
        fresh :class:`ChartPart` is created on this slide's package containing
        a deep copy of the source's ``c:chartSpace`` XML, every relationship
        the source chart part carries (embedded xlsx, theme override, chart
        images, ...) is re-established on the new chart part, and the
        embedded workbook is duplicated into its own
        :class:`EmbeddedXlsxPart` so each chart retains an independent "Edit
        Data" workbook.

        `source_chart` may live in this same presentation (e.g. copying a
        chart from slide 1 to slide 5) or in a different presentation
        entirely (cross-presentation copy). In the cross-package case every
        referenced part — chart, xlsx, theme-override, image — is
        materialised in this package; no cross-package references are
        retained.

        Returns the newly-created |GraphicFrame| containing the duplicated
        chart; use its :attr:`~pptx.shapes.graphfrm.GraphicFrame.chart`
        property to reach the underlying |Chart| object.
        """
        new_chart_part = ChartPart.clone_from(source_chart.part, self.part.package)
        rId = self.part.relate_to(new_chart_part, RT.CHART)
        graphicFrame = self._add_chart_graphicFrame(rId, x, y, cx, cy)
        self._recalculate_extents()
        return cast("GraphicFrame", self._shape_factory(graphicFrame))

    def add_connector(
        self,
        connector_type: MSO_CONNECTOR_TYPE,
        begin_x: Length,
        begin_y: Length,
        end_x: Length,
        end_y: Length,
    ) -> Connector:
        """Add a newly created connector shape to the end of this shape tree.

        `connector_type` is a member of the :ref:`MsoConnectorType` enumeration and the end-point
        values are specified as EMU values. The returned connector is of type `connector_type` and
        has begin and end points as specified.
        """
        cxnSp = self._add_cxnSp(connector_type, begin_x, begin_y, end_x, end_y)
        self._recalculate_extents()
        return cast(Connector, self._shape_factory(cxnSp))

    def add_group_shape(self, shapes: Iterable[BaseShape] = ()) -> GroupShape:
        """Return a |GroupShape| object newly appended to this shape tree.

        The group shape is empty and must be populated with shapes using methods on its shape
        tree, available on its `.shapes` property. The position and extents of the group shape are
        determined by the shapes it contains; its position and extents are recalculated each time
        a shape is added to it.
        """
        shapes = tuple(shapes)
        grpSp = self._element.add_grpSp()
        for shape in shapes:
            grpSp.insert_element_before(
                shape._element, "p:extLst"  # pyright: ignore[reportPrivateUsage]
            )
        if shapes:
            grpSp.recalculate_extents()
        return cast(GroupShape, self._shape_factory(grpSp))

    def add_ole_object(
        self,
        object_file: str | IO[bytes],
        prog_id: PROG_ID | str,
        left: Length,
        top: Length,
        width: Length | None = None,
        height: Length | None = None,
        icon_file: str | IO[bytes] | None = None,
        icon_width: Length | None = None,
        icon_height: Length | None = None,
        extension: str | None = None,
    ) -> GraphicFrame:
        """Return newly-created GraphicFrame shape embedding `object_file`.

        The returned graphic-frame shape contains `object_file` as an embedded OLE object. It is
        displayed as an icon at `left`, `top` with size `width`, `height`. `width` and `height`
        may be omitted when `prog_id` is a member of `PROG_ID`, in which case the default icon
        size is used. This is advised for best appearance where applicable because it avoids an
        icon with a "stretched" appearance.

        `object_file` may either be a str path to a file or file-like object (such as
        `io.BytesIO`) containing the bytes of the object to be embedded (such as an Excel file).

        `prog_id` can be either a member of `pptx.enum.shapes.PROG_ID` or a str value like
        `"Adobe.Exchange.7"` determined by inspecting the XML generated by PowerPoint for an
        object of the desired type. The built-in members include well-known office file types
        (``DOCX``, ``PPTX``, ``XLSX``) as well as generic convenience members for common
        file formats (``ZIP``, ``PDF``, ``DOC``, ``HTML``). Any other progId accepted by
        PowerPoint may also be passed as a str.

        `icon_file` may either be a str path to an image file or a file-like object containing the
        image. The image provided will be displayed in lieu of the OLE object; double-clicking on
        the image opens the object (subject to operating-system limitations). The image file can
        be any supported image file. Those produced by PowerPoint itself are generally EMF and can
        be harvested from a PPTX package that embeds such an object. PNG and JPG also work fine.

        `icon_width` and `icon_height` are `Length` values (e.g. Emu() or Inches()) that describe
        the size of the icon image within the shape. These should be omitted unless a custom
        `icon_file` is provided. The dimensions must be discovered by inspecting the XML.
        Automatic resizing of the OLE-object shape can occur when the icon is double-clicked if
        these values are not as set by PowerPoint. This behavior may only manifest in the Windows
        version of PowerPoint.

        `extension` is an optional str file-extension (without the leading dot, e.g. ``"zip"``)
        used to name the embedded-object part inside the package, e.g.
        ``/ppt/embeddings/oleObject1.zip``. When omitted, the extension is derived from
        `object_file` if it is a str path, or from `prog_id` if it is a |PROG_ID| member;
        otherwise it defaults to ``"bin"``. Providing this argument explicitly is recommended
        when `object_file` is a file-like object and `prog_id` is an arbitrary str, so the
        embedded file is stored with a meaningful extension in the package.
        """
        graphicFrame = _OleObjectElementCreator.graphicFrame(
            self,
            self._next_shape_id,
            object_file,
            prog_id,
            left,
            top,
            width,
            height,
            icon_file,
            icon_width,
            icon_height,
            extension,
        )
        self._spTree.append(graphicFrame)
        self._recalculate_extents()
        return cast(GraphicFrame, self._shape_factory(graphicFrame))

    def add_picture(
        self,
        image_file: str | IO[bytes],
        left: Length,
        top: Length,
        width: Length | None = None,
        height: Length | None = None,
    ) -> Picture:
        """Add picture shape displaying image in `image_file`.

        `image_file` can be either a path to a file (a string) or a file-like object. The picture
        is positioned with its top-left corner at (`top`, `left`). If `width` and `height` are
        both |None|, the native size of the image is used. If only one of `width` or `height` is
        used, the unspecified dimension is calculated to preserve the aspect ratio of the image.
        If both are specified, the picture is stretched to fit, without regard to its native
        aspect ratio.

        The file-like form does not require a filename or a seekable stream; a non-seekable
        reader (e.g. one backed by a browser blob URL or an HTTP-upload stream) is fully
        materialized in memory so the image's format is still detected from its bytes. See
        issue #866.
        """
        image_part, rId = self.part.get_or_add_image_part(image_file)
        pic = self._add_pic_from_image_part(image_part, rId, left, top, width, height)
        self._recalculate_extents()
        return cast(Picture, self._shape_factory(pic))

    def add_shape(
        self, autoshape_type_id: MSO_SHAPE, left: Length, top: Length, width: Length, height: Length
    ) -> Shape:
        """Return new |Shape| object appended to this shape tree.

        `autoshape_type_id` is a member of :ref:`MsoAutoShapeType` e.g. `MSO_SHAPE.RECTANGLE`
        specifying the type of shape to be added. The remaining arguments specify the new shape's
        position and size.
        """
        autoshape_type = AutoShapeType(autoshape_type_id)
        sp = self._add_sp(autoshape_type, left, top, width, height)
        self._recalculate_extents()
        return cast(Shape, self._shape_factory(sp))

    def add_table(
        self, rows: int, cols: int, left: Length, top: Length, width: Length, height: Length
    ) -> GraphicFrame:
        """Add a |GraphicFrame| object containing a table.

        The table has the specified number of `rows` and `cols` and the specified position and
        size. `width` is the *total* width of the table (not the width of a single column); it is
        evenly distributed between the columns of the new table. Likewise, `height` is the *total*
        (authored) height of the table (not the height of a single row); it is evenly distributed
        between the rows, so each row is initially ``height // rows`` EMU tall. The last column
        and last row absorb any integer-division remainder so ``sum(col_widths) == width`` and
        ``sum(row_heights) == height``. Once PowerPoint opens the file its layout engine may
        grow individual rows to fit their text content — see the "Table height and row height"
        section of the user guide for the caveat. Note that the `.table` property on the returned
        |GraphicFrame| shape must be used to access the enclosed |Table| object.

        This method is available on |SlideShapes|, |LayoutShapes|, |MasterShapes|, and
        |GroupShapes| — a table may be authored directly inside a group. When added to a group
        the group's position and extents are recalculated to include the new table.
        """
        graphicFrame = self._add_graphicFrame_containing_table(rows, cols, left, top, width, height)
        self._recalculate_extents()
        return cast(GraphicFrame, self._shape_factory(graphicFrame))

    def add_textbox(self, left: Length, top: Length, width: Length, height: Length) -> Shape:
        """Return newly added text box shape appended to this shape tree.

        The text box is of the specified size, located at the specified position on the slide.
        """
        sp = self._add_textbox_sp(left, top, width, height)
        self._recalculate_extents()
        return cast(Shape, self._shape_factory(sp))

    def build_freeform(
        self, start_x: float = 0, start_y: float = 0, scale: tuple[float, float] | float = 1.0
    ) -> FreeformBuilder:
        """Return |FreeformBuilder| object to specify a freeform shape.

        The optional `start_x` and `start_y` arguments specify the starting pen position in local
        coordinates. They will be rounded to the nearest integer before use and each default to
        zero.

        The optional `scale` argument specifies the size of local coordinates proportional to
        slide coordinates (EMU). If the vertical scale is different than the horizontal scale
        (local coordinate units are "rectangular"), a pair of numeric values can be provided as
        the `scale` argument, e.g. `scale=(1.0, 2.0)`. In this case the first number is
        interpreted as the horizontal (X) scale and the second as the vertical (Y) scale.

        A convenient method for calculating scale is to divide a |Length| object by an equivalent
        count of local coordinate units, e.g. `scale = Inches(1)/1000` for 1000 local units per
        inch.
        """
        x_scale, y_scale = scale if isinstance(scale, tuple) else (scale, scale)

        return FreeformBuilder.new(self, start_x, start_y, x_scale, y_scale)

    def index(self, shape: BaseShape) -> int:
        """Return the index of `shape` in this sequence.

        Raises |ValueError| if `shape` is not in the collection.
        """
        shape_elms = list(self._element.iter_shape_elms())
        return shape_elms.index(shape.element)

    def _add_chart_graphicFrame(
        self, rId: str, x: Length, y: Length, cx: Length, cy: Length
    ) -> CT_GraphicalObjectFrame:
        """Return new `p:graphicFrame` element appended to this shape tree.

        The `p:graphicFrame` element has the specified position and size and refers to the chart
        part identified by `rId`.
        """
        shape_id = self._next_shape_id
        name = "Chart %d" % (shape_id - 1)
        graphicFrame = CT_GraphicalObjectFrame.new_chart_graphicFrame(
            shape_id, name, rId, x, y, cx, cy
        )
        self._element.insert_element_before(graphicFrame, "p:extLst")
        return graphicFrame

    def _add_graphicFrame_containing_table(
        self, rows: int, cols: int, x: Length, y: Length, cx: Length, cy: Length
    ) -> CT_GraphicalObjectFrame:
        """Return a newly added `p:graphicFrame` element containing a table as specified."""
        _id = self._next_shape_id
        name = "Table %d" % (_id - 1)
        graphicFrame = self._element.add_table(_id, name, rows, cols, x, y, cx, cy)
        return graphicFrame

    def _add_cxnSp(
        self,
        connector_type: MSO_CONNECTOR_TYPE,
        begin_x: Length,
        begin_y: Length,
        end_x: Length,
        end_y: Length,
    ) -> CT_Connector:
        """Return a newly-added `p:cxnSp` element as specified.

        The `p:cxnSp` element is for a connector of `connector_type` beginning at (`begin_x`,
        `begin_y`) and extending to (`end_x`, `end_y`).
        """
        id_ = self._next_shape_id
        name = "Connector %d" % (id_ - 1)

        flipH, flipV = begin_x > end_x, begin_y > end_y
        x, y = min(begin_x, end_x), min(begin_y, end_y)
        cx, cy = abs(end_x - begin_x), abs(end_y - begin_y)

        return self._element.add_cxnSp(id_, name, connector_type, x, y, cx, cy, flipH, flipV)

    def _add_pic_from_image_part(
        self,
        image_part: ImagePart,
        rId: str,
        x: Length,
        y: Length,
        cx: Length | None,
        cy: Length | None,
    ) -> CT_Picture:
        """Return a newly appended `p:pic` element as specified.

        The `p:pic` element displays the image in `image_part` with size and position specified by
        `x`, `y`, `cx`, and `cy`. The element is appended to the shape tree, causing it to be
        displayed first in z-order on the slide.
        """
        id_ = self._next_shape_id
        scaled_cx, scaled_cy = image_part.scale(cx, cy)
        name = "Picture %d" % (id_ - 1)
        desc = image_part.desc
        pic = self._grpSp.add_pic(id_, name, desc, rId, x, y, scaled_cx, scaled_cy)
        return pic

    def _add_sp(
        self, autoshape_type: AutoShapeType, x: Length, y: Length, cx: Length, cy: Length
    ) -> CT_Shape:
        """Return newly-added `p:sp` element as specified.

        `p:sp` element is of `autoshape_type` at position (`x`, `y`) and of size (`cx`, `cy`).
        """
        id_ = self._next_shape_id
        name = "%s %d" % (autoshape_type.basename, id_ - 1)
        sp = self._grpSp.add_autoshape(id_, name, autoshape_type.prst, x, y, cx, cy)
        return sp

    def _add_textbox_sp(self, x: Length, y: Length, cx: Length, cy: Length) -> CT_Shape:
        """Return newly-appended textbox `p:sp` element.

        Element has position (`x`, `y`) and size (`cx`, `cy`).
        """
        id_ = self._next_shape_id
        name = "TextBox %d" % (id_ - 1)
        sp = self._spTree.add_textbox(id_, name, x, y, cx, cy)
        return sp

    def _recalculate_extents(self) -> None:
        """Adjust position and size to incorporate all contained shapes.

        This would typically be called when a contained shape is added, removed, or its position
        or size updated.
        """
        # ---default behavior is to do nothing, GroupShapes overrides to
        #    produce the distinctive behavior of groups and subgroups.---
        pass


class GroupShapes(_BaseGroupShapes):
    """The sequence of child shapes belonging to a group shape.

    Note that this collection can itself contain a group shape, making this part of a recursive,
    tree data structure (acyclic graph).
    """

    def _recalculate_extents(self) -> None:
        """Adjust position and size to incorporate all contained shapes.

        This would typically be called when a contained shape is added, removed, or its position
        or size updated.
        """
        self._grpSp.recalculate_extents()


class SlideShapes(_BaseGroupShapes):
    """Sequence of shapes appearing on a slide.

    The first shape in the sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), index(), and iteration.
    """

    parent: Slide  # pyright: ignore[reportIncompatibleMethodOverride]

    def add_picture_svg(
        self,
        svg_file: str | IO[bytes],
        left: Length,
        top: Length,
        width: Length | None = None,
        height: Length | None = None,
        png_fallback: str | IO[bytes] | None = None,
    ) -> Picture:
        """Add picture shape referencing the editable SVG in `svg_file`.

        PowerPoint 365 (Office 2016+) stores SVG imagery as a *pair*: a
        rasterized PNG the renderer paints for compatibility, and the
        original SVG kept for editing. Both parts are embedded in the
        package and referenced from a single ``p:pic`` via an
        ``asvg:svgBlip`` extension on the ``a:blip`` that points at the
        PNG. See issue #358.

        Because python-pptx does not bundle an SVG rasterizer, the caller
        is expected to supply the PNG themselves via `png_fallback`. Pass
        a path (``str``) or a file-like object containing PNG bytes. When
        `png_fallback` is |None| a tiny built-in 1x1 transparent PNG is
        used — the SVG will still render correctly in PowerPoint 365 and
        newer, but older clients that fall back to the raster blip will
        show the placeholder. Supply a properly rendered PNG whenever
        appearance outside PowerPoint 365 matters.

        `left` and `top` position the top-left corner of the picture.
        `width` and `height` specify the display size; because the SVG
        has no PIL-measurable intrinsic size, the fallback when either is
        |None| is 1 inch (914400 EMU). Supply both explicitly to preserve
        the SVG's own aspect ratio.

        Returns the newly-added |Picture| shape.

        .. versionadded:: 2026.05.0
        """
        default = Emu(914400)
        cx = width if width is not None else default
        cy = height if height is not None else default

        slide_part = self.part
        _, svg_rId = slide_part.get_or_add_svg_part(svg_file)
        png_source = (
            png_fallback if png_fallback is not None else io.BytesIO(_SVG_PNG_FALLBACK_BYTES)
        )
        _, png_rId = slide_part.get_or_add_image_part(png_source)

        shape_id = self._next_shape_id
        name = "Picture %d" % (shape_id - 1)
        desc = _svg_desc(svg_file)
        pic = self._grpSp.add_pic_svg(shape_id, name, desc, png_rId, svg_rId, left, top, cx, cy)
        self._recalculate_extents()
        return cast(Picture, self._shape_factory(pic))

    def add_picture_link(
        self,
        url: str,
        left: Length,
        top: Length,
        width: Length | None = None,
        height: Length | None = None,
    ) -> Picture:
        """Add picture shape that *links* to the image at `url` instead of embedding it.

        An external relationship of type
        ``http://schemas.openxmlformats.org/officeDocument/2006/relationships/image``
        with ``Target=url`` and ``TargetMode="External"`` is added to the slide part and
        referenced via ``<a:blip r:link="rIdX"/>``. No image bytes are downloaded or stored
        in the package; PowerPoint resolves the URL when the slide is displayed.

        `url` is a string URL (typically ``http://`` or ``https://``). `left` and `top`
        specify the position of the top-left corner of the picture. `width` and `height`
        specify the display size; because the image is not available locally, no aspect-
        ratio calculation is possible. When either value is |None|, it defaults to 1 inch
        (914400 EMU). To preserve the image's native aspect ratio, callers should supply
        both `width` and `height` explicitly.

        .. versionadded:: 2026.05.0
        """
        default = Emu(914400)
        cx = width if width is not None else default
        cy = height if height is not None else default
        rId = self.part.relate_to(url, RT.IMAGE, is_external=True)
        shape_id = self._next_shape_id
        name = "Picture %d" % (shape_id - 1)
        pic = self._grpSp.add_pic_link(shape_id, name, url, rId, left, top, cx, cy)
        self._recalculate_extents()
        return cast(Picture, self._shape_factory(pic))

    def add_movie(
        self,
        movie_file: str | IO[bytes],
        left: Length,
        top: Length,
        width: Length,
        height: Length,
        poster_frame_image: str | IO[bytes] | None = None,
        mime_type: str = CT.VIDEO,
        autoplay: bool = False,
    ) -> GraphicFrame:
        """Return newly added movie shape containing video (or audio) in `movie_file`.

        **EXPERIMENTAL.** This method has important limitations:

        * The size must be specified; no auto-scaling such as that provided by :meth:`add_picture`
          is performed.
        * The MIME type of the media file should be specified, e.g. 'video/mp4' or
          'audio/mpeg'. The provided media file is not interrogated for its type. The MIME type
          `video/unknown` is used by default (and works fine in tests as of this writing).
        * A poster frame image must be provided, it cannot be automatically extracted from the
          media file. If no poster frame is provided, the default "media loudspeaker" image will
          be used.

        When `mime_type` begins with `'audio/'` (e.g. `'audio/mpeg'`, `'audio/wav'`), the
        resulting shape is an audio clip and the emitted XML uses `<a:audioFile>` in place of
        `<a:videoFile>`; otherwise a video clip is produced.

        Return a newly added movie (or audio) shape to the slide, positioned at (`left`, `top`),
        having size (`width`, `height`), and containing `movie_file`. Before the video is
        started, `poster_frame_image` is displayed as a placeholder for the video.

        When `autoplay` is |True|, the returned movie is configured to start automatically with
        the slide (``start_condition="withPrevious"``). The default of |False| preserves
        PowerPoint's click-to-play behavior (``start_condition="onClick"``). For finer control
        (e.g. ``"afterPrevious"`` or a delayed start), assign directly to
        :attr:`Movie.start_condition` / :attr:`Movie.start_time` on the returned shape.
        """
        movie_pic = _MoviePicElementCreator.new_movie_pic(
            self,
            self._next_shape_id,
            movie_file,
            left,
            top,
            width,
            height,
            poster_frame_image,
            mime_type,
        )
        self._spTree.append(movie_pic)
        self._add_video_timing(movie_pic)
        movie = cast(Movie, self._shape_factory(movie_pic))
        if autoplay:
            movie.start_condition = "withPrevious"
        return cast(GraphicFrame, movie)

    def add_movie_link(
        self,
        url: str,
        poster_frame_image: str | IO[bytes],
        left: Length,
        top: Length,
        width: Length,
        height: Length,
    ) -> Movie:
        """Return newly added URL-linked movie shape for the video at `url`.

        This emits a ``p:pic`` matching PowerPoint's "Insert Online Video"
        output: the video is referenced by an external-mode relationship and
        only the poster-frame image is embedded in the package. No media
        bytes are copied into the ``.pptx``.

        `url` is the external URL for the video. PowerPoint requires a URL
        the media player can load in-slide — for YouTube, use the
        ``https://www.youtube.com/embed/VIDEO_ID`` form (not the
        ``watch?v=…`` form). Other CDN hosts work when they serve the video
        directly over HTTPS and allow cross-origin embedding.

        `poster_frame_image` is a path to an image file or a binary
        file-like object; it is displayed in-slide before playback begins.
        Unlike :meth:`add_movie`, a poster image is **required** because
        there is no media part from which to derive a default.

        `left`, `top`, `width`, `height` specify the position and size of
        the shape in EMU (use :mod:`pptx.util` helpers like :class:`Inches`).

        .. versionadded:: 2026.05.0
        """
        slide_part = self.part
        # -- NoRels external URL relationship with VIDEO reltype --
        video_rId = slide_part.relate_to(url, RT.VIDEO, is_external=True)
        # -- embed poster PNG as a normal image part (internal rel) --
        _, poster_rId = slide_part.get_or_add_image_part(poster_frame_image)
        shape_id = self._next_shape_id
        shape_name = "Video %d" % (shape_id - 1)
        movie_pic = CT_Picture.new_video_link_pic(
            shape_id,
            shape_name,
            video_rId,
            poster_rId,
            left,
            top,
            width,
            height,
        )
        self._spTree.append(movie_pic)
        self._add_video_timing(movie_pic)
        return cast(Movie, self._shape_factory(movie_pic))

    def clone_layout_placeholders(self, slide_layout: SlideLayout) -> None:
        """Add placeholder shapes based on those in `slide_layout`.

        Z-order of placeholders is preserved. Latent placeholders (date, slide number, and footer)
        are not cloned.
        """
        for placeholder in slide_layout.iter_cloneable_placeholders():
            self.clone_placeholder(placeholder)

    def find_all_by_name(self, name: str, include_descendants: bool = False) -> list[BaseShape]:
        """Return all shapes in this collection whose ``name`` equals `name`.

        Shape names in PowerPoint are not required to be unique, so this method
        returns every match in z-order (backmost first, topmost last). Returns
        an empty list when no shape has the given name. By default the search
        is restricted to direct members of this slide's shape tree; shapes
        nested inside a group are not included.

        When `include_descendants` is |True| the search walks every group on
        the slide recursively so names assigned to shapes inside a group are
        also matched. Matches are returned in document (z-order) sequence,
        with any matched |GroupShape| appearing before its own matched
        children — the same order :meth:`descendants` yields.

        .. versionadded:: 2026.05.0
        """
        if include_descendants:
            return [shape for shape in self.descendants() if shape.name == name]
        return [shape for shape in self if shape.name == name]

    def get_by_name(self, name: str, include_descendants: bool = False) -> BaseShape | None:
        """Return the first shape in this collection whose ``name`` equals `name`.

        Returns |None| if no shape has that name. When several shapes share the
        same name, the one earliest in z-order (backmost) is returned; use
        :meth:`find_all_by_name` to retrieve every match. An explicit method is
        provided instead of overloading ``__getitem__`` because the subscript
        operator is already defined for integer indexing.

        When `include_descendants` is |True| the search extends into every
        group on the slide, matching shape names assigned inside groups as
        well. A matched |GroupShape| is returned in preference to any match
        that lives inside it (the container is walked before its contents).

        .. versionadded:: 2026.05.0
        """
        iterable = self.descendants() if include_descendants else iter(self)
        for shape in iterable:
            if shape.name == name:
                return shape
        return None

    @property
    def placeholders(self) -> SlidePlaceholders:
        """Sequence of placeholder shapes in this slide."""
        return self.parent.placeholders

    @property
    def title(self) -> Shape | None:
        """The title placeholder shape on the slide.

        |None| if the slide has no title placeholder.
        """
        for elm in self._spTree.iter_ph_elms():
            if elm.ph_idx == 0:
                return cast(Shape, self._shape_factory(elm))
        return None

    def _add_video_timing(self, pic: CT_Picture) -> None:
        """Add a `p:video` element under `p:sld/p:timing`.

        The element will refer to the specified `pic` element by its shape id, and cause the video
        play controls to appear for that video.
        """
        sld = self._spTree.xpath("/p:sld")[0]
        childTnLst = sld.get_or_add_childTnLst()
        childTnLst.add_video(pic.shape_id)

    def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape:
        """Return an instance of the appropriate shape proxy class for `shape_elm`."""
        return SlideShapeFactory(shape_elm, self)


class LayoutShapes(_BaseGroupShapes):
    """Sequence of shapes appearing on a slide layout.

    The first shape in the sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), index(), and iteration.

    Like |SlideShapes|, this collection supports adding non-placeholder shapes
    (e.g. pictures, text boxes, auto-shapes, connectors, and group shapes) via
    :meth:`add_picture`, :meth:`add_textbox`, :meth:`add_shape`,
    :meth:`add_connector`, and :meth:`add_group_shape`. This lets callers
    customize a slide layout with branding elements such as logos or
    background images that should appear on every slide using that layout.
    """

    def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape:
        """Return an instance of the appropriate shape proxy class for `shape_elm`."""
        return _LayoutShapeFactory(shape_elm, self)


class MasterShapes(_BaseGroupShapes):
    """Sequence of shapes appearing on a slide master.

    The first shape in the sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), and iteration.

    Like |SlideShapes|, this collection supports adding non-placeholder shapes
    (e.g. pictures, text boxes, auto-shapes, connectors, and group shapes) via
    :meth:`add_picture`, :meth:`add_textbox`, :meth:`add_shape`,
    :meth:`add_connector`, and :meth:`add_group_shape`. This lets callers
    customize the slide master with branding elements such as a client logo
    or background image that should appear on every slide in the
    presentation (see issue #575).
    """

    def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape:
        """Return an instance of the appropriate shape proxy class for `shape_elm`."""
        return _MasterShapeFactory(shape_elm, self)


class NotesSlideShapes(_BaseShapes):
    """Sequence of shapes appearing on a notes slide.

    The first shape in the sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), index(), and iteration.
    """

    def ph_basename(self, ph_type: PP_PLACEHOLDER) -> str:
        """Return the base name for a placeholder of `ph_type` in this shape collection.

        A notes slide uses a different name for the body placeholder and has some unique
        placeholder types, so this method overrides the default in the base class.
        """
        return {
            PP_PLACEHOLDER.BODY: "Notes Placeholder",
            PP_PLACEHOLDER.DATE: "Date Placeholder",
            PP_PLACEHOLDER.FOOTER: "Footer Placeholder",
            PP_PLACEHOLDER.HEADER: "Header Placeholder",
            PP_PLACEHOLDER.SLIDE_IMAGE: "Slide Image Placeholder",
            PP_PLACEHOLDER.SLIDE_NUMBER: "Slide Number Placeholder",
        }[ph_type]

    def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape:
        """Return appropriate shape object for `shape_elm` appearing on a notes slide."""
        return _NotesSlideShapeFactory(shape_elm, self)


class BasePlaceholders(_BaseShapes):
    """Base class for placeholder collections.

    Subclasses differentiate behaviors for a master, layout, and slide. By default, placeholder
    shapes are constructed using |BaseShapeFactory|. Subclasses should override
    :method:`_shape_factory` to use custom placeholder classes.
    """

    @staticmethod
    def _is_member_elm(shape_elm: ShapeElement) -> bool:
        """True if `shape_elm` is a placeholder shape, False otherwise."""
        return shape_elm.has_ph_elm


class LayoutPlaceholders(BasePlaceholders):
    """Sequence of |LayoutPlaceholder| instance for each placeholder shape on a slide layout."""

    __iter__: Callable[  # pyright: ignore[reportIncompatibleMethodOverride]
        [], Iterator[LayoutPlaceholder]
    ]

    def get(self, idx: int, default: LayoutPlaceholder | None = None) -> LayoutPlaceholder | None:
        """The first placeholder shape with matching `idx` value, or `default` if not found."""
        for placeholder in self:
            if placeholder.element.ph_idx == idx:
                return placeholder
        return default

    def _shape_factory(self, shape_elm: ShapeElement) -> BaseShape:
        """Return an instance of the appropriate shape proxy class for `shape_elm`."""
        return _LayoutShapeFactory(shape_elm, self)


class MasterPlaceholders(BasePlaceholders):
    """Sequence of MasterPlaceholder representing the placeholder shapes on a slide master."""

    __iter__: Callable[  # pyright: ignore[reportIncompatibleMethodOverride]
        [], Iterator[MasterPlaceholder]
    ]

    def get(self, ph_type: PP_PLACEHOLDER, default: MasterPlaceholder | None = None):
        """Return the first placeholder shape with type `ph_type` (e.g. 'body').

        Returns `default` if no such placeholder shape is present in the collection.
        """
        for placeholder in self:
            if placeholder.ph_type == ph_type:
                return placeholder
        return default

    def _shape_factory(  # pyright: ignore[reportIncompatibleMethodOverride]
        self, placeholder_elm: CT_Shape
    ) -> MasterPlaceholder:
        """Return an instance of the appropriate shape proxy class for `shape_elm`."""
        return cast(MasterPlaceholder, _MasterShapeFactory(placeholder_elm, self))


class NotesSlidePlaceholders(MasterPlaceholders):
    """Sequence of placeholder shapes on a notes slide."""

    __iter__: Callable[  # pyright: ignore[reportIncompatibleMethodOverride]
        [], Iterator[NotesSlidePlaceholder]
    ]

    def _shape_factory(  # pyright: ignore[reportIncompatibleMethodOverride]
        self, placeholder_elm: CT_Shape
    ) -> NotesSlidePlaceholder:
        """Return an instance of the appropriate placeholder proxy class for `placeholder_elm`."""
        return cast(NotesSlidePlaceholder, _NotesSlideShapeFactory(placeholder_elm, self))


class SlidePlaceholders(ParentedElementProxy):
    """Collection of placeholder shapes on a slide.

    Supports iteration, :func:`len`, and dictionary-style lookup on the `idx` value of the
    placeholders it contains.
    """

    _element: CT_GroupShape

    def __getitem__(self, idx: int):
        """Access placeholder shape having `idx`.

        Note that while this looks like list access, `idx` is actually a dictionary key and will
        raise |KeyError| if no placeholder with that `idx` value is in the collection. Use
        :meth:`get` for a non-raising lookup.
        """
        for e in self._element.iter_ph_elms():
            if e.ph_idx == idx:
                return SlideShapeFactory(e, self)
        raise KeyError("no placeholder on this slide with idx == %d" % idx)

    def __iter__(self):
        """Generate placeholder shapes in `idx` order."""
        ph_elms = sorted([e for e in self._element.iter_ph_elms()], key=lambda e: e.ph_idx)
        return (SlideShapeFactory(e, self) for e in ph_elms)

    def __len__(self) -> int:
        """Return count of placeholder shapes."""
        return len(list(self._element.iter_ph_elms()))

    def get(self, idx: int, default: BaseShape | None = None) -> BaseShape | None:
        """Return the placeholder shape having `idx`, or `default` if not found.

        This is the non-raising counterpart to `placeholders[idx]` and provides a convenient way
        to locate a specific placeholder (e.g. a picture placeholder with `idx == 1`) without
        needing a try/except. `idx` is the ``p:ph/@idx`` attribute value, not a list position.

        .. versionadded:: 2026.05.0
        """
        for e in self._element.iter_ph_elms():
            if e.ph_idx == idx:
                return SlideShapeFactory(e, self)
        return default


def BaseShapeFactory(shape_elm: ShapeElement, parent: ProvidesPart) -> BaseShape:
    """Return an instance of the appropriate shape proxy class for `shape_elm`."""
    tag = shape_elm.tag

    if isinstance(shape_elm, CT_Picture):
        mediaFiles = shape_elm.xpath(
            "./p:nvPicPr/p:nvPr/a:videoFile | ./p:nvPicPr/p:nvPr/a:audioFile"
        )
        if mediaFiles:
            return Movie(shape_elm, parent)
        return Picture(shape_elm, parent)

    shape_cls = {
        qn("p:cxnSp"): Connector,
        qn("p:grpSp"): GroupShape,
        qn("p:sp"): Shape,
        qn("p:graphicFrame"): GraphicFrame,
    }.get(tag, BaseShape)

    return shape_cls(shape_elm, parent)  # pyright: ignore[reportArgumentType]


def _LayoutShapeFactory(shape_elm: ShapeElement, parent: ProvidesPart) -> BaseShape:
    """Return appropriate shape object for `shape_elm` on a slide layout."""
    if isinstance(shape_elm, CT_Shape) and shape_elm.has_ph_elm:
        return LayoutPlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


def _MasterShapeFactory(shape_elm: ShapeElement, parent: ProvidesPart) -> BaseShape:
    """Return appropriate shape object for `shape_elm` on a slide master."""
    if isinstance(shape_elm, CT_Shape) and shape_elm.has_ph_elm:
        return MasterPlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


def _NotesSlideShapeFactory(shape_elm: ShapeElement, parent: ProvidesPart) -> BaseShape:
    """Return appropriate shape object for `shape_elm` on a notes slide."""
    if isinstance(shape_elm, CT_Shape) and shape_elm.has_ph_elm:
        return NotesSlidePlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


def _SlidePlaceholderFactory(shape_elm: ShapeElement, parent: ProvidesPart):
    """Return a placeholder shape of the appropriate type for `shape_elm`."""
    tag = shape_elm.tag
    if tag == qn("p:sp"):
        Constructor = {
            PP_PLACEHOLDER.BITMAP: PicturePlaceholder,
            PP_PLACEHOLDER.CHART: ChartPlaceholder,
            PP_PLACEHOLDER.PICTURE: PicturePlaceholder,
            PP_PLACEHOLDER.TABLE: TablePlaceholder,
        }.get(shape_elm.ph_type, SlidePlaceholder)
    elif tag == qn("p:graphicFrame"):
        Constructor = PlaceholderGraphicFrame
    elif tag == qn("p:pic"):
        Constructor = PlaceholderPicture
    else:
        Constructor = BaseShapeFactory
    return Constructor(shape_elm, parent)  # pyright: ignore[reportArgumentType]


def SlideShapeFactory(shape_elm: ShapeElement, parent: ProvidesPart) -> BaseShape:
    """Return appropriate shape object for `shape_elm` on a slide."""
    if shape_elm.has_ph_elm:
        return _SlidePlaceholderFactory(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class _MoviePicElementCreator(object):
    """Functional service object for creating a new movie p:pic element.

    It's entire external interface is its :meth:`new_movie_pic` class method that returns a new
    `p:pic` element containing the specified video. This class is not intended to be constructed
    or an instance of it retained by the caller; it is a "one-shot" object, really a function
    wrapped in a object such that its helper methods can be organized here.
    """

    def __init__(
        self,
        shapes: SlideShapes,
        shape_id: int,
        movie_file: str | IO[bytes],
        x: Length,
        y: Length,
        cx: Length,
        cy: Length,
        poster_frame_file: str | IO[bytes] | None,
        mime_type: str | None,
    ):
        super(_MoviePicElementCreator, self).__init__()
        self._shapes = shapes
        self._shape_id = shape_id
        self._movie_file = movie_file
        self._x, self._y, self._cx, self._cy = x, y, cx, cy
        self._poster_frame_file = poster_frame_file
        self._mime_type = mime_type

    @classmethod
    def new_movie_pic(
        cls,
        shapes: SlideShapes,
        shape_id: int,
        movie_file: str | IO[bytes],
        x: Length,
        y: Length,
        cx: Length,
        cy: Length,
        poster_frame_image: str | IO[bytes] | None,
        mime_type: str | None,
    ) -> CT_Picture:
        """Return a new `p:pic` element containing video in `movie_file`.

        If `mime_type` is None, 'video/unknown' is used. If `poster_frame_file` is None, the
        default "media loudspeaker" image is used.
        """
        return cls(shapes, shape_id, movie_file, x, y, cx, cy, poster_frame_image, mime_type)._pic

    @property
    def _media_rId(self) -> str:
        """Return the rId of RT.MEDIA relationship to video part.

        For historical reasons, there are two relationships to the same part; one is the video rId
        and the other is the media rId.
        """
        return self._video_part_rIds[0]

    @lazyproperty
    def _pic(self) -> CT_Picture:
        """Return the new `p:pic` element referencing the video or audio media."""
        return CT_Picture.new_video_pic(
            self._shape_id,
            self._shape_name,
            self._video_rId,
            self._media_rId,
            self._poster_frame_rId,
            self._x,
            self._y,
            self._cx,
            self._cy,
            is_audio=self._is_audio,
        )

    @property
    def _is_audio(self) -> bool:
        """Return True when the media file is audio rather than video.

        Determined from the MIME-type prefix, e.g. 'audio/mpeg' -> audio,
        'video/mp4' -> video. Unknown MIME-types default to video, preserving
        backward-compatibility with the `mime_type=None` case.
        """
        mime_type = self._mime_type or ""
        return mime_type.lower().startswith("audio/")

    @lazyproperty
    def _poster_frame_image_file(self) -> str | IO[bytes]:
        """Return the image file for video placeholder image.

        If no poster frame file is provided, the default "media loudspeaker" image is used.
        """
        poster_frame_file = self._poster_frame_file
        if poster_frame_file is None:
            return io.BytesIO(SPEAKER_IMAGE_BYTES)
        return poster_frame_file

    @lazyproperty
    def _poster_frame_rId(self) -> str:
        """Return the rId of relationship to poster frame image.

        The poster frame is the image used to represent the video before it's played.
        """
        _, poster_frame_rId = self._slide_part.get_or_add_image_part(self._poster_frame_image_file)
        return poster_frame_rId

    @property
    def _shape_name(self) -> str:
        """Return the appropriate shape name for the p:pic shape.

        A movie shape is named with the base filename of the video, with the file
        extension stripped. This matches PowerPoint's own behavior in the selection
        pane (e.g. ``"intro.mp4"`` becomes ``"intro"``).
        """
        return os.path.splitext(self._video.filename)[0]

    @property
    def _slide_part(self) -> SlidePart:
        """Return SlidePart object for slide containing this movie."""
        return self._shapes.part

    @lazyproperty
    def _video(self) -> Video:
        """Return a |Video| object containing the movie file."""
        return Video.from_path_or_file_like(self._movie_file, self._mime_type)

    @lazyproperty
    def _video_part_rIds(self) -> tuple[str, str]:
        """Return the rIds for relationships to media part for video.

        This is where the media part and its relationships to the slide are actually created.
        """
        media_rId, video_rId = self._slide_part.get_or_add_video_media_part(self._video)
        return media_rId, video_rId

    @property
    def _video_rId(self) -> str:
        """Return the rId of RT.VIDEO relationship to video part.

        For historical reasons, there are two relationships to the same part; one is the video rId
        and the other is the media rId.
        """
        return self._video_part_rIds[1]


class _OleObjectElementCreator(object):
    """Functional service object for creating a new OLE-object p:graphicFrame element.

    It's entire external interface is its :meth:`graphicFrame` class method that returns a new
    `p:graphicFrame` element containing the specified embedded OLE-object shape. This class is not
    intended to be constructed or an instance of it retained by the caller; it is a "one-shot"
    object, really a function wrapped in a object such that its helper methods can be organized
    here.
    """

    def __init__(
        self,
        shapes: _BaseGroupShapes,
        shape_id: int,
        ole_object_file: str | IO[bytes],
        prog_id: PROG_ID | str,
        x: Length,
        y: Length,
        cx: Length | None,
        cy: Length | None,
        icon_file: str | IO[bytes] | None,
        icon_width: Length | None,
        icon_height: Length | None,
        extension: str | None = None,
    ):
        self._shapes = shapes
        self._shape_id = shape_id
        self._ole_object_file = ole_object_file
        self._prog_id_arg = prog_id
        self._x = x
        self._y = y
        self._cx_arg = cx
        self._cy_arg = cy
        self._icon_file_arg = icon_file
        self._icon_width_arg = icon_width
        self._icon_height_arg = icon_height
        self._extension_arg = extension

    @classmethod
    def graphicFrame(
        cls,
        shapes: _BaseGroupShapes,
        shape_id: int,
        ole_object_file: str | IO[bytes],
        prog_id: PROG_ID | str,
        x: Length,
        y: Length,
        cx: Length | None,
        cy: Length | None,
        icon_file: str | IO[bytes] | None,
        icon_width: Length | None,
        icon_height: Length | None,
        extension: str | None = None,
    ) -> CT_GraphicalObjectFrame:
        """Return new `p:graphicFrame` element containing embedded `ole_object_file`."""
        return cls(
            shapes,
            shape_id,
            ole_object_file,
            prog_id,
            x,
            y,
            cx,
            cy,
            icon_file,
            icon_width,
            icon_height,
            extension,
        )._graphicFrame

    @lazyproperty
    def _graphicFrame(self) -> CT_GraphicalObjectFrame:
        """Newly-created `p:graphicFrame` element referencing embedded OLE-object."""
        return CT_GraphicalObjectFrame.new_ole_object_graphicFrame(
            self._shape_id,
            self._shape_name,
            self._ole_object_rId,
            self._progId,
            self._icon_rId,
            self._x,
            self._y,
            self._cx,
            self._cy,
            self._icon_width,
            self._icon_height,
        )

    @lazyproperty
    def _cx(self) -> Length:
        """Emu object specifying width of "show-as-icon" image for OLE shape."""
        # --- a user-specified width overrides any default ---
        if self._cx_arg is not None:
            return self._cx_arg

        # --- the default width is specified by the PROG_ID member if prog_id is one,
        # --- otherwise it gets the default icon width.
        return (
            Emu(self._prog_id_arg.width) if isinstance(self._prog_id_arg, PROG_ID) else Emu(965200)
        )

    @lazyproperty
    def _cy(self) -> Length:
        """Emu object specifying height of "show-as-icon" image for OLE shape."""
        # --- a user-specified width overrides any default ---
        if self._cy_arg is not None:
            return self._cy_arg

        # --- the default height is specified by the PROG_ID member if prog_id is one,
        # --- otherwise it gets the default icon height.
        return (
            Emu(self._prog_id_arg.height) if isinstance(self._prog_id_arg, PROG_ID) else Emu(609600)
        )

    @lazyproperty
    def _icon_height(self) -> Length:
        """Vertical size of enclosed EMF icon within the OLE graphic-frame.

        This must be specified when a custom icon is used, to avoid stretching of the image and
        possible undesired resizing by PowerPoint when the OLE shape is double-clicked to open it.

        The correct size can be determined by creating an example PPTX using PowerPoint and then
        inspecting the XML of the OLE graphics-frame (p:oleObj.imgH).
        """
        return self._icon_height_arg if self._icon_height_arg is not None else Emu(609600)

    @lazyproperty
    def _icon_image_file(self) -> str | IO[bytes]:
        """Reference to image file containing icon to show in lieu of this object.

        This can be either a str path or a file-like object (io.BytesIO typically).
        """
        # --- a user-specified icon overrides any default ---
        if self._icon_file_arg is not None:
            return self._icon_file_arg

        # --- A prog_id belonging to PROG_ID gets its icon filename from there. A
        # --- user-specified (str) prog_id gets the default icon.
        icon_filename = (
            self._prog_id_arg.icon_filename
            if isinstance(self._prog_id_arg, PROG_ID)
            else "generic-icon.emf"
        )

        _thisdir = os.path.split(__file__)[0]
        return os.path.abspath(os.path.join(_thisdir, "..", "templates", icon_filename))

    @lazyproperty
    def _icon_rId(self) -> str:
        """str rId like "rId7" of rel to icon (image) representing OLE-object part."""
        _, rId = self._slide_part.get_or_add_image_part(self._icon_image_file)
        return rId

    @lazyproperty
    def _icon_width(self) -> Length:
        """Width of enclosed EMF icon within the OLE graphic-frame.

        This must be specified when a custom icon is used, to avoid stretching of the image and
        possible undesired resizing by PowerPoint when the OLE shape is double-clicked to open it.
        """
        return self._icon_width_arg if self._icon_width_arg is not None else Emu(965200)

    @lazyproperty
    def _extension(self) -> str | None:
        """Optional str file-extension (without dot) for the embedded-object part-name.

        Resolution order:

        * an explicit `extension` argument to `add_ole_object()` wins,
        * otherwise, when `ole_object_file` is a str path, its filename extension is used,
        * otherwise, `None` is returned (the factory will fall back to "bin" or the
          extension carried by a |PROG_ID| member).
        """
        if self._extension_arg is not None:
            return self._extension_arg
        if isinstance(self._ole_object_file, str):
            _, ext = os.path.splitext(self._ole_object_file)
            if ext.startswith("."):
                ext = ext[1:]
            return ext or None
        return None

    @lazyproperty
    def _ole_object_rId(self) -> str:
        """str rId like "rId6" of relationship to embedded ole_object part.

        This is where the ole_object part and its relationship to the slide are actually created.
        """
        return self._slide_part.add_embedded_ole_object_part(
            self._prog_id_arg, self._ole_object_file, self._extension
        )

    @lazyproperty
    def _progId(self) -> str:
        """str like "Excel.Sheet.12" identifying program used to open object.

        This value appears in the `progId` attribute of the `p:oleObj` element for the object.
        """
        prog_id_arg = self._prog_id_arg

        # --- member of PROG_ID enumeration knows its progId keyphrase, otherwise caller
        # --- has specified it explicitly (as str)
        return prog_id_arg.progId if isinstance(prog_id_arg, PROG_ID) else prog_id_arg

    @lazyproperty
    def _shape_name(self) -> str:
        """str name like "Object 1" for the embedded ole_object shape.

        The name is formed from the prefix "Object " and the shape-id decremented by 1.
        """
        return "Object %d" % (self._shape_id - 1)

    @lazyproperty
    def _slide_part(self) -> SlidePart:
        """SlidePart object for this slide."""
        return self._shapes.part
