"""Placeholder-related objects.

Specific to shapes having a `p:ph` element. A placeholder has distinct behaviors
depending on whether it appears on a slide, layout, or master. Hence there is a
non-trivial class inheritance structure.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from pptx.oxml.shapes.picture import CT_Picture
from pptx.shapes.autoshape import Shape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.util import Emu

if TYPE_CHECKING:
    from pptx.oxml.shapes.autoshape import CT_Shape


class _InheritsDimensions(object):
    """
    Mixin class that provides inherited dimension behavior. Specifically,
    left, top, width, and height report the value from the layout placeholder
    where they would have otherwise reported |None|. This behavior is
    distinctive to placeholders. :meth:`_base_placeholder` must be overridden
    by all subclasses to provide lookup of the appropriate base placeholder
    to inherit from.
    """

    @property
    def height(self):
        """
        The effective height of this placeholder shape; its directly-applied
        height if it has one, otherwise the height of its parent layout
        placeholder.
        """
        return self._effective_value("height")

    @height.setter
    def height(self, value):
        self._element.cy = value

    @property
    def left(self):
        """
        The effective left of this placeholder shape; its directly-applied
        left if it has one, otherwise the left of its parent layout
        placeholder.
        """
        return self._effective_value("left")

    @left.setter
    def left(self, value):
        self._element.x = value

    @property
    def shape_type(self):
        """
        Member of :ref:`MsoShapeType` specifying the type of this shape.
        Unconditionally ``MSO_SHAPE_TYPE.PLACEHOLDER`` in this case.
        Read-only.
        """
        return MSO_SHAPE_TYPE.PLACEHOLDER

    @property
    def top(self):
        """
        The effective top of this placeholder shape; its directly-applied
        top if it has one, otherwise the top of its parent layout
        placeholder.
        """
        return self._effective_value("top")

    @top.setter
    def top(self, value):
        self._element.y = value

    @property
    def width(self):
        """
        The effective width of this placeholder shape; its directly-applied
        width if it has one, otherwise the width of its parent layout
        placeholder.
        """
        return self._effective_value("width")

    @width.setter
    def width(self, value):
        self._element.cx = value

    @property
    def _base_placeholder(self):
        """
        Return the layout or master placeholder shape this placeholder
        inherits from. Not to be confused with an instance of
        |BasePlaceholder| (necessarily).
        """
        raise NotImplementedError("Must be implemented by all subclasses.")

    def _effective_value(self, attr_name):
        """
        The effective value of *attr_name* on this placeholder shape; its
        directly-applied value if it has one, otherwise the value on the
        layout placeholder it inherits from.
        """
        directly_applied_value = getattr(super(_InheritsDimensions, self), attr_name)
        if directly_applied_value is not None:
            return directly_applied_value
        return self._inherited_value(attr_name)

    def _inherited_value(self, attr_name):
        """
        Return the attribute value, e.g. 'width' of the base placeholder this
        placeholder inherits from.
        """
        base_placeholder = self._base_placeholder
        if base_placeholder is None:
            return None
        inherited_value = getattr(base_placeholder, attr_name)
        return inherited_value


class _BaseSlidePlaceholder(_InheritsDimensions, Shape):
    """Base class for placeholders on slides.

    Provides common behaviors such as inherited dimensions.
    """

    @property
    def is_placeholder(self):
        """
        Boolean indicating whether this shape is a placeholder.
        Unconditionally |True| in this case.
        """
        return True

    @property
    def shape_type(self):
        """
        Member of :ref:`MsoShapeType` specifying the type of this shape.
        Unconditionally ``MSO_SHAPE_TYPE.PLACEHOLDER`` in this case.
        Read-only.
        """
        return MSO_SHAPE_TYPE.PLACEHOLDER

    def insert_chart(self, chart_type, chart_data):
        """Return a |PlaceholderGraphicFrame| object containing a new chart.

        The new chart is of `chart_type`, depicts `chart_data`, and has the same position
        and size as this placeholder. `chart_type` is one of the :ref:`XlChartType`
        enumeration values. `chart_data` is a |ChartData| object populated with the
        categories and series values for the chart. Note that the new |Chart| object is
        not returned directly. The chart object may be accessed using the
        :attr:`~.PlaceholderGraphicFrame.chart` property of the returned
        |PlaceholderGraphicFrame| object. Like other placeholder rich-content insertion
        methods, this placeholder reference becomes invalid after this call; use the
        return value (or re-fetch the placeholder by idx) to interact with the new shape.
        """
        rId = self.part.add_chart_part(chart_type, chart_data)
        graphicFrame = self._new_chart_graphicFrame(
            rId, self.left, self.top, self.width, self.height
        )
        self._replace_placeholder_with(graphicFrame)
        return PlaceholderGraphicFrame(graphicFrame, self._parent)

    def insert_picture(self, image_file, crop=True):
        """Return a |PlaceholderPicture| object depicting the image in `image_file`.

        `image_file` may be either a path (string) or a file-like object.

        When `crop` is |True| (the default) the image is stretched proportionately and
        cropped to fill the entire space of the placeholder. This matches PowerPoint's
        "Fill" picture layout and preserves the prior behavior of this method.

        When `crop` is |False| the image is scaled to fit entirely inside the
        placeholder, preserving its aspect ratio without any cropping, and centered
        horizontally and vertically within the placeholder bounds. This matches
        PowerPoint's "Fit" picture layout and is useful when the image content must
        remain fully visible (e.g. a logo, chart, or screenshot).

        A |PlaceholderPicture| object has all the properties and methods of a |Picture|
        shape except that the value of its
        :attr:`~._BaseSlidePlaceholder.shape_type` property is
        `MSO_SHAPE_TYPE.PLACEHOLDER` instead of `MSO_SHAPE_TYPE.PICTURE`.

        This method is available on all slide placeholders (inherited from
        :class:`._BaseSlidePlaceholder`), not only on a specialized
        :class:`.PicturePlaceholder`. This lets callers drop an image into a generic
        "content", "body", or "object" placeholder without having to change the
        placeholder type first.
        """
        pic = self._new_placeholder_pic(image_file, crop=crop)
        self._replace_placeholder_with(pic)
        return PlaceholderPicture(pic, self._parent)

    def insert_table(self, rows, cols):
        """Return |PlaceholderGraphicFrame| object containing a `rows` by `cols` table.

        The position and width of the table are those of the placeholder and its height
        is proportional to the number of rows. A |PlaceholderGraphicFrame| object has
        all the properties and methods of a |GraphicFrame| shape except that the value
        of its :attr:`~._BaseSlidePlaceholder.shape_type` property is unconditionally
        `MSO_SHAPE_TYPE.PLACEHOLDER`. Note that the return value is not the new table
        but rather *contains* the new table. The table can be accessed using the
        :attr:`~.PlaceholderGraphicFrame.table` property of the returned
        |PlaceholderGraphicFrame| object.

        This method is available on all slide placeholders (inherited from
        :class:`._BaseSlidePlaceholder`), not only on a specialized
        :class:`.TablePlaceholder`. This lets callers add a table to a generic
        "content", "body", or "object" placeholder without having to change the
        placeholder type first.
        """
        graphicFrame = self._new_placeholder_table(rows, cols)
        self._replace_placeholder_with(graphicFrame)
        return PlaceholderGraphicFrame(graphicFrame, self._parent)

    @property
    def _base_placeholder(self):
        """
        Return the layout placeholder this slide placeholder inherits from.
        Not to be confused with an instance of |BasePlaceholder|
        (necessarily).
        """
        layout, idx = self.part.slide_layout, self._element.ph_idx
        return layout.placeholders.get(idx=idx)

    def _new_chart_graphicFrame(self, rId, x, y, cx, cy):
        """Return a newly created `p:graphicFrame` element.

        The returned element has the specified position and size and contains the chart
        identified by `rId`.
        """
        id_, name = self.shape_id, self.name
        return CT_GraphicalObjectFrame.new_chart_graphicFrame(id_, name, rId, x, y, cx, cy)

    def _new_placeholder_pic(self, image_file, crop=True):
        """Return a new `p:pic` element depicting the image in *image_file*.

        Suitable for use as a placeholder. When *crop* is |True| the image is
        cropped to fill the placeholder and no `a:xfrm` element is emitted,
        allowing the picture's extents to be inherited from its layout
        placeholder. When *crop* is |False| an explicit `a:xfrm` is added with
        extents computed to fit the image inside the placeholder bounds
        preserving its aspect ratio, and with offsets that center the image
        within those bounds.
        """
        rId, desc, image_size = self._get_or_add_image(image_file)
        shape_id, name = self.shape_id, self.name
        pic = CT_Picture.new_ph_pic(shape_id, name, desc, rId)
        if crop:
            pic.crop_to_fit(image_size, (self.width, self.height))
        else:
            self._fit_pic_to_placeholder(pic, image_size)
        return pic

    def _new_placeholder_table(self, rows, cols):
        """Return a newly added `p:graphicFrame` element containing an empty table.

        The table has *rows* rows and *cols* columns, positioned at the location
        of this placeholder and having its same width. The table's height is
        determined by the number of rows.
        """
        shape_id, name, height = self.shape_id, self.name, Emu(rows * 370840)
        return CT_GraphicalObjectFrame.new_table_graphicFrame(
            shape_id, name, rows, cols, self.left, self.top, self.width, height
        )

    def _fit_pic_to_placeholder(self, pic, image_size):
        """Add an `a:xfrm` child to *pic* sized to fit *image_size* inside the placeholder.

        Preserves aspect ratio, and centered horizontally and vertically within
        the placeholder's bounds. No cropping is applied.
        """
        ph_x, ph_y = self.left, self.top
        ph_cx, ph_cy = self.width, self.height
        image_width, image_height = image_size
        # ---preserve aspect ratio, scale to fit inside the placeholder---
        if image_width * ph_cy > image_height * ph_cx:
            # ---image wider than placeholder aspect; width binds---
            fit_cx = ph_cx
            fit_cy = int(image_height * ph_cx / image_width)
        else:
            # ---image taller than (or equal to) placeholder aspect; height binds---
            fit_cy = ph_cy
            fit_cx = int(image_width * ph_cy / image_height)
        # ---center within placeholder---
        off_x = ph_x + (ph_cx - fit_cx) // 2
        off_y = ph_y + (ph_cy - fit_cy) // 2
        xfrm = pic.spPr.get_or_add_xfrm()
        xfrm.x, xfrm.y, xfrm.cx, xfrm.cy = off_x, off_y, fit_cx, fit_cy

    def _get_or_add_image(self, image_file):
        """Return an (rId, description, image_size) 3-tuple for *image_file*.

        Identifies the related image part containing *image_file* and describing
        the image.
        """
        image_part, rId = self.part.get_or_add_image_part(image_file)
        desc, image_size = image_part.desc, image_part._px_size
        return rId, desc, image_size

    def _replace_placeholder_with(self, element):
        """
        Substitute *element* for this placeholder element in the shapetree.
        This placeholder's `._element` attribute is set to |None| and its
        original element is free for garbage collection. Any attribute access
        (including a method call) on this placeholder after this call raises
        |AttributeError|.
        """
        element._nvXxPr.nvPr._insert_ph(self._element.ph)
        self._element.addprevious(element)
        self._element.getparent().remove(self._element)
        self._element = None


class BasePlaceholder(Shape):
    """
    NOTE: This class is deprecated and will be removed from a future release
    along with the properties *idx*, *orient*, *ph_type*, and *sz*. The *idx*
    property will be available via the .placeholder_format property. The
    others will be accessed directly from the oxml layer as they are only
    used for internal purposes.

    Base class for placeholder subclasses that differentiate the varying
    behaviors of placeholders on a master, layout, and slide.
    """

    @property
    def idx(self):
        """
        Integer placeholder 'idx' attribute, e.g. 0
        """
        return self._sp.ph_idx

    @property
    def orient(self):
        """
        Placeholder orientation, e.g. ST_Direction.HORZ
        """
        return self._sp.ph_orient

    @property
    def ph_type(self):
        """
        Placeholder type, e.g. PP_PLACEHOLDER.CENTER_TITLE
        """
        return self._sp.ph_type

    @property
    def sz(self):
        """
        Placeholder 'sz' attribute, e.g. ST_PlaceholderSize.FULL
        """
        return self._sp.ph_sz


class LayoutPlaceholder(_InheritsDimensions, Shape):
    """Placeholder shape on a slide layout.

    Provides differentiated behavior for slide layout placeholders, in particular, inheriting
    shape properties from the master placeholder having the same type, when a matching one exists.
    """

    element: CT_Shape  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def _base_placeholder(self):
        """
        Return the master placeholder this layout placeholder inherits from.
        """
        base_ph_type = {
            PP_PLACEHOLDER.BODY: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.CHART: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.BITMAP: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.CENTER_TITLE: PP_PLACEHOLDER.TITLE,
            PP_PLACEHOLDER.ORG_CHART: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.DATE: PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER: PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.MEDIA_CLIP: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.PICTURE: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.SLIDE_NUMBER: PP_PLACEHOLDER.SLIDE_NUMBER,
            PP_PLACEHOLDER.SUBTITLE: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.TABLE: PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.TITLE: PP_PLACEHOLDER.TITLE,
        }[self._element.ph_type]
        slide_master = self.part.slide_master
        return slide_master.placeholders.get(base_ph_type, None)


class MasterPlaceholder(BasePlaceholder):
    """Placeholder shape on a slide master."""

    element: CT_Shape  # pyright: ignore[reportIncompatibleMethodOverride]


class NotesSlidePlaceholder(_InheritsDimensions, Shape):
    """
    Placeholder shape on a notes slide. Inherits shape properties from the
    placeholder on the notes master that has the same type (e.g. 'body').
    """

    @property
    def _base_placeholder(self):
        """
        Return the notes master placeholder this notes slide placeholder
        inherits from, or |None| if no placeholder of the matching type is
        present.
        """
        notes_master = self.part.notes_master
        ph_type = self.element.ph_type
        return notes_master.placeholders.get(ph_type=ph_type)


class SlidePlaceholder(_BaseSlidePlaceholder):
    """Placeholder shape on a slide.

    Inherits shape properties from its corresponding slide layout placeholder.

    This is the class returned for generic "content", "body", and "object"
    placeholders, along with any other placeholder type not mapped to a
    specialized subclass. It inherits the full set of rich-content insertion
    methods from :class:`._BaseSlidePlaceholder`:

    * :meth:`~._BaseSlidePlaceholder.insert_chart` — replace the placeholder
      with a chart;
    * :meth:`~._BaseSlidePlaceholder.insert_picture` — replace the placeholder
      with a picture;
    * :meth:`~._BaseSlidePlaceholder.insert_table` — replace the placeholder
      with a table.

    This means any slide placeholder — for example a generic body or content
    placeholder — can be populated with a chart, picture, or table without
    needing a specialized :class:`.ChartPlaceholder`,
    :class:`.PicturePlaceholder`, or :class:`.TablePlaceholder`.
    """


class ChartPlaceholder(_BaseSlidePlaceholder):
    """Placeholder shape that can only accept a chart.

    The :meth:`~._BaseSlidePlaceholder.insert_chart` method is inherited from
    :class:`._BaseSlidePlaceholder` and is also available on generic
    :class:`.SlidePlaceholder` instances so that charts may be inserted into any
    slide placeholder, not only specialized chart placeholders.
    """


class PicturePlaceholder(_BaseSlidePlaceholder):
    """Placeholder shape that is specialized for holding a picture.

    The :meth:`~._BaseSlidePlaceholder.insert_picture` method is inherited from
    :class:`._BaseSlidePlaceholder` and is also available on generic
    :class:`.SlidePlaceholder` instances so that pictures may be inserted into
    any slide placeholder, not only specialized picture placeholders.
    """


class PlaceholderGraphicFrame(GraphicFrame):
    """
    Placeholder shape populated with a table, chart, or smart art.
    """

    @property
    def is_placeholder(self):
        """
        Boolean indicating whether this shape is a placeholder.
        Unconditionally |True| in this case.
        """
        return True


class PlaceholderPicture(_InheritsDimensions, Picture):
    """
    Placeholder shape populated with a picture.
    """

    @property
    def _base_placeholder(self):
        """
        Return the layout placeholder this picture placeholder inherits from.
        """
        layout, idx = self.part.slide_layout, self._element.ph_idx
        return layout.placeholders.get(idx=idx)


class TablePlaceholder(_BaseSlidePlaceholder):
    """Placeholder shape that is specialized for holding a table.

    The :meth:`~._BaseSlidePlaceholder.insert_table` method is inherited from
    :class:`._BaseSlidePlaceholder` and is also available on generic
    :class:`.SlidePlaceholder` instances so that tables may be inserted into any
    slide placeholder, not only specialized table placeholders.
    """
