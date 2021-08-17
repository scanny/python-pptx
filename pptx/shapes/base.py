# encoding: utf-8

"""Base shape-related objects such as BaseShape."""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.action import ActionSetting
from pptx.dml.effect import ShadowFormat
from pptx.shared import ElementProxy
from pptx.util import lazyproperty
from pptx.dml.color import ColorFormat


class BaseShape(object):
    """Base class for shape objects.

    Subclasses include |Shape|, |Picture|, and |GraphicFrame|.
    """

    def __init__(self, shape_elm, parent):
        super(BaseShape, self).__init__()
        self._element = shape_elm
        self._parent = parent

    def __eq__(self, other):
        """|True| if this shape object proxies the same element as *other*.

        Equality for proxy objects is defined as referring to the same XML
        element, whether or not they are the same proxy object instance.
        """
        if not isinstance(other, BaseShape):
            return False
        return self._element is other._element

    def __ne__(self, other):
        if not isinstance(other, BaseShape):
            return True
        return self._element is not other._element

    @lazyproperty
    def click_action(self):
        """|ActionSetting| instance providing access to click behaviors.

        Click behaviors are hyperlink-like behaviors including jumping to
        a hyperlink (web page) or to another slide in the presentation. The
        click action is that defined on the overall shape, not a run of text
        within the shape. An |ActionSetting| object is always returned, even
        when no click behavior is defined on the shape.
        """
        cNvPr = self._element._nvXxPr.cNvPr
        return ActionSetting(cNvPr, self)

    @property
    def element(self):
        """`lxml` element for this shape, e.g. a CT_Shape instance.

        Note that manipulating this element improperly can produce an invalid
        presentation file. Make sure you know what you're doing if you use
        this to change the underlying XML.
        """
        return self._element

    @property
    def has_chart(self):
        """
        |True| if this shape is a graphic frame containing a chart object.
        |False| otherwise. When |True|, the chart object can be accessed
        using the ``.chart`` property.
        """
        # This implementation is unconditionally False, the True version is
        # on GraphicFrame subclass.
        return False

    @property
    def has_table(self):
        """
        |True| if this shape is a graphic frame containing a table object.
        |False| otherwise. When |True|, the table object can be accessed
        using the ``.table`` property.
        """
        # This implementation is unconditionally False, the True version is
        # on GraphicFrame subclass.
        return False

    @property
    def has_text_frame(self):
        """
        |True| if this shape can contain text.
        """
        # overridden on Shape to return True. Only <p:sp> has text frame
        return False

    @property
    def height(self):
        """
        Read/write. Integer distance between top and bottom extents of shape
        in EMUs
        """
        return self._element.cy

    @height.setter
    def height(self, value):
        self._element.cy = value

    @property
    def is_placeholder(self):
        """
        True if this shape is a placeholder. A shape is a placeholder if it
        has a <p:ph> element.
        """
        return self._element.has_ph_elm

    @property
    def left(self):
        """
        Read/write. Integer distance of the left edge of this shape from the
        left edge of the slide, in English Metric Units (EMU)
        """
        return self._element.x

    @left.setter
    def left(self, value):
        self._element.x = value

    @property
    def name(self):
        """
        Name of this shape, e.g. 'Picture 7'
        """
        return self._element.shape_name

    @name.setter
    def name(self, value):
        self._element._nvXxPr.cNvPr.name = value

    
    @property
    def hidden(self):
        """
        Read/write visiblity status of this shape.  Defaults to False
        """
        return self._element.hidden

    @hidden.setter
    def hidden(self, value):
        self._element._nvXxPr.cNvPr.hidden = value

    
    @property
    def part(self):
        """The package part containing this shape.

        A |BaseSlidePart| subclass in this case. Access to a slide part
        should only be required if you are extending the behavior of |pp| API
        objects.
        """
        return self._parent.part

    @property
    def placeholder_format(self):
        """
        A |_PlaceholderFormat| object providing access to
        placeholder-specific properties such as placeholder type. Raises
        |ValueError| on access if the shape is not a placeholder.
        """
        if not self.is_placeholder:
            raise ValueError("shape is not a placeholder")
        return _PlaceholderFormat(self._element.ph)

    @property
    def rotation(self):
        """
        Read/write float. Degrees of clockwise rotation. Negative values can
        be assigned to indicate counter-clockwise rotation, e.g. assigning
        -45.0 will change setting to 315.0.
        """
        return self._element.rot

    @rotation.setter
    def rotation(self, value):
        self._element.rot = value

    @lazyproperty
    def shadow(self):
        """|ShadowFormat| object providing access to shadow for this shape.

        A |ShadowFormat| object is always returned, even when no shadow is
        explicitly defined on this shape (i.e. it inherits its shadow
        behavior).
        """
        return ShadowFormat(self._element.spPr)

    @property
    def shape_id(self):
        """Read-only positive integer identifying this shape.

        The id of a shape is unique among all shapes on a slide.
        """
        return self._element.shape_id

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, like
        ``MSO_SHAPE_TYPE.CHART``. Must be implemented by subclasses.
        """
        # # This one returns |None| unconditionally to account for shapes
        # # that haven't been implemented yet, like group shape and chart.
        # # Once those are done this should raise |NotImplementedError|.
        # msg = 'shape_type property must be implemented by subclasses'
        # raise NotImplementedError(msg)
        return None

    @property
    def top(self):
        """
        Read/write. Integer distance of the top edge of this shape from the
        top edge of the slide, in English Metric Units (EMU)
        """
        return self._element.y

    @top.setter
    def top(self, value):
        self._element.y = value

    @property
    def width(self):
        """
        Read/write. Integer distance between left and right extents of shape
        in EMUs
        """
        return self._element.cx

    @width.setter
    def width(self, value):
        self._element.cx = value

    @lazyproperty
    def style(self):
        """
        Returns a |ShapeStyle| object or None if not available
        """
        obj = self._element.style
        if obj is None:
            return None
        return ShapeStyle(obj, self._element)
        
    def remove_style(self):
        """
        Removes any Style formatting
        """
        self._element.remove_style()


    @property
    def flip_h(self):
        """
        Read/Write.  Boolean for flipping a shape horizontally.
        """
        return self._element.flipH

    @flip_h.setter
    def flip_h(self, value):
        self._element.flipH = value


    @property
    def flip_v(self):
        """
        Read/Write.  Boolean for flipping a shape horizontally.
        """
        return self._element.flipV

    @flip_v.setter
    def flip_v(self, value):
        self._element.flipV = value

    @property
    def custom_geometry(self):
        if not self._element.has_custom_geometry:
            return None
        else:
            return CustomGeometry(self._element.custGeom)

    def add_custom_geometry(self):
        if self._element.has_custom_geometry:
            return CustomGeometry(self._element.custGeom)
        return CustomGeometry(self._element.spPr._add_custGeom())
    
    @property
    def preset_geometry(self):
        if self._element.has_custom_geometry:
            return None
        return PresetGeometry(self._element.spPr.prstGeom)

class _PlaceholderFormat(ElementProxy):
    """
    Accessed via the :attr:`~.BaseShape.placeholder_format` property of
    a placeholder shape, provides properties specific to placeholders, such
    as the placeholder type.
    """

    @property
    def element(self):
        """
        The `p:ph` element proxied by this object.
        """
        return super(_PlaceholderFormat, self).element

    @property
    def idx(self):
        """
        Integer placeholder 'idx' attribute.
        """
        return self._element.idx

    @property
    def type(self):
        """
        Placeholder type, a member of the :ref:`PpPlaceholderType`
        enumeration, e.g. PP_PLACEHOLDER.CHART
        """
        return self._element.type



class ShapeStyle(object):
    """
    The `p:style` object that handles the references for the shape
    to the theme's styles
    """

    def __init__(self, style_elm, parent):
        super(ShapeStyle, self).__init__()
        self._element = style_elm
        self._parent = parent

    @property
    def line_ref(self):
        """
        A Line Reference
        """
        return StyleMatrixReference(self._element.lnRef)

    @property
    def fill_ref(self):
        """
        A Fill Reference
        """
        return StyleMatrixReference(self._element.fillRef)

    @property
    def effect_ref(self):
        """
        An Effect Reference
        """
        return StyleMatrixReference(self._element.effectRef)

    @property
    def font_ref(self):
        """
        A Font Reference
        """
        return StyleMatrixReference(self._element.fontRef)




class StyleMatrixReference(object):
    """
    Object that is used by |ShapeStyle| to handle the four different references
    """

    def __init__(self, matrix_ref):
        super(StyleMatrixReference, self).__init__()
        self._reference = matrix_ref

    @lazyproperty
    def color_reference(self):
        """Return |ColorFormat| object matching reference color."""
        return ColorFormat.from_colorchoice_parent(self._reference)

    @property
    def idx(self):
        """ Read/Write the IDX of the referenced theme. """
        return self._reference.idx

    @idx.setter
    def idx(self, value):
        self._reference.idx = value


class CustomGeometry(ElementProxy):
    """
    Class that proxies the ``<a:custGeom>`` tag used for
    custom shape geometry.  
    """

    @property
    def adjust_values(self):
        return GeometryGuideList(self._element.get_or_add_avLst())

    @property
    def shape_guides(self):
        return GeometryGuideList(self._element.get_or_add_gdLst())

    # @property
    # def shape_handles(self):
    #     return self._element.ahLst

    @property
    def connection_sites(self):
        return ConnectionSiteList(self._element.get_or_add_cxnLst())

    @property
    def rectangle(self):
        return GeometricRectangle(self._element.get_or_add_rect())

    @property
    def paths(self):
        return PathList(self._element.get_or_add_pathLst())

class PresetGeometry(ElementProxy):
    """
    Class that proxies the ``<a:prstGeom>`` tag used for
    preset shape geometry.  
    """

    @property
    def adjust_values(self):
        return GeometryGuideList(self._element.get_or_add_avLst())


class GeometryGuideList(ElementProxy):
    """List of Shape Guides used by |CustomGeometry|.
    """

    def __init__(self, guide_list):
        super(GeometryGuideList, self).__init__(guide_list)

    @property
    def guide_list(self):
        gd = self._element.gd_lst
        if gd is None:
            return []
        return [GeometryGuide(guide) for guide in gd]

    def add_guide(self):
        return GeometryGuide(self._element._add_gd())


class GeometryGuide(ElementProxy):
    """ Individual Geometry Guides for elements `a:gd`
    """
    def __init__(self, guide):
        super(GeometryGuide, self).__init__(guide)
    
    @property
    def name(self):
        return self._element.name

    @name.setter
    def name(self, value):
        self._element.name = value

    @property
    def formula(self):
        return self._element.fmla

    @formula.setter
    def formula(self, value):
        self._element.fmla = value


class ConnectionSiteList(ElementProxy):
    """List of Connection Sites used by |CustomGeometry|.
    """

    def __init__(self, cxn_site_list):
        super(ConnectionSiteList, self).__init__(cxn_site_list)

    @property
    def sites_list(self):
        sites = self._element.cxn_lst
        if sites is None:
            return []
        return [ConnectionSite(site) for site in sites]

    def add_site(self):
        return ConnectionSite(self._element._add_cxn())


class ConnectionSite(ElementProxy):
    """ Connection Sites represented by `a:cxn`
    """
    def __init__(self, site):
        super(ConnectionSite, self).__init__(site)
    
    @property
    def angle(self):
        return self._element.ang

    @angle.setter
    def angle(self, value):
        self._element.ang = value

    @property
    def position(self):
        pos = self._element.pos
        return (pos.x, pos.y)

    @position.setter
    def position(self, coords):
        pos = self._element.get_or_add_pos()
        pos.x = coords[0]
        pos.y = coords[1]
        

class GeometricRectangle(ElementProxy):
    """ Object to define a rectangle for a textbox used by |CustomGeometry|
    """

    def __init__(self, rectangle):
        super(GeometricRectangle, self).__init__(rectangle)
    
    @property
    def left(self):
        return self._element.l
    
    @left.setter
    def left(self, value):
        self._element.l = value

    @property
    def right(self):
        return self._element.r
    
    @right.setter
    def right(self, value):
        self._element.r = value

    @property
    def top(self):
        return self._element.t
    
    @top.setter
    def top(self, value):
        self._element.t = value

    @property
    def bottom(self):
        return self._element.b
    
    @bottom.setter
    def bottom(self, value):
        self._element.b = value


class PathList(ElementProxy):
    """List of Shape Paths used by |CustomGeometry|.
    """

    def __init__(self, path_list):
        super(PathList, self).__init__(path_list)

    @property
    def path_list(self):
        paths = self._element.path_lst
        if paths is None:
            return []
        return [ShapePath(path) for path in paths]

    def add_path(self):
        return ShapePath(self._element._add_path())


class ShapePath(ElementProxy):
    """ Individual Shape Paths contained in `a:path`
    """
    def __init__(self, path):
        super(ShapePath, self).__init__(path)
    
    @property
    def width(self):
        return self._element.w

    @width.setter
    def width(self, value):
        self._element.w = value

    @property
    def height(self):
        return self._element.h

    @height.setter
    def height(self, value):
        self._element.h = value

    @property
    def fill_mode(self):
        return self._element.fill

    @fill_mode.setter
    def fill_mode(self, value):
        self._element.fill = value

    @property
    def stroke(self):
        return self._element.stroke

    @stroke.setter
    def stroke(self, value):
        self._element.stroke = value

    @property
    def extrusion(self):
        return self._element.extrusionOk

    @extrusion.setter
    def extrusion(self, value):
        self._element.extrusionOk = value

    def add_close(self):
        return self._element.add_close()

    def add_line_to(self, x, y):
        return self._element.add_lnTo(x,y)

    def add_move_to(self, x, y):
        return self._element.add_moveTo(x, y)

    def add_arc_to(self, width_radius, height_radius, start_angle, swing_angle):
        return self._element.add_arcTo(width_radius, height_radius, start_angle, swing_angle)
    
    def add_quad_bez_to(self, point1, point2):
        return self._element.add_quadBezTo(point1, point2)

    def add_cubic_bez_to(self, point1, point2, point3):
        return self._element.add_cubicBezTo(point1, point2, point3)

    @property
    def path_sequence(self):
        return list(self._element)
            



