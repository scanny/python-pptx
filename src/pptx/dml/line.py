"""DrawingML objects related to line formatting."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.fill import FillFormat
from pptx.enum.dml import (
    MSO_FILL,
    MSO_LINE_END_LENGTH,
    MSO_LINE_END_TYPE,
    MSO_LINE_END_WIDTH,
)
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.dml.line import CT_LineEndProperties
    from pptx.oxml.shapes.shared import CT_LineProperties


class LineFormat(object):
    """Provides access to line properties such as color, style, and width.

    A LineFormat object is typically accessed via the ``.line`` property of
    a shape such as |Shape| or |Picture|.
    """

    def __init__(self, parent):
        super(LineFormat, self).__init__()
        self._parent = parent

    @lazyproperty
    def begin_arrow(self) -> LineEndFormat:
        """|LineEndFormat| instance for the head end of this line.

        Provides access to the line-end decoration (arrowhead) at the
        *beginning* of the line (the `a:headEnd` child element). Read/write
        sub-properties are ``type``, ``width``, and ``length``.
        """
        return LineEndFormat(self, "headEnd")

    @lazyproperty
    def end_arrow(self) -> LineEndFormat:
        """|LineEndFormat| instance for the tail end of this line.

        Provides access to the line-end decoration (arrowhead) at the *end*
        of the line (the `a:tailEnd` child element). Read/write
        sub-properties are ``type``, ``width``, and ``length``.
        """
        return LineEndFormat(self, "tailEnd")

    @lazyproperty
    def color(self):
        """
        The |ColorFormat| instance that provides access to the color settings
        for this line. Essentially a shortcut for ``line.fill.fore_color``.
        As a side-effect, accessing this property causes the line fill type
        to be set to ``MSO_FILL.SOLID``. If this sounds risky for your use
        case, use ``line.fill.type`` to non-destructively discover the
        existing fill type.
        """
        if self.fill.type != MSO_FILL.SOLID:
            self.fill.solid()
        return self.fill.fore_color

    @property
    def dash_style(self):
        """Return value indicating line style.

        Returns a member of :ref:`MsoLineDashStyle` indicating line style, or
        |None| if no explicit value has been set. When no explicit value has
        been set, the line dash style is inherited from the style hierarchy.

        Assigning |None| removes any existing explicitly-defined dash style.
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.prstDash_val

    @dash_style.setter
    def dash_style(self, dash_style):
        if dash_style is None:
            ln = self._ln
            if ln is None:
                return
            ln._remove_prstDash()
            ln._remove_custDash()
            return
        ln = self._get_or_add_ln()
        ln.prstDash_val = dash_style

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this line, providing access to fill
        properties such as foreground color.
        """
        ln = self._get_or_add_ln()
        return FillFormat.from_fill_parent(ln, self._parent)

    @property
    def width(self):
        """
        The width of the line expressed as an integer number of :ref:`English
        Metric Units <EMU>`. The returned value is an instance of |Length|,
        a value class having properties such as `.inches`, `.cm`, and `.pt`
        for converting the value into convenient units.
        """
        ln = self._ln
        if ln is None:
            return Emu(0)
        return ln.w

    @width.setter
    def width(self, emu):
        if emu is None:
            emu = 0
        ln = self._get_or_add_ln()
        ln.w = emu

    def _get_or_add_ln(self) -> CT_LineProperties:
        """
        Return the ``<a:ln>`` element containing the line format properties
        in the XML.
        """
        return self._parent.get_or_add_ln()

    @property
    def _ln(self) -> CT_LineProperties | None:
        return self._parent.ln


class LineEndFormat(object):
    """Provides access to one end-decoration (arrowhead) on a line.

    An instance is obtained via the ``line.begin_arrow`` or
    ``line.end_arrow`` properties. Each instance exposes ``type``, ``width``,
    and ``length`` read/write properties.
    """

    _VALID_END_TAGS = ("headEnd", "tailEnd")

    def __init__(self, line: LineFormat, end_tag: str):
        super(LineEndFormat, self).__init__()
        if end_tag not in self._VALID_END_TAGS:
            raise ValueError("end_tag must be 'headEnd' or 'tailEnd', got %r" % end_tag)
        self._line = line
        self._end_tag = end_tag

    @property
    def type(self) -> MSO_LINE_END_TYPE | None:
        """Member of :ref:`MsoLineEndType` or |None| if not explicitly set."""
        end = self._end
        if end is None:
            return None
        return end.type

    @type.setter
    def type(self, value: MSO_LINE_END_TYPE | None):
        if value is None:
            end = self._end
            if end is None:
                return
            end.type = None
            self._prune_if_empty()
            return
        end = self._get_or_add_end()
        end.type = value

    @property
    def width(self) -> MSO_LINE_END_WIDTH | None:
        """Member of :ref:`MsoLineEndWidth` or |None| if not explicitly set."""
        end = self._end
        if end is None:
            return None
        return end.w

    @width.setter
    def width(self, value: MSO_LINE_END_WIDTH | None):
        if value is None:
            end = self._end
            if end is None:
                return
            end.w = None
            self._prune_if_empty()
            return
        end = self._get_or_add_end()
        end.w = value

    @property
    def length(self) -> MSO_LINE_END_LENGTH | None:
        """Member of :ref:`MsoLineEndLength` or |None| if not explicitly set."""
        end = self._end
        if end is None:
            return None
        return end.len

    @length.setter
    def length(self, value: MSO_LINE_END_LENGTH | None):
        if value is None:
            end = self._end
            if end is None:
                return
            end.len = None
            self._prune_if_empty()
            return
        end = self._get_or_add_end()
        end.len = value

    @property
    def _end(self) -> CT_LineEndProperties | None:
        """The `a:headEnd` or `a:tailEnd` element, or |None| if not present."""
        ln = self._line._ln  # pyright: ignore[reportPrivateUsage]
        if ln is None:
            return None
        return getattr(ln, self._end_tag)

    def _get_or_add_end(self) -> CT_LineEndProperties:
        """Return the `a:headEnd`/`a:tailEnd` element, creating it if needed."""
        ln = self._line._get_or_add_ln()  # pyright: ignore[reportPrivateUsage]
        return getattr(ln, "get_or_add_" + self._end_tag)()

    def _prune_if_empty(self):
        """Remove the end-properties element if no attribute is set."""
        end = self._end
        if end is None:
            return
        if end.type is None and end.w is None and end.len is None:
            ln = self._line._ln  # pyright: ignore[reportPrivateUsage]
            if ln is None:
                return
            getattr(ln, "_remove_" + self._end_tag)()
