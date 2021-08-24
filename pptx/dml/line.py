# encoding: utf-8

"""DrawingML objects related to line formatting."""

from __future__ import absolute_import, division, print_function, unicode_literals
from pptx.oxml.shapes.shared import (
    CT_LineJoinBevel, 
    CT_LineJoinMiterProperties,
    CT_LineJoinRound,
)

from ..enum.dml import MSO_FILL
from .fill import FillFormat
from ..util import Emu, lazyproperty


class LineFormat(object):
    """Provides access to line properties such as color, style, and width.

    A LineFormat object is typically accessed via the ``.line`` property of
    a shape such as |Shape| or |Picture|.
    """

    def __init__(self, parent):
        super(LineFormat, self).__init__()
        self._parent = parent

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
        return FillFormat.from_fill_parent(ln)

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

    @lazyproperty
    def join(self):
        """
        |JoinFormat| instance for this line providing access to join properties
        """
        ln = self._get_or_add_ln()
        return JoinFormat.from_join_parent(ln)

    @lazyproperty
    def head(self):
        """
        |LineEndFormat| instance for the head of this line providing acces
        to line end properties
        """
        ln = self._get_or_add_ln()
        return LineEndFormat(ln.get_or_add_headEnd())

    @lazyproperty
    def tail(self):
        """
        |LineEndFormat| instance for the tail of this line providing acces
        to line end properties
        """
        ln = self._get_or_add_ln()
        return LineEndFormat(ln.get_or_add_tailEnd())

    @property
    def cap(self):
        """
        Return value indicating cap style.  There are three valid options:
            - "rnd" (Round Line Cap)
            - "sq" (Square Line Cap)
            - "flat" (Flat Line Cap)
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.cap

    @cap.setter
    def cap(self, val):
        if val is None:
            return
        ln = self._get_or_add_ln()
        ln.cap = val


    @property
    def compound(self):
        """
        Return value indicating compound line style.  
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.cmpd

    @compound.setter
    def compound(self, val):
        if val is None:
            return
        ln = self._get_or_add_ln()
        ln.cmpd = val

    @property
    def align(self):
        """
        Return value indicationg the line alignment
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.algn

    @align.setter
    def align(self, val):
        if val is None:
            return
        ln = self._get_or_add_ln()
        ln.algn = val

    def _get_or_add_ln(self):
        """
        Return the ``<a:ln>`` element containing the line format properties
        in the XML.
        """
        return self._parent.get_or_add_ln()

    @property
    def _ln(self):
        return self._parent.ln



class LineStyle(object):
    """Provides access to line properties such as color, style, and width.

    While LineFormat is accessed from the parent, this is accessed directly
    and LineStyle is read only.  It is used in the Themes
    """

    def __init__(self, ln):
        super(LineStyle, self).__init__()
        self._ln = ln

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

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this line, providing access to fill
        properties such as foreground color.
        """
        ln = self._ln
        return FillFormat.from_fill_parent(ln)

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




class JoinFormat(object):
    """
    Provides access to the current line join properties object and provides
    methods to change the join type
    """

    def __init__(self, eg_join_properties_parent, join_obj):
        super(JoinFormat, self).__init__()
        self._xPr = eg_join_properties_parent
        self._join = join_obj

    @classmethod
    def from_join_parent(cls, eg_joinProperties_parent):
        """
        Return a |JoinFormat| instance initialized to the settings contained
        in *eg_joinProperties_parent*, which must be an emelent having 
        eg_lineJoinProperties in its child element sequence in the XML schemea.
        """
        join_elm = eg_joinProperties_parent.eg_lineJoinProperties
        join = _Join(join_elm)
        join_format = cls(eg_joinProperties_parent, join)
        return join_format

    def round(self):
        """
        Sets the join type to Round
        """
        roundJoin = self._xPr.get_or_change_to_round()
        self._join = _RoundJoin(roundJoin)

    def bevel(self):
        """
        Sets the join type to Bevel
        """
        bevelJoin = self._xPr.get_or_change_to_bevel()
        self._join = _BevelJoin(bevelJoin)

    def miter(self):
        """
        Sets the join type to Miter
        """
        miterJoin = self._xPr.get_or_change_to_miter()
        self._join = _MiterJoin(miterJoin)

    @property
    def limit(self):
        """
        The setting for the limit on a Miter Join. 
        """
        return self._join.limit

    @limit.setter
    def limit(self, val):
        join = self._xPr.get_or_change_to_miter()
        self._join = _MiterJoin(join)
        self._join.limit = val

    @property
    def type(self):
        """
        Returns the type value string
        """
        return self._join.type
    

class _Join(object):
    """
    Object factor for join object of class matching join element.
    """

    def __new__(cls, xJoin):
        if xJoin is None:
            join_cls = _NoneJoin
        elif isinstance(xJoin, CT_LineJoinRound):
            join_cls = _RoundJoin
        elif isinstance(xJoin, CT_LineJoinBevel):
            join_cls = _BevelJoin
        elif isinstance(xJoin, CT_LineJoinMiterProperties):
            join_cls = _MiterJoin
        else:
            join_cls = _Join
        return super(_Join, cls).__new__(join_cls)

    
class _NoneJoin(_Join):
    @property
    def type(self):
        return None

class _RoundJoin(_Join):
    @property
    def type(self):
        return "round"

class _BevelJoin(_Join):
    @property
    def type(self):
        return "bevel"

class _MiterJoin(_Join):
    def __init__(self, miterJoin):
        self._element = miterJoin

    @property
    def type(self):
        return "miter"
    
    @property
    def limit(self):
        """
        Property defining the angle of the miter.  Must be a positive percentage
        """
        return self._element.lim

    @limit.setter
    def limit(self, val):
        self._element.lim = val

class LineEndFormat(object):
    
    def __init__(self, end):
        self._end = end

    @property
    def width(self):
        return self._end.w

    @width.setter
    def width(self, val):
        if val is None:
            return
        self._end.w = val
    
    
    @property
    def length(self):
        return self._end.len

    @length.setter
    def length(self, val):
        if val is None:
            return
        self._end.len = val
    
    
    @property
    def type(self):
        return self._end.endType

    @type.setter
    def type(self, val):
        if val is None:
            return
        self._end.endType = val
    
    
