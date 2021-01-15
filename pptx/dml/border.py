# encoding: utf-8

"""DrawingML objects related to border/line formatting."""

from __future__ import absolute_import, division, print_function, unicode_literals

from ..enum.dml import MSO_FILL
from .fill import FillFormat
from .line import LineFormat
from ..util import Emu, lazyproperty


class BorderFormat(LineFormat):
    """Provides access to line properties such as color, style, and width.
    This simply inherits the standard |LineFormat| object but abstracts the
    _ln property and _get_or_add_ln method to allow for a side designation

    A |BorderFormat| object is typically accessed via the ``.border_xxx`` property of
    a Table Cell.
    """

    def __init__(self, parent, side):
        super(BorderFormat, self).__init__(parent)
        self.side = side

    @property
    def _ln(self):
        if self.side == "left":
            return self._parent.border_left
        elif self.side == "right":
            return self._parent.border_right
        elif self.side == "top":
            return self._parent.border_top
        elif self.side == "bottom":
            return self._parent.border_bottom
        elif self.side == "bl_tr":
            return self._parent.border_bl_tr
        elif self.side == "tl_br":
            return self._parent.border_tl_br
        else:
            return None


    def _get_or_add_ln(self):
        if self.side == "left":
            return self._parent.get_or_add_borL()
        elif self.side == "right":
            return self._parent.get_or_add_borR()
        elif self.side == "top":
            return self._parent.get_or_add_borT()
        elif self.side == "bottom":
            return self._parent.get_or_add_borB()
        elif self.side == "bl_tr":
            return self._parent.get_or_add_borBlTr()
        elif self.side == "tl_br":
            return self._parent.get_or_add_borTlBr()
        else:
            return None
