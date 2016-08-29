# encoding: utf-8

"""
Connector (line) shape and related objects. A connector is a line shape
having end-points that can be connected to other objects (but not to other
connectors). A line can be straight, have elbows, or can be curved.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .base import BaseShape
from ..util import Emu


class Connector(BaseShape):
    """
    Connector (line) shape. A connector is a linear shape having end-points
    that can be connected to other objects (but not to other connectors).
    A line can be straight, have elbows, or can be curved.
    """
    @property
    def begin_x(self):
        """
        Return the X-position of the begin point of this connector, in
        English Metric Units (as an |Emu| object).
        """
        cxnSp = self._element
        x, cx, flipH = cxnSp.x, cxnSp.cx, cxnSp.flipH
        begin_x = x+cx if flipH else x
        return Emu(begin_x)

    @begin_x.setter
    def begin_x(self, value):
        cxnSp = self._element
        x, cx, flipH, new_x = cxnSp.x, cxnSp.cx, cxnSp.flipH, int(value)

        if flipH:
            old_x = x + cx
            dx = abs(new_x - old_x)
            if new_x >= old_x:
                cxnSp.cx = cx + dx
            elif dx <= cx:
                cxnSp.cx = cx - dx
            else:
                cxnSp.flipH = False
                cxnSp.x = new_x
                cxnSp.cx = dx - cx
        else:
            dx = abs(new_x - x)
            if new_x <= x:
                cxnSp.x = new_x
                cxnSp.cx = cx + dx
            elif dx <= cx:
                cxnSp.x = new_x
                cxnSp.cx = cx - dx
            else:
                cxnSp.flipH = True
                cxnSp.x = x + cx
                cxnSp.cx = dx - cx

    @property
    def begin_y(self):
        """
        Return the Y-position of the begin point of this connector, in
        English Metric Units (as an |Emu| object).
        """
        cxnSp = self._element
        y, cy, flipV = cxnSp.y, cxnSp.cy, cxnSp.flipV
        begin_y = y+cy if flipV else y
        return Emu(begin_y)

    @begin_y.setter
    def begin_y(self, value):
        cxnSp = self._element
        y, cy, flipV, new_y = cxnSp.y, cxnSp.cy, cxnSp.flipV, int(value)

        if flipV:
            old_y = y + cy
            dy = abs(new_y - old_y)
            if new_y >= old_y:
                cxnSp.cy = cy + dy
            elif dy <= cy:
                cxnSp.cy = cy - dy
            else:
                cxnSp.flipV = False
                cxnSp.y = new_y
                cxnSp.cy = dy - cy
        else:
            dy = abs(new_y - y)
            if new_y <= y:
                cxnSp.y = new_y
                cxnSp.cy = cy + dy
            elif dy <= cy:
                cxnSp.y = new_y
                cxnSp.cy = cy - dy
            else:
                cxnSp.flipV = True
                cxnSp.y = y + cy
                cxnSp.cy = dy - cy

    @property
    def end_x(self):
        """
        Return the X-position of the end point of this connector, in English
        Metric Units (as an |Emu| object).
        """
        cxnSp = self._element
        x, cx, flipH = cxnSp.x, cxnSp.cx, cxnSp.flipH
        end_x = x if flipH else x+cx
        return Emu(end_x)

    @end_x.setter
    def end_x(self, value):
        cxnSp = self._element
        x, cx, flipH, new_x = cxnSp.x, cxnSp.cx, cxnSp.flipH, int(value)

        if flipH:
            dx = abs(new_x - x)
            if new_x <= x:
                cxnSp.x = new_x
                cxnSp.cx = cx + dx
            elif dx <= cx:
                cxnSp.x = new_x
                cxnSp.cx = cx - dx
            else:
                cxnSp.flipH = False
                cxnSp.x = x + cx
                cxnSp.cx = dx - cx
        else:
            old_x = x + cx
            dx = abs(new_x - old_x)
            if new_x >= old_x:
                cxnSp.cx = cx + dx
            elif dx <= cx:
                cxnSp.cx = cx - dx
            else:
                cxnSp.flipH = True
                cxnSp.x = new_x
                cxnSp.cx = dx - cx

    @property
    def end_y(self):
        """
        Return the Y-position of the end point of this connector, in English
        Metric Units (as an |Emu| object).
        """
        cxnSp = self._element
        y, cy, flipV = cxnSp.y, cxnSp.cy, cxnSp.flipV
        end_y = y if flipV else y+cy
        return Emu(end_y)
