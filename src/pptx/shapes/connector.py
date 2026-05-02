"""Connector (line) shape and related objects.

A connector is a line shape having end-points that can be connected to other
objects (but not to other connectors). A connector can be straight, have
elbows, or can be curved.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Adjustment
from pptx.shapes.base import BaseShape
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.shapes.autoshape import CT_PresetGeometry2D


# -- Default `a:avLst` adjustment values for connector preset geometries. --
# -- Keys are the `prst` attribute value; values are a tuple of ``(name, default)`` pairs as --
# -- they appear in the connector's ``a:avLst``. Connector presets with no adjustments --
# -- (``line``, ``straightConnector1``, ``bentConnector2``, ``curvedConnector2``) are not listed; --
# -- any unknown `prst` is treated as having no adjustments. Values mirror the ISO/IEC 29500 --
# -- preset-shape definitions (`spec/ISO-IEC-29500-1/schemas/dml-geometries/...`). --
_CONNECTOR_DEFAULT_ADJUSTMENT_VALUES: dict[str, tuple[tuple[str, int], ...]] = {
    "bentConnector3": (("adj1", 50000),),
    "bentConnector4": (("adj1", 50000), ("adj2", 50000)),
    "bentConnector5": (("adj1", 50000), ("adj2", 50000), ("adj3", 50000)),
    "curvedConnector3": (("adj1", 50000),),
    "curvedConnector4": (("adj1", 50000), ("adj2", 50000)),
    "curvedConnector5": (("adj1", 50000), ("adj2", 50000), ("adj3", 50000)),
}


class ConnectorAdjustmentCollection:
    """Sequence of |Adjustment| instances for a connector shape.

    Each represents an available adjustment for a connector of its preset type. Supports
    ``len()`` and indexed access, e.g. ``connector.adjustments[0] = 0.25``. Indexing into the
    collection retrieves or assigns the *effective value* of the corresponding adjustment, a
    |float| nominally in the range 0.0 to 1.0.

    Straight connectors and two-segment connectors have no adjustments, so the collection will
    be empty (``len(adjustments) == 0``) for those preset types.
    """

    def __init__(self, prstGeom: CT_PresetGeometry2D | None):
        super(ConnectorAdjustmentCollection, self).__init__()
        self._prstGeom = prstGeom
        self._adjustments_ = self._initialized_adjustments(prstGeom)

    def __getitem__(self, idx: int) -> float:
        """Provides indexed access, (e.g. ``adjustments[0]``)."""
        return self._adjustments_[idx].effective_value

    def __setitem__(self, idx: int, value: float):
        """Provides item assignment via an indexed expression, e.g. ``adjustments[0] = 0.25``.

        Causes all adjustment values in collection to be written to the XML.
        """
        self._adjustments_[idx].effective_value = value
        self._rewrite_guides()

    def __len__(self):
        """Implement built-in function ``len()``."""
        return len(self._adjustments_)

    @property
    def _adjustments(self) -> tuple[Adjustment, ...]:
        """Sequence of |Adjustment| objects contained in collection."""
        return tuple(self._adjustments_)

    @staticmethod
    def _initialized_adjustments(prstGeom: CT_PresetGeometry2D | None) -> list[Adjustment]:
        """Return a list of |Adjustment| objects reflecting the current state of `prstGeom`."""
        if prstGeom is None:
            return []
        # -- connector prst values are not in MSO_AUTO_SHAPE_TYPE so avoid the descriptor --
        prst = prstGeom.get("prst")
        if prst is None:
            return []
        davs = _CONNECTOR_DEFAULT_ADJUSTMENT_VALUES.get(prst, ())
        adjustments = [Adjustment(name, def_val) for name, def_val in davs]
        adjustments_by_name = {adj.name: adj for adj in adjustments}
        for gd in prstGeom.gd_lst:
            adjustment = adjustments_by_name.get(gd.name)
            if adjustment is None:
                continue
            adjustment.actual = int(gd.fmla[4:])
        return adjustments

    def _rewrite_guides(self):
        """Write `a:gd` elements to the XML, one for each adjustment value in collection.

        Any existing `a:gd` children of `a:avLst` are overwritten.
        """
        assert self._prstGeom is not None  # -- non-empty collection has a prstGeom --
        guides = [(adj.name, adj.val) for adj in self._adjustments_]
        self._prstGeom.rewrite_guides(guides)


class Connector(BaseShape):
    """Connector (line) shape.

    A connector is a linear shape having end-points that can be connected to
    other objects (but not to other connectors). A connector can be straight,
    have elbows, or can be curved.
    """

    @lazyproperty
    def adjustments(self) -> ConnectorAdjustmentCollection:
        """|ConnectorAdjustmentCollection| instance for this connector.

        Connector adjustments are the normalized values (nominally 0.0 to 1.0) stored in the
        connector's `a:avLst` that control the position of bend points on elbow (bent) and
        curved connectors. Straight connectors and two-segment connectors have no adjustments;
        the collection is empty (``len == 0``) for those preset types.
        """
        return ConnectorAdjustmentCollection(self._element.spPr.prstGeom)

    def begin_connect(self, shape, cxn_pt_idx):
        """
        **EXPERIMENTAL** - *The current implementation only works properly
        with rectangular shapes, such as pictures and rectangles. Use with
        other shape types may cause unexpected visual alignment of the
        connected end-point and could lead to a load error if cxn_pt_idx
        exceeds the connection point count available on the connected shape.
        That said, a quick test should reveal what to expect when using this
        method with other shape types.*

        Connect the beginning of this connector to *shape* at the connection
        point specified by *cxn_pt_idx*. Each shape has zero or more
        connection points and they are identified by index, starting with 0.
        Generally, the first connection point of a shape is at the top center
        of its bounding box and numbering proceeds counter-clockwise from
        there. However this is only a convention and may vary, especially
        with non built-in shapes.
        """
        self._connect_begin_to(shape, cxn_pt_idx)
        self._move_begin_to_cxn(shape, cxn_pt_idx)

    @property
    def begin_x(self):
        """
        Return the X-position of the begin point of this connector, in
        English Metric Units (as a |Length| object).
        """
        cxnSp = self._element
        x, cx, flipH = cxnSp.x, cxnSp.cx, cxnSp.flipH
        begin_x = x + cx if flipH else x
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
        English Metric Units (as a |Length| object).
        """
        cxnSp = self._element
        y, cy, flipV = cxnSp.y, cxnSp.cy, cxnSp.flipV
        begin_y = y + cy if flipV else y
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

    def end_connect(self, shape, cxn_pt_idx):
        """
        **EXPERIMENTAL** - *The current implementation only works properly
        with rectangular shapes, such as pictures and rectangles. Use with
        other shape types may cause unexpected visual alignment of the
        connected end-point and could lead to a load error if cxn_pt_idx
        exceeds the connection point count available on the connected shape.
        That said, a quick test should reveal what to expect when using this
        method with other shape types.*

        Connect the ending of this connector to *shape* at the connection
        point specified by *cxn_pt_idx*.
        """
        self._connect_end_to(shape, cxn_pt_idx)
        self._move_end_to_cxn(shape, cxn_pt_idx)

    @property
    def end_x(self):
        """
        Return the X-position of the end point of this connector, in English
        Metric Units (as a |Length| object).
        """
        cxnSp = self._element
        x, cx, flipH = cxnSp.x, cxnSp.cx, cxnSp.flipH
        end_x = x if flipH else x + cx
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
        Metric Units (as a |Length| object).
        """
        cxnSp = self._element
        y, cy, flipV = cxnSp.y, cxnSp.cy, cxnSp.flipV
        end_y = y if flipV else y + cy
        return Emu(end_y)

    @end_y.setter
    def end_y(self, value):
        cxnSp = self._element
        y, cy, flipV, new_y = cxnSp.y, cxnSp.cy, cxnSp.flipV, int(value)

        if flipV:
            dy = abs(new_y - y)
            if new_y <= y:
                cxnSp.y = new_y
                cxnSp.cy = cy + dy
            elif dy <= cy:
                cxnSp.y = new_y
                cxnSp.cy = cy - dy
            else:
                cxnSp.flipV = False
                cxnSp.y = y + cy
                cxnSp.cy = dy - cy
        else:
            old_y = y + cy
            dy = abs(new_y - old_y)
            if new_y >= old_y:
                cxnSp.cy = cy + dy
            elif dy <= cy:
                cxnSp.cy = cy - dy
            else:
                cxnSp.flipV = True
                cxnSp.y = new_y
                cxnSp.cy = dy - cy

    def get_or_add_ln(self):
        """Helper method required by |LineFormat|."""
        return self._element.spPr.get_or_add_ln()

    @lazyproperty
    def line(self):
        """|LineFormat| instance for this connector.

        Provides access to line properties such as line color, width, and
        line style.
        """
        return LineFormat(self)

    @property
    def ln(self):
        """Helper method required by |LineFormat|.

        The ``<a:ln>`` element containing the line format properties such as
        line color and width. |None| if no `<a:ln>` element is present.
        """
        return self._element.spPr.ln

    @property
    def shape_type(self):
        """Member of `MSO_SHAPE_TYPE` identifying the type of this shape.

        Unconditionally `MSO_SHAPE_TYPE.LINE` for a `Connector` object.
        """
        return MSO_SHAPE_TYPE.LINE

    def _connect_begin_to(self, shape, cxn_pt_idx):
        """
        Add or update a stCxn element for this connector that connects its
        begin point to the connection point of *shape* specified by
        *cxn_pt_idx*.
        """
        cNvCxnSpPr = self._element.nvCxnSpPr.cNvCxnSpPr
        stCxn = cNvCxnSpPr.get_or_add_stCxn()
        stCxn.id = shape.shape_id
        stCxn.idx = cxn_pt_idx

    def _connect_end_to(self, shape, cxn_pt_idx):
        """
        Add or update an endCxn element for this connector that connects its
        end point to the connection point of *shape* specified by
        *cxn_pt_idx*.
        """
        cNvCxnSpPr = self._element.nvCxnSpPr.cNvCxnSpPr
        endCxn = cNvCxnSpPr.get_or_add_endCxn()
        endCxn.id = shape.shape_id
        endCxn.idx = cxn_pt_idx

    def _move_begin_to_cxn(self, shape, cxn_pt_idx):
        """
        Move the begin point of this connector to coordinates of the
        connection point of *shape* specified by *cxn_pt_idx*.
        """
        x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
        self.begin_x, self.begin_y = {
            0: (int(x + cx / 2), y),
            1: (x, int(y + cy / 2)),
            2: (int(x + cx / 2), y + cy),
            3: (x + cx, int(y + cy / 2)),
        }[cxn_pt_idx]

    def _move_end_to_cxn(self, shape, cxn_pt_idx):
        """
        Move the end point of this connector to the coordinates of the
        connection point of *shape* specified by *cxn_pt_idx*.
        """
        x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
        self.end_x, self.end_y = {
            0: (int(x + cx / 2), y),
            1: (x, int(y + cy / 2)),
            2: (int(x + cx / 2), y + cy),
            3: (x + cx, int(y + cy / 2)),
        }[cxn_pt_idx]
