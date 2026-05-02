"""Value objects representing the path geometry of a shape.

The `Shape.path_geometry` accessor returns a |PathGeometry| which is a sequence of |Path|
objects (one per `a:path` element). Each path contains drawing operations
(|MoveTo|, |LineTo|, |CubicBezierTo|, |QuadBezierTo|, |ArcTo|, |Close|) whose coordinates
are |Length| instances expressed in shape-local EMU (English Metric Units).

The accessor supports two geometry kinds:

1. Custom geometry (freeform) — the `a:custGeom/a:pathLst` element tree is walked directly.
2. Preset geometry — a subset of preset shapes (those whose definition has no adjustment
   guides) is resolved against the shape's own bounding-box dimensions. Dynamic preset
   shapes (with `a:avLst` adjustments) are not currently supported and the accessor returns
   |None| for them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, Sequence, cast

from pptx.oxml.ns import qn
from pptx.oxml.shapes.preset_geometries import STATIC_PRESETS, resolve_token
from pptx.util import Emu

if TYPE_CHECKING:
    from lxml.etree import _Element  # pyright: ignore[reportPrivateUsage]

    from pptx.oxml.shapes.autoshape import CT_Path2DList, CT_Shape
    from pptx.util import Length


class _DrawingOp:
    """Base class for an immutable drawing-operation value object."""

    __slots__ = ("_points",)

    def __init__(self, points: tuple[tuple[Length, Length], ...]):
        self._points = points

    @property
    def points(self) -> tuple[tuple[Length, Length], ...]:
        """Sequence of (x, y) |Length| pairs defining this operation's vertices."""
        return self._points

    def __eq__(self, other: object) -> bool:
        if type(self) is not type(other):
            return NotImplemented
        return self._points == cast("_DrawingOp", other)._points

    def __hash__(self) -> int:
        return hash((type(self).__name__, self._points))

    def __repr__(self) -> str:
        return "%s(%r)" % (type(self).__name__, list(self._points))


class MoveTo(_DrawingOp):
    """Move-to operation. Sets the current pen location without drawing a segment."""

    def __init__(self, x: Length, y: Length):
        super().__init__(((x, y),))

    @property
    def x(self) -> Length:
        """X-coordinate of the move-to destination, in shape-local EMU."""
        return self._points[0][0]

    @property
    def y(self) -> Length:
        """Y-coordinate of the move-to destination, in shape-local EMU."""
        return self._points[0][1]


class LineTo(_DrawingOp):
    """Line-to operation. Draws a straight segment from the current pen location."""

    def __init__(self, x: Length, y: Length):
        super().__init__(((x, y),))

    @property
    def x(self) -> Length:
        """X-coordinate of the line-to endpoint, in shape-local EMU."""
        return self._points[0][0]

    @property
    def y(self) -> Length:
        """Y-coordinate of the line-to endpoint, in shape-local EMU."""
        return self._points[0][1]


class CubicBezierTo(_DrawingOp):
    """Cubic-Bezier-to operation with two control points and an endpoint."""

    def __init__(
        self,
        c1: tuple[Length, Length],
        c2: tuple[Length, Length],
        end: tuple[Length, Length],
    ):
        super().__init__((c1, c2, end))


class QuadBezierTo(_DrawingOp):
    """Quadratic-Bezier-to operation with a single control point and an endpoint."""

    def __init__(self, c1: tuple[Length, Length], end: tuple[Length, Length]):
        super().__init__((c1, end))


class ArcTo(_DrawingOp):
    """Arc-to operation.

    An arc-to is defined by ellipse width/height radii (`wR`, `hR`), a start angle, and a
    sweep angle in 1/60,000ths of a degree. The resulting drawing op has no enumerable points
    until the arc is flattened.
    """

    __slots__ = ("_wR", "_hR", "_stAng", "_swAng")

    def __init__(self, wR: Length, hR: Length, stAng: int, swAng: int):
        super().__init__(())
        self._wR = wR
        self._hR = hR
        self._stAng = stAng
        self._swAng = swAng

    @property
    def wR(self) -> Length:
        """Ellipse x-radius (width radius) for the arc, in shape-local EMU."""
        return self._wR

    @property
    def hR(self) -> Length:
        """Ellipse y-radius (height radius) for the arc, in shape-local EMU."""
        return self._hR

    @property
    def stAng(self) -> int:
        """Start angle of the arc, in 1/60,000ths of a degree."""
        return self._stAng

    @property
    def swAng(self) -> int:
        """Sweep angle of the arc, in 1/60,000ths of a degree."""
        return self._swAng

    def __eq__(self, other: object) -> bool:
        if type(self) is not type(other):
            return NotImplemented
        o = cast("ArcTo", other)
        return (
            self._wR == o._wR
            and self._hR == o._hR
            and self._stAng == o._stAng
            and self._swAng == o._swAng
        )

    def __hash__(self) -> int:
        return hash(("ArcTo", self._wR, self._hR, self._stAng, self._swAng))


class Close(_DrawingOp):
    """Close operation. Closes the current contour by drawing a segment to its start."""

    def __init__(self):
        super().__init__(())


class Path(Sequence["_DrawingOp"]):
    """A single contour in a |PathGeometry|.

    Iterable sequence of drawing-operation objects. A path's `width` and `height` are the
    dimensions of the path's coordinate system — for preset shapes resolved by this library
    they are always equal to the shape's bounding-box dimensions.
    """

    def __init__(self, width: Length, height: Length, ops: tuple[_DrawingOp, ...]):
        self._width = width
        self._height = height
        self._ops = ops

    def __getitem__(  # pyright: ignore[reportIncompatibleMethodOverride]
        self, idx: int
    ) -> _DrawingOp:
        return self._ops[idx]

    def __iter__(self) -> Iterator[_DrawingOp]:
        return iter(self._ops)

    def __len__(self) -> int:
        return len(self._ops)

    @property
    def width(self) -> Length:
        """Width of the path's coordinate system, in EMU."""
        return self._width

    @property
    def height(self) -> Length:
        """Height of the path's coordinate system, in EMU."""
        return self._height


class PathGeometry(Sequence[Path]):
    """Read-only sequence of |Path| objects describing a shape's rendered outline.

    For preset shapes (`a:prstGeom`), the path list is derived from the ECMA-376 preset
    definition after resolving the shape's width and height. Only shapes whose geometry does
    not depend on adjustment values are currently supported — for dynamic shapes, the
    `Shape.path_geometry` accessor returns |None|.

    For custom-geometry (freeform) shapes, each `a:path` element becomes one |Path|.
    """

    def __init__(self, paths: tuple[Path, ...]):
        self._paths = paths

    def __getitem__(self, idx: int) -> Path:  # pyright: ignore[reportIncompatibleMethodOverride]
        return self._paths[idx]

    def __iter__(self) -> Iterator[Path]:
        return iter(self._paths)

    def __len__(self) -> int:
        return len(self._paths)

    @classmethod
    def from_shape(cls, sp: CT_Shape) -> PathGeometry | None:
        """Return a |PathGeometry| for `sp` or |None| if geometry cannot be resolved.

        `None` is returned when the shape is a placeholder with no geometry, or when its
        preset definition requires adjustment-value evaluation (dynamic preset).
        """
        # -- custom geometry: walk the a:pathLst element tree directly --
        if sp.has_custom_geometry:
            custGeom = sp.spPr.custGeom
            assert custGeom is not None
            pathLst = custGeom.pathLst
            if pathLst is None:
                return cls(())
            return _build_from_pathLst(pathLst)

        # -- preset geometry: resolve against shape dimensions --
        prstGeom = sp.prstGeom
        # -- `CT_Shape.prstGeom` is typed as non-optional but may actually be None
        # -- at runtime (e.g. placeholder shapes without `a:prstGeom`).
        if prstGeom is None:  # pyright: ignore[reportUnnecessaryComparison]
            return None

        prst = prstGeom.prst
        prst_xml = prst.xml_value if hasattr(prst, "xml_value") else str(prst)
        if prst_xml not in STATIC_PRESETS:
            return None

        # -- a present-but-non-empty avLst means the shape may have non-default
        # -- adjustment values set, so we bail out and treat this as a dynamic preset. --
        if prstGeom.gd_lst:
            return None

        cx, cy = sp.cx, sp.cy
        # -- `BaseShapeElement.cx`/`cy` are typed as `Length` but can return None
        # -- when there is no `a:xfrm/a:ext` descendant. --
        if cx is None or cy is None:  # pyright: ignore[reportUnnecessaryComparison]
            return None
        return _build_from_preset(prst_xml, cx, cy)


# -- internal helpers --------------------------------------------------------------------

_MOVE_TO = qn("a:moveTo")
_LINE_TO = qn("a:lnTo")
_CUBIC_BEZ_TO = qn("a:cubicBezTo")
_QUAD_BEZ_TO = qn("a:quadBezTo")
_ARC_TO = qn("a:arcTo")
_CLOSE = qn("a:close")
_PT = qn("a:pt")


def _build_from_pathLst(pathLst: CT_Path2DList) -> PathGeometry:
    """Return a |PathGeometry| built from the `a:path` children of `pathLst`."""
    paths: list[Path] = []
    for path_el in pathLst.findall(qn("a:path")):
        paths.append(_build_custom_path(path_el))
    return PathGeometry(tuple(paths))


def _build_custom_path(path_el: _Element) -> Path:
    """Return a |Path| from the drawing-op children of `a:path` element `path_el`."""
    ops: list[_DrawingOp] = []
    for child in path_el.iterchildren():
        tag = child.tag
        if tag == _MOVE_TO:
            pt = _read_pt(child)
            if pt is not None:
                ops.append(MoveTo(*pt))
        elif tag == _LINE_TO:
            pt = _read_pt(child)
            if pt is not None:
                ops.append(LineTo(*pt))
        elif tag == _CUBIC_BEZ_TO:
            pts = _read_pts(child)
            if len(pts) == 3:
                ops.append(CubicBezierTo(pts[0], pts[1], pts[2]))
        elif tag == _QUAD_BEZ_TO:
            pts = _read_pts(child)
            if len(pts) == 2:
                ops.append(QuadBezierTo(pts[0], pts[1]))
        elif tag == _ARC_TO:
            wR = Emu(int(child.get("wR", "0")))
            hR = Emu(int(child.get("hR", "0")))
            stAng = int(child.get("stAng", "0"))
            swAng = int(child.get("swAng", "0"))
            ops.append(ArcTo(wR, hR, stAng, swAng))
        elif tag == _CLOSE:
            ops.append(Close())

    w_attr = path_el.get("w")
    h_attr = path_el.get("h")
    w = Emu(int(w_attr)) if w_attr is not None else Emu(0)
    h = Emu(int(h_attr)) if h_attr is not None else Emu(0)
    return Path(w, h, tuple(ops))


def _read_pt(op_el: _Element) -> tuple[Length, Length] | None:
    """Return the (x, y) pair from the `a:pt` child of `op_el`, or |None| if absent."""
    pt = op_el.find(_PT)
    if pt is None:
        return None
    return Emu(int(pt.get("x", "0"))), Emu(int(pt.get("y", "0")))


def _read_pts(op_el: _Element) -> list[tuple[Length, Length]]:
    """Return the list of (x, y) pairs for all `a:pt` children of `op_el`."""
    return [(Emu(int(pt.get("x", "0"))), Emu(int(pt.get("y", "0")))) for pt in op_el.findall(_PT)]


def _build_from_preset(prst: str, cx: Length, cy: Length) -> PathGeometry:
    """Return |PathGeometry| for static preset `prst` given shape width/height `cx`,`cy`."""
    preset_paths = STATIC_PRESETS[prst]
    paths: list[Path] = []
    for preset_path in preset_paths:
        if preset_path.w is None and preset_path.h is None:
            x_scale, y_scale = int(cx), int(cy)
        else:
            assert preset_path.w is not None
            assert preset_path.h is not None
            x_scale, y_scale = preset_path.w, preset_path.h
        path_w, path_h = Emu(int(cx)), Emu(int(cy))
        ops: list[_DrawingOp] = []
        for preset_op in preset_path.ops:
            resolved_pts = [
                (
                    _resolve(x_tok, cx, cy, x_scale, y_scale, axis="x"),
                    _resolve(y_tok, cx, cy, x_scale, y_scale, axis="y"),
                )
                for x_tok, y_tok in preset_op.points
            ]
            ops.append(_op_from_preset(preset_op.name, resolved_pts))
        paths.append(Path(path_w, path_h, tuple(ops)))
    return PathGeometry(tuple(paths))


def _resolve(token: str, cx: Length, cy: Length, scale_w: int, scale_h: int, axis: str) -> Length:
    """Resolve preset-path coord `token` to a shape-local EMU |Length|.

    Named tokens (e.g. `"hc"`) use shape-bbox reference values (`cx`, `cy`). Integer literals
    are interpreted as a fraction `literal / scale` of the axis' shape-bbox dimension.
    """
    try:
        int_literal = int(token)
    except ValueError:
        raw = resolve_token(token, int(cx), int(cy))
        return Emu(int(raw))
    if axis == "x":
        return Emu(int(cx) * int_literal // scale_w)
    return Emu(int(cy) * int_literal // scale_h)


def _op_from_preset(op_name: str, pts: list[tuple[Length, Length]]) -> _DrawingOp:
    """Build a drawing-op value-object from a (op_name, resolved-points) pair."""
    if op_name == "moveTo":
        return MoveTo(pts[0][0], pts[0][1])
    if op_name == "lnTo":
        return LineTo(pts[0][0], pts[0][1])
    if op_name == "cubicBezTo":
        return CubicBezierTo(pts[0], pts[1], pts[2])
    if op_name == "quadBezTo":
        return QuadBezierTo(pts[0], pts[1])
    if op_name == "close":
        return Close()
    raise ValueError("unexpected preset drawing op %r" % op_name)
