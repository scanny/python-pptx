"""Static preset-shape path definitions from ECMA-376 Part 1, Annex D.

Only the subset of preset shapes whose geometry does not depend on any adjustment values
(`a:avLst`) or computed guides (`a:gdLst`) is encoded here. Dynamic shapes (those that
require evaluating the DrawingML formula language against an `a:avLst`) are intentionally
excluded.

Callers should consult `STATIC_PRESETS` keyed by the `prst` string used on `a:prstGeom`
(e.g. `"rect"`, `"flowChartProcess"`). Each entry is a list of `PresetPath` records.

The data is derived from the presetShapeDefinitions.xml reference in
`spec/ISO-IEC-29500-1/schemas/dml-geometries/`.
"""

from __future__ import annotations

from typing import NamedTuple


class PresetOp(NamedTuple):
    """A single drawing operation on a preset path.

    `name` is one of `"moveTo"`, `"lnTo"`, `"close"`, `"quadBezTo"`, `"cubicBezTo"`.
    `points` is the sequence of `(x_token, y_token)` token pairs for the operation
    (empty for `"close"`).
    """

    name: str
    points: tuple[tuple[str, str], ...]


class PresetPath(NamedTuple):
    """A single `a:path` record in a static preset definition.

    `w` and `h` are the path's own coordinate-system dimensions (both |None| when the
    path inherits the shape's bounding-box coordinate system).
    """

    w: int | None
    h: int | None
    ops: tuple[PresetOp, ...]


# -- Named coordinate constants referenced by the static preset path data. The numeric
# -- value is computed from the shape's bounding-box width `w` and height `h`.
NAMED_X_TOKENS = {"l", "r", "hc", "wd2"}
NAMED_Y_TOKENS = {"t", "b", "vc", "hd2"}


def _p(w: int | None, h: int | None, *ops: tuple[str, tuple[tuple[str, str], ...]]) -> PresetPath:
    """Return a |PresetPath| built from a terser literal tuple form used below."""
    return PresetPath(w, h, tuple(PresetOp(name, points) for name, points in ops))


STATIC_PRESETS: dict[str, tuple[PresetPath, ...]] = {
    "actionButtonBlank": (
        _p(
            None,
            None,
            ("moveTo", (("l", "t"),)),
            ("lnTo", (("r", "t"),)),
            ("lnTo", (("r", "b"),)),
            ("lnTo", (("l", "b"),)),
            ("close", ()),
        ),
    ),
    "chartPlus": (
        _p(
            10,
            10,
            ("moveTo", (("5", "0"),)),
            ("lnTo", (("5", "10"),)),
            ("moveTo", (("0", "5"),)),
            ("lnTo", (("10", "5"),)),
        ),
        _p(
            10,
            10,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("0", "10"),)),
            ("lnTo", (("10", "10"),)),
            ("lnTo", (("10", "0"),)),
            ("close", ()),
        ),
    ),
    "chartStar": (
        _p(
            10,
            10,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("10", "10"),)),
            ("moveTo", (("0", "10"),)),
            ("lnTo", (("10", "0"),)),
            ("moveTo", (("5", "0"),)),
            ("lnTo", (("5", "10"),)),
        ),
        _p(
            10,
            10,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("0", "10"),)),
            ("lnTo", (("10", "10"),)),
            ("lnTo", (("10", "0"),)),
            ("close", ()),
        ),
    ),
    "chartX": (
        _p(
            10,
            10,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("10", "10"),)),
            ("moveTo", (("0", "10"),)),
            ("lnTo", (("10", "0"),)),
        ),
        _p(
            10,
            10,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("0", "10"),)),
            ("lnTo", (("10", "10"),)),
            ("lnTo", (("10", "0"),)),
            ("close", ()),
        ),
    ),
    "flowChartInternalStorage": (
        _p(
            1,
            1,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("1", "0"),)),
            ("lnTo", (("1", "1"),)),
            ("lnTo", (("0", "1"),)),
            ("close", ()),
        ),
        _p(
            8,
            8,
            ("moveTo", (("1", "0"),)),
            ("lnTo", (("1", "8"),)),
            ("moveTo", (("0", "1"),)),
            ("lnTo", (("8", "1"),)),
        ),
        _p(
            1,
            1,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("1", "0"),)),
            ("lnTo", (("1", "1"),)),
            ("lnTo", (("0", "1"),)),
            ("close", ()),
        ),
    ),
    "flowChartManualInput": (
        _p(
            5,
            5,
            ("moveTo", (("0", "1"),)),
            ("lnTo", (("5", "0"),)),
            ("lnTo", (("5", "5"),)),
            ("lnTo", (("0", "5"),)),
            ("close", ()),
        ),
    ),
    "flowChartProcess": (
        _p(
            1,
            1,
            ("moveTo", (("0", "0"),)),
            ("lnTo", (("1", "0"),)),
            ("lnTo", (("1", "1"),)),
            ("lnTo", (("0", "1"),)),
            ("close", ()),
        ),
    ),
    "flowChartPunchedCard": (
        _p(
            5,
            5,
            ("moveTo", (("0", "1"),)),
            ("lnTo", (("1", "0"),)),
            ("lnTo", (("5", "0"),)),
            ("lnTo", (("5", "5"),)),
            ("lnTo", (("0", "5"),)),
            ("close", ()),
        ),
    ),
    "lineInv": (
        _p(
            None,
            None,
            ("moveTo", (("l", "b"),)),
            ("lnTo", (("r", "t"),)),
        ),
    ),
    "rect": (
        _p(
            None,
            None,
            ("moveTo", (("l", "t"),)),
            ("lnTo", (("r", "t"),)),
            ("lnTo", (("r", "b"),)),
            ("lnTo", (("l", "b"),)),
            ("close", ()),
        ),
    ),
}
"""Mapping of preset identifier (`prst` attribute value) to its decoded path list.

Only static shapes are included (those with no `a:avLst` adjustments and no computed
`a:gdLst` guides). Keys are the `prst` string literal used in the `a:prstGeom` element
(e.g. `"rect"`, `"flowChartProcess"`).
"""


def is_static_preset(prst: str) -> bool:
    """True if `prst` identifies a preset shape whose geometry is fully static."""
    return prst in STATIC_PRESETS


def resolve_token(token: str, w: int, h: int) -> int:
    """Return integer coordinate value for `token` given shape width `w` and height `h`.

    `token` is one of the named constants (e.g. `"hc"`, `"r"`) or an integer literal.
    """
    if token == "l" or token == "t":
        return 0
    if token == "r":
        return w
    if token == "b":
        return h
    if token == "hc" or token == "wd2":
        return w // 2
    if token == "vc" or token == "hd2":
        return h // 2
    # -- integer literal; path-local coordinate --
    return int(token)
