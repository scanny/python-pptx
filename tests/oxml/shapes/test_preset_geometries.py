"""Unit-test suite for `pptx.oxml.shapes.preset_geometries` module."""

from __future__ import annotations

import pytest

from pptx.oxml.shapes.preset_geometries import (
    STATIC_PRESETS,
    PresetOp,
    PresetPath,
    is_static_preset,
    resolve_token,
)


class DescribeStaticPresetsRegistry(object):
    """Unit-test suite for the `STATIC_PRESETS` registry."""

    def it_includes_the_rectangle_preset(self):
        assert "rect" in STATIC_PRESETS

    def it_omits_dynamic_presets(self):
        # -- roundRect is a dynamic preset (has an adjustment) --
        assert "roundRect" not in STATIC_PRESETS
        assert "chevron" not in STATIC_PRESETS

    def it_returns_a_tuple_of_PresetPath_records(self):
        paths = STATIC_PRESETS["rect"]
        assert isinstance(paths, tuple)
        assert len(paths) == 1
        path = paths[0]
        assert isinstance(path, PresetPath)
        assert path.w is None
        assert path.h is None
        assert all(isinstance(op, PresetOp) for op in path.ops)

    def it_exposes_flowChartProcess_with_a_1x1_local_coord_system(self):
        paths = STATIC_PRESETS["flowChartProcess"]
        assert len(paths) == 1
        assert paths[0].w == 1
        assert paths[0].h == 1

    def it_exposes_chartPlus_as_two_sub_paths(self):
        paths = STATIC_PRESETS["chartPlus"]
        assert len(paths) == 2


class DescribeIsStaticPreset(object):
    """Unit-test suite for `is_static_preset`."""

    def it_returns_True_for_known_static_preset(self):
        assert is_static_preset("rect") is True
        assert is_static_preset("flowChartProcess") is True

    def it_returns_False_for_unknown_preset(self):
        assert is_static_preset("roundRect") is False
        assert is_static_preset("bogus") is False


class DescribeResolveToken(object):
    """Unit-test suite for `resolve_token`."""

    @pytest.mark.parametrize(
        ("token", "w", "h", "expected"),
        [
            ("l", 100, 200, 0),
            ("t", 100, 200, 0),
            ("r", 100, 200, 100),
            ("b", 100, 200, 200),
            ("hc", 100, 200, 50),
            ("wd2", 100, 200, 50),
            ("vc", 100, 200, 100),
            ("hd2", 100, 200, 100),
            ("5", 100, 200, 5),
            ("10", 100, 200, 10),
        ],
    )
    def it_resolves_named_and_integer_tokens(self, token: str, w: int, h: int, expected: int):
        assert resolve_token(token, w, h) == expected
