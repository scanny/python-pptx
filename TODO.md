# python-pptx — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

_None tracked yet._

## Done

- Issue #627 — author tables and charts (plus textbox / picture / auto-shape /
  connector) directly inside a `GroupShape` via `group.shapes.add_table(...)`
  etc. `add_table` promoted from `SlideShapes` onto the shared
  `_BaseGroupShapes` base so every subclass inherits it. Branch
  `feat/issue-627-group-shapes-authoring`.

- Issue #872 — set marker border color on line-marker charts. Verified the
  existing `_MarkerMixin.marker` + `Marker.format` surface already supports
  the recipe `series.marker.format.line.color.rgb = RGBColor(...)` and
  emits the expected `c:ser/c:marker/c:spPr/a:ln/a:solidFill/a:srgbClr @val`.
  Pinned with `tests/test_issue_872_marker_border_color_verify.py` +
  new behave scenario in `features/cht-marker-props.feature`. Branch
  `feat/issue-872-marker-border-color`.

- Issue #1047 — box-and-whisker chartex passthrough MVP. Added
  `GraphicFrame.chartex_type` returning the raw `cx:series/@layoutId`
  string (e.g. `"boxWhisker"`, `"waterfall"`, `"funnel"`, ...) so callers
  can distinguish chartex kinds without parsing the chartex part
  themselves. Detection + round-trip preservation are the shipped
  contract; structured authoring remains blocked on F4 (see
  `docs/dev/analysis/chartex-foundation.rst`). Pinned with
  `tests/test_issue_1047_box_whisker_passthrough.py`; design target for
  the eventual authoring follow-up captured in
  `docs/dev/analysis/chartex-box-whisker.rst`. Branch
  `feat/issue-1047-box-whisker-passthrough`.
