# python-pptx — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

- **FU-1: `_FontColorFormat` cache staleness (from W17 #537).** On a fresh
  run with no `a:solidFill`, `font.color.rgb = RGBColor(...)` writes the
  XML correctly, but subsequent reads of `font.color.rgb` raise because
  `_FontColorFormat._promote()` caches the pre-promote `_NoneColor`
  proxy into `font.__dict__["color"]` before the base setter upgrades
  to `_SRgbColor`. Fix: invalidate / re-resolve the `lazyproperty`
  after promotion.

- **FU-2: `Chart.value_axis` returns secondary axis on combo charts (from
  W17 #833).** `chart.value_axis` returns the LAST `c:valAx` in the
  chartSpace, so on a combo chart with bar + line + secondary axis it
  hands back the secondary axis. Callers must dig into
  `chart._chartSpace.plotArea.primary_valAx` to touch the primary.
  Options: (a) return primary by default and add
  `chart.primary_value_axis` alias; (b) keep behavior, add
  `chart.primary_value_axis` + deprecation note; (c) error on combo
  charts and require explicit side-selection. Needs user-facing
  decision before any code change.

- **FU-3: Duplicate `slideLayout7.xml` on save-after-reopen (from W12
  #956).** After opening a default deck and re-saving, the zip emits
  `UserWarning: Duplicate name: 'ppt/slideLayouts/slideLayout7.xml'`.
  Two distinct layout parts end up with the same partname when
  `iter_parts()` walks the reachable graph. Investigate whether the
  cloner or the reachability walk is the source.

- **FU-4: `_rel_ref_count` doesn't count `@r:embed` attribute references
  (flagged in early waves).** Causes shared-image refcount
  underreporting. When two shapes share the same image rel, the ref
  count ignores one because it only counts `r:link`-style attributes
  and misses `r:embed` occurrences on `a:blip` and `p:blipFill` leaves.



## Done

- **FU-6 (test-only, landed; surfaced a real chart-GC gap).**
  ``tests/test_fu6_clear_gc.py`` pins that
  ``Slide.clear_shapes()`` + ``save()`` compose with the save-time
  ``iter_parts()`` reachability walk: clearing a picture-bearing
  deck recovers ~60 KB (every ``ppt/media/*`` entry gone, the three
  added image SHAs absent from the reopened package), and
  placeholders are preserved under the default
  ``preserve_placeholders=True``. The investigation also surfaced a
  real gap: :class:`GraphicFrame` carries no ``.delete()`` override,
  so clearing a chart-bearing slide does *not* drop the slide ->
  chart rel and the chart + embedded xlsx linger. Distinct from FU-4
  (refcount bug): the chart rel is never dropped at all. A companion
  scenario in ``test_fu6_clear_gc.py`` pins the current behaviour so
  a future ``GraphicFrame.delete()`` fix flips the assertion loudly;
  the fix itself is left as a follow-up feature (candidate FU-7) for
  explicit scoping.

- **FU-5 (reviewed and dropped).** The scratch branch
  `scratch/wave-15-orphan-parts-residue` carried two WIP commits
  (`ac064932`, `1136c764`) targeting orphan-parts round-trip — the
  same changes to `src/pptx/opc/package.py`, `tests/opc/test_package.py`,
  and a new `tests/test_chartex_roundtrip.py` regression test. Those
  changes were fully superseded by commit `0271c930` on master
  (``fix(opc): preserve chartEx fallback parts + embedded xlsx on
  round-trip``), which landed the identical implementation plus the
  regression test. The scratch branch's `src/pptx/opc/serialized.py`
  delta was additionally stale (lacked the later RMS-protection
  handling now on master). Dropped the scratch branch; orphan-parts
  handling on master is correct — parts declared in
  ``[Content_Types].xml`` but unreachable from the package-root rels
  graph (e.g. ``mc:AlternateContent`` fallback charts) are discovered
  during load, their dependent rels subtree is walked, and empty
  ``.rels`` files that were physically present in the input are
  preserved on save. Verified by ``tests/test_chartex_roundtrip.py``
  and the full pytest run (6557 passed).

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
