# python-pptx — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

_None tracked yet._


## Done

- **Wave 24 (overnight polish pass).** Five parallel agents, all merged
  to master (``5adf390b..8aa60af5``).
  - **24-A deeper types**: pyright ``src/pptx`` 6255 → 6208 (-47),
    ruff ``src/pptx`` 50 → 0 (-50). 27 files touched. Caught two real
    bugs: missing ``tbl is not None`` guard in
    ``GraphicFrame.table``, and ``TextFrame.rotation`` return type
    was ``float | None`` where it should be ``float``.
    Branch ``chore/wave-24-a-deeper-types``.
  - **24-B Sphinx warnings**: actionable (non-epilog) warnings
    1120 → 0, total 13942 → 10186. Fixed by adding ``.. module::``
    directive to ``docs/api/animation.rst``, enum symbol
    registration for ``MSO_ANIMATION_TYPE`` / ``_TRIGGER``, a
    disambiguated cross-reference in pictures-svg, and a targeted
    349-entry ``nitpick_ignore`` list in ``docs/conf.py`` covering
    internal proxy classes. Branch
    ``chore/wave-24-b-sphinx-warnings``.
  - **24-C skipped behave scenarios**: the 2 scenarios in
    ``features/interop-validate.feature`` intentionally skip when
    a developer-local PowerPoint fixture, the ``ooxml-validate``
    sibling, or LibreOffice is absent. Pinned with an explanatory
    header comment documenting the three skip triggers. Branch
    ``chore/wave-24-c-skipped-scenarios``.
  - **24-D corpus conformance setup**: ``ooxml-validate`` is a
    sibling ``loadfix`` git-only package (not on PyPI), and
    ``../ooxml-reference-corpus/`` is similarly git-only. Added a
    "Corpus conformance tests" section to ``docs/dev/runtests.rst``
    explaining the sibling-checkout + ``pip install -e`` workflow.
    Tests continue to auto-skip gracefully when siblings absent.
    Branch ``build/wave-24-d-corpus-conformance``.
  - **24-E behave coverage round 2**: +34 scenarios (1501 → 1535)
    in ``features/prs-round-trip-2.feature`` covering group-shape
    mixed content + nested, sections (rename/remove/reorder/
    cross-section-move), transitions (speed/direction/advance),
    shape animations (fade_in/pulse/fade_out/clear/sequence),
    extended properties, password-protected round-trip (requires
    ``msoffcrypto-tool``), chartex deep double round-trip,
    cross-presentation ``add_slide_from_external`` +
    ``Presentation.merge``. Branch ``chore/wave-24-e-behave-round-2``.

- **Wave 23-B type cleanup.** Reduced pyright strict-mode error count by
  removing unnecessary ``# type: ignore`` / ``# pyright: ignore``
  comments, pruning unused imports, and adding missing parameter
  annotations on fork-era module-level helpers. Scope: ``src/pptx/api.py``,
  ``src/pptx/animation.py``, ``src/pptx/oxml/chart/series.py``,
  ``src/pptx/oxml/text.py``, ``src/pptx/oxml/timing.py``,
  ``src/pptx/parts/extprops.py``, ``src/pptx/parts/image.py``,
  ``src/pptx/shapes/base.py``, ``src/pptx/slide.py``, plus mirrored
  cleanups in ``tests/``. All 6619 pytest tests still pass.
  Branch ``chore/wave-23-b-type-cleanup``.

- **FU-7 (fixed).** :class:`~pptx.shapes.graphfrm.GraphicFrame` now
  overrides :meth:`~pptx.shapes.base.BaseShape.delete` to drop every
  slide-part rel the graphic frame carries — classic ``c:chart``,
  Office 2016+ extended ``cx:chart``, embedded or linked ``p:oleObj``
  plus the companion icon image ``a:blip/@r:embed``, the four
  SmartArt ``dgm:relIds`` attributes, and ``am3d:model3D/@r:embed``.
  Closes the chart-GC gap surfaced by FU-6: calling
  ``slide.shapes.clear()`` on a chart-bearing slide now composes with
  the save-time :meth:`iter_parts` reachability walk so the chart
  part and its embedded xlsx dependency are pruned on the next save.
  Inverted the FU-6 ``but_chart_parts_are_NOT_gc_ed_by_clear_shapes``
  pin to ``it_gcs_chart_parts_on_clear_shapes``, added
  ``tests/test_fu7_graphfrm_delete_rels.py`` (chart / chartex / OLE /
  SmartArt / 3D-model / table coverage), and expanded
  ``tests/shapes/test_graphfrm.py`` with per-kind delete-and-drop-rel
  unit tests. Branch ``fix/fu7-graphic-frame-delete-rels``.

- **FU-1 (fixed).** `_FontColorFormat._promote()` now drops the cached
  `color` entry on `font.__dict__` and returns the live `ColorFormat`
  (`font.fill.fore_color`), so the deferred-promotion proxy no longer
  leaves a stale pre-promotion `_NoneColor` in the `lazyproperty`
  cache slot after a first `font.color.rgb = RGBColor(...)` or
  `.theme_color = ...` assignment. Regression suite
  `tests/test_fu1_font_color_cache.py` (7 scenarios: held-reference,
  chained-access, `color.type` read, theme-color, double-write,
  round-trip save-and-reopen, cache-consistency). Branch
  `fix/fu1-font-color-cache`.

- **FU-2 (shipped as additive API).** Added `Chart.primary_value_axis`
  returning the first `c:valAx` in the plot area, unambiguous on
  combo charts where `Chart.value_axis` legacy "last-wins" behaviour
  resolves to the secondary axis. Both accessors raise `ValueError`
  when the chart has no value axis; `Chart.value_axis` docstring now
  spells out the combo-chart caveat explicitly. FEATURES.md bullet,
  `docs/user/charts.rst` caution block, `docs/api/chart.rst` entry,
  new unit tests in `tests/chart/test_chart.py`, behave scenario,
  and `tests/test_issue_833_two_y_axis_verify.py` all updated.
  Branch `feat/fu2-primary-value-axis`.

- **FU-4 (fixed).** `XmlPart._rel_ref_count` xpath extended to count
  every rId-bearing attribute (`@r:id`, `@r:embed`, `@r:link`)
  instead of only `@r:id`. Fixes shared-image shapes (two pictures
  from the same image file sharing a `@r:embed` on `a:blip`) where
  deleting one would undercount the remaining reference, cause
  `drop_rel` to remove the shared relationship prematurely, and
  leave the second picture with a dangling rId. Same bug affected
  `p14:media`, `a:videoFile`, `a:audioFile`, linked-image `a:blip`,
  `am3d:model3D`. Regression test in
  `tests/test_fu4_rel_ref_count_embed.py`. Branch
  `fix/fu4-rel-ref-count-embed`.

- **FU-3 (investigated, not reproducible, pinned as invariant).** Could
  not reproduce the W12 #956 report of ``UserWarning: Duplicate name:
  'ppt/slideLayouts/slideLayout7.xml'`` on save-after-reopen of a
  default deck. Tried default save-reopen-save, triple cycle, every
  built-in layout used, and chart + notes round-trip — none emit a
  duplicate-name warning. The saved zip contains exactly one entry per
  ``slideLayout{1..11}.xml`` on every cycle. The original report likely
  referred to a deck manipulated by in-progress Wave 12 #956 clone /
  delete code paths that did not land on master. Pinned the current
  correct behavior in ``tests/test_fu3_duplicate_layout_partname.py``
  (8 tests covering the warning-free invariant and the zip-entry
  uniqueness invariant) as a regression guard. Branch
  ``feat/fu3-verify-no-dupe-warning``.

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
