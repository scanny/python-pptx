# python-pptx — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Audit findings 2026-05-05

Captured from a project audit run on 2026-05-05. Each bullet is an
actionable follow-up; promote into "Open" (or straight to "Done") as
work lands.

1. **Cut release for Unreleased CLO-1..12 ledger.** Master is
   `8e382107`, 23+ commits past the `2026.05.2` tag released earlier
   today. The Unreleased changelog section logs CLO-1 through CLO-12
   cross-tree cloning (`Slide.clone_shapes_from`, `TextFrame.clone_from`,
   `BaseShape.clone_onto`, `Table.clone_from`, `_Cell.clone_from`,
   `add_chart_from`, `add_picture_from`, `GroupShape.clone_onto`) plus
   two slide-clone hotfixes. Warrants a `2026.05.3` tag.
2. **Close GitHub issue #3** — bare `KeyError('[Content_Types].xml')`
   raised on a package missing its content-types part. Wrap in a typed
   `PackageNotFoundError` (or similarly descriptive exception). Only
   open issue on the tracker.
3. **Close CLO-8 follow-up** — in-place
   `Chart.clone_from(source_chart) -> Self`. **(Done — see CLO-8
   follow-up entry under "Done" below; branch
   `feat/clo8-chart-clone-from-inplace`.)**
4. **Bump README version string.** README currently advertises
   `2026.05.0`; should be `2026.05.2` (or `2026.05.3` once item 1
   ships).
5. **Seal submodule oxml leakage.** 665 `CT_*` / `ST_*` names are
   reachable via `pptx.oxml.*` and friends — up from the W11-E baseline
   of ~500 because the CLO series added XML classes. Add explicit
   `__all__` on the top-level package and the major subpackages so the
   public surface is enumerable.
6. **Update `pyproject.toml` classifiers.** Classifiers stop at
   Python 3.12 despite `requires-python >= 3.8`. Add 3.13, and drop 3.8
   if the runtime test matrix allows.
7. **Move or delete scratch audit artefacts.** `DOCS_AUDIT.md` (58 KB)
   and `TEST_AUDIT.md` (29 KB) are stale internal meta-documents
   tracked in the repo root. Relocate to an `audits/` directory or
   delete outright.
8. **Git-tag 2026.05.1 and 2026.05.2.** Verify both releases carry
   git tags; add them if missing.

## Open

(none — all tracked items in the Audit-findings section above)


## Done

- **CLO-8 (follow-up): ``Chart.clone_from(source_chart) -> Self``.**
  In-place read-mutate counterpart to the CLO-8 primary
  ``SlideShapes.add_chart_from``. Drops every rel referenced by this
  chart's current ``c:chartSpace`` (embedded xlsx, user-shapes, chart
  images, theme-override, chart-style / chart-colors), deep-copies the
  source's ``c:chartSpace`` with every ``r:id`` / ``r:embed`` /
  ``r:link`` rewritten against freshly-allocated rels on this part,
  replaces the chartSpace content in place (element identity preserved
  so caller-held ``Chart`` references remain valid), and re-embeds the
  source's ``.xlsx`` workbook as an independent copy so PowerPoint's
  "Edit Data" dialog stays wired per-chart. The chart part's partname
  is preserved throughout, so the enclosing ``p:graphicFrame`` on the
  slide and its rel still reference this chart. Returns ``self`` for
  chaining. Supports same-presentation and cross-presentation sources.
  Branch ``feat/clo8-chart-clone-from-inplace``.

- **CLO-1: ``Slide.clone_shapes_from(other_slide, include_placeholders=True)
  -> None``.** Composite wrapper on top of ``BaseShape.clone_onto``: walks
  ``other_slide.shapes`` and appends a deep-copy of each top-level shape
  onto this slide's shape tree, re-embedding every referenced package
  part (images, charts, OLE, SmartArt, 3D-model media, hyperlinks) and
  assigning fresh unique ``cNvPr/@id`` / ``@name`` to each clone. Source
  z-order is preserved. Placeholders are handled specially because CLO-2
  raises on them; the default ``include_placeholders=True`` deep-copies
  the source placeholder's text body onto this slide's matching-``idx``
  placeholder via ``_Paragraph.clone_from`` (or skips when no matching
  idx), and ``include_placeholders=False`` skips placeholders entirely.
  Source is not mutated; returns ``None``. Branch
  ``feat/clo1-slide-clone-shapes-from``.
- **CLO-4 / CLO-5: whole-text-frame clone (paired).**
  Added ``TextFrame.clone_from(other_text_frame) -> Self`` and
  ``BaseShape.replace_text_frame_from(other_shape) -> Self``.
  ``TextFrame.clone_from`` replaces this text frame's ``txBody``
  children (``a:bodyPr``, ``a:lstStyle``, every ``a:p``, any trailing
  ``a:endParaRPr``) with deep-copies of the source's; the destination
  ``txBody`` element is kept intact so the enclosing shape's wiring is
  preserved. ``BaseShape.replace_text_frame_from`` is the
  shape-scoped wrapper — delegates to ``TextFrame.clone_from`` once
  both shapes are confirmed to have a text frame; raises
  ``ValueError`` when either lacks one (connector, chart or table
  graphic frame, group shape). Replaces the fidelity-losing idiom
  ``shape.text_frame.text = src.text_frame.text``. Both return
  ``self`` for chaining and deep-copy (sources untouched, destinations
  independent). Works across slides / presentations. Branch
  ``feat/clo4-5-text-frame-clone-from``.
- **CLO-11: `GroupShape.clone_onto` correctness verified.**
  CLO-2 (``BaseShape.clone_onto``) already handled groups correctly —
  this task added a dedicated regression suite confirming it. Covered
  by ``tests/test_clo11_group_clone.py`` (9 pytest tests) and five new
  scenarios in ``features/shp-clone-onto.feature`` exercising: a group
  of three autoshapes with preserved child positions; nested groups
  with every descendant surviving and unique fresh ids; a group
  containing a picture re-embedding its image blob; a group containing
  a chart with every ``r:id`` resolving on the target slide-part; an
  outer-group position override that preserves the child coordinate
  system (``a:chOff`` / ``a:chExt`` unchanged); plus cross-presentation
  re-embedding and round-trip save/reload. No production-code changes
  were needed.
- **CLO-12: ``SlideShapes.add_picture_from(other_picture, left=None,
  top=None, width=None, height=None) -> Picture``.** Ergonomic picture-
  specific alias for ``BaseShape.clone_onto``. Re-embeds the source's
  image bytes on this slide's package and automatically preserves crop
  (``a:srcRect``), outline, effects, rotation, and flip — everything
  the manual ``BytesIO(src.image.blob)`` + ``add_picture`` dance
  discards. ``left`` / ``top`` / ``width`` / ``height`` default to
  ``None`` (preserves the source value). Supports intra- and cross-
  presentation sources. Branch ``feat/clo12-add-picture-from``.


- **CLO-8 (primary deliverable): full-fidelity chart-clone front door.**
  Added ``SlideShapes.add_chart_from(source_chart, left, top,
  width=None, height=None)`` — an ergonomic wrapper on the existing
  ``clone_chart`` primitive that emits a deep clone preserving every
  styling attribute the source carries (title text, axis label fonts,
  series fills, plot-area position, trendlines, error bars, data
  labels, legend formatting), where ``add_chart(type, data)`` emits
  fresh default-styled XML. ``width`` and ``height`` default to 5 × 3
  inches. Works intra- and cross-presentation. Read-mutate
  ``Chart.clone_from`` is still open as a follow-up (see CLO-8 partial
  entry above). Branch ``feat/clo8-chart-clone-from``.
- **CLO-9: ``_Cell.clone_from(other_cell) -> Self``.** Copies full
  ``a:txBody`` (paragraphs, runs, run-level formatting) and ``a:tcPr``
  (fill, edge and diagonal borders, margins, vertical anchor) subtrees
  from one cell onto another while preserving position and merge-state.
  Returns ``self``; source is not mutated. Branch
  ``feat/clo9-cell-clone-from``.

- **CLO-10: ``Table.clone_from(source_table) -> Self``.** Whole-table
  counterpart to CLO-9: copies every column width, every row height,
  every cell (delegating per-cell to ``_Cell.clone_from``), the source's
  ``style_id`` when present, and all six header / banding flags. Returns
  ``self``; source is not mutated. Raises ``ValueError`` when dimensions
  differ. Works intra- and cross-presentation. Branch
  ``feat/clo10-table-clone-from``.

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
