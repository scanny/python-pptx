# python-pptx — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

### Fidelity / clone API gaps

Surfaced during a round-trip replication exercise on
``Dickinson_Sample_Slides.pptx`` (2026-05-05) where we wanted to
read an arbitrary deck with ``python-pptx`` and write out a
faithful replica using the library only. The script currently has
to fall through to raw ``lxml`` (``_element`` + ``copy.deepcopy``)
at every fidelity boundary because the proxy layer's setters lose
formatting that readers can see. Each gap below, if shipped, would
eliminate an XML escape hatch from that replicator.

Priority ordering: **CLO-1** and **CLO-2** together collapse the
replicator from ~150 lines of XML gymnastics to ~6 lines of
library calls; everything else is narrower but composes with them.

- **CLO-2: `BaseShape.clone_onto(shape_tree, left=None, top=None) -> BaseShape`.**
  Copy a shape (the full ``<p:sp>`` / ``<p:pic>`` /
  ``<p:graphicFrame>`` / ``<p:cxnSp>`` / ``<p:grpSp>`` subtree) into
  another shape tree, optionally at a new position. Whole-slide
  clones already ship (``Slides.duplicate`` /
  ``Slides.add_slide_from_external``); this is the missing
  per-shape granularity. Must rewrite shape-id to avoid collisions
  in the destination, and re-embed any referenced image / chart /
  OLE / media / SmartArt / 3D-model parts just like FU-7's
  ``GraphicFrame.delete`` inversed (drop rels → add rels).

- **CLO-3: `BaseShape.replace_spPr_from(other_shape) -> Self`.**
  Clone fill / line / effect / geometry-hints from another shape
  onto this one. Lets callers mix-and-match "new shape shell, old
  look" without touching ``_element``. Composes with CLO-2 (which
  replaces the whole shape) as a narrower alternative.

- **CLO-6: `_Paragraph.clone_from(other_paragraph) -> Self` and
  `_Run.clone_from(other_run) -> Self`.**
  The next layer down from CLO-5 — per-paragraph / per-run
  formatting clone. Covers the case where a caller wants to mix
  text from two sources inside one text frame.

- **CLO-7: Extended `BaseShape.theme_style_refs`.**
  W14 #447 shipped read/write for the four ``<p:style>`` index
  refs, but the property always writes the default ``accent1``
  color choice on the four child ``<a:schemeClr>`` entries. Extend
  so callers can preserve or specify each ref's actual color
  choice (``<a:schemeClr val="bg1"/>`` vs ``<a:schemeClr val="accent2"/>``
  etc.). Cleanest shape: a second NamedTuple field like
  ``ThemeStyleRefColors(line, fill, effect, font)``.

- **CLO-8 (partial): `Chart.clone_from(source_chart) -> Self` /
  `Chart.apply_template(source_chart)` accepting a live Chart.**
  The primary deliverable — ``SlideShapes.add_chart_from`` — shipped;
  see *Done* below. Read-mutate ``Chart.clone_from`` (replace this
  chart's ``c:chartSpace`` with the source's while preserving the
  target part's partname and rebuilding its rels) is still open;
  tracked as a follow-up because it requires draining the chart
  part's existing relationships graph (embedded xlsx, user-shapes,
  image, theme-override) before re-running the F1 / F5 cloning
  against the existing part.

- **CLO-9: `_Cell.clone_from(other_cell) -> Self`.**
  Copy runs, borders (including diagonals), fill, margins,
  vertical anchor, and padding from another cell. Today
  ``cell.text = "..."`` loses every one of those.

<<<<<<< HEAD
- **CLO-11: `GroupShape.clone_onto(shape_tree) -> GroupShape`.**
  Nominally covered by CLO-2 in the protocol sense, but groups
  carry nested content that needs correct recursion (each child
  shape's local transform + the group's own ``chOff`` / ``chExt``
  viewport). Call out explicitly so the implementation of CLO-2
  doesn't silently degrade on groups.
=======
- **CLO-10: `Table.clone_from(source_table) -> Self`.**
  Whole-table version of CLO-9 — column widths + row heights +
  every cell (delegating per-cell to CLO-9) + table style +
  header/first-row/banded-rows flags.
>>>>>>> feat/clo11-group-shape-clone

- **CLO-12: `SlideShapes.add_picture_from(other_picture, left=None, top=None) -> Picture`.**
  One-call equivalent of "extract blob, re-embed as picture, copy
  spPr". Saves the ``BytesIO(other.image.blob)`` dance and
  automatically preserves crop / outline / effects. A narrow
  alias CLO-2 makes unnecessary but is the most ergonomic name
  for the common picture-only case.




## Done

<<<<<<< HEAD
<<<<<<< HEAD
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
=======
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
>>>>>>> feat/clo4-5-text-frame-clone-from
=======
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
>>>>>>> feat/clo11-group-shape-clone

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
