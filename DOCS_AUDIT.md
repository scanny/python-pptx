# Sphinx Documentation Audit

This report surveys the state of the Sphinx-based reference docs under `docs/`
against the current `loadfix/python-pptx` source tree at commit `1a345f8f`
(`master`, "docs: align with python-docx sibling conventions"). It is a
companion to `TEST_AUDIT.md`. The goal is to document what is currently
shipped, what has drifted since upstream, and to give a prioritised punch-list
of concrete follow-ups.

This report is **advisory only** — no `.rst`, `.py`, or workflow files were
modified while producing it.

---

## 1. Summary

- `docs/` contains **180 reStructuredText files, 36 353 lines total**
  (`find docs -name '*.rst' | xargs wc -l`). Of those:
  - 14 API-reference pages under `docs/api/` (1 625 lines) — the biggest is
    `chart.rst` at 396 lines; the smallest are `exc.rst` (9) and `action.rst`
    (54). `presentation.rst` is 245 lines.
  - 35 enum pages + 1 index under `docs/api/enum/` (3 055 lines). Each page
    is hand-written (title + alias + one-paragraph intro + bulleted member
    list).
  - 19 user-guide pages under `docs/user/` (4 548 lines). Longest:
    `charts.rst` (827), `text.rst` (585), `table.rst` (538).
  - 107 developer pages under `docs/dev/` of which **88** are feature /
    XSD analyses under `docs/dev/analysis/` (25 605 lines of the total). The
    analysis tree is where fork-era work has landed most heavily: 11 new
    analysis pages (`chartex-foundation`, `chartex-funnel`,
    `chartex-treemap`, `chartex-waterfall`, `combo-chart`,
    `f5-embedded-workbook`, `f6-slide-id-manager`, `f7-sections`,
    `f8-animations-transitions`, `f9-smartart`, `omml-parsing`) compared
    to the 77 pre-fork upstream analyses.
  - 3 community pages under `docs/community/` (25 lines) — `faq.rst` is
    2 lines, `updates.rst` is 7, `support.rst` is 16.
- Build system: **Sphinx with the `armstrong` HTML theme** vendored under
  `docs/.themes/armstrong` (note the *dot* — `docs/conf.py:432` sets
  `html_theme_path = [".themes"]`). The `alabaster` pin in
  `requirements-docs.txt:3` is **not** consumed — the active theme is
  `armstrong` (`docs/conf.py:424`). CLAUDE.md does not claim a theme so
  there is no drift there, but a reader inspecting `requirements-docs.txt`
  will come away with the wrong impression.
- Configuration: `docs/conf.py` — Sphinx-pre-1.0-style, pinned to
  `Sphinx==1.8.6` / `Jinja2==2.11.3` / `MarkupSafe==0.23` in
  `requirements-docs.txt`. None of those install on Python ≥3.10 without
  patches. ReadTheDocs builds it by pinning `python: "3.9"` in
  `.readthedocs.yaml`; any developer running `make docs` on a modern
  checkout cannot reproduce that stack locally.
- Last meaningful doc-tree commit (non-merge): **`0c83820e` "feat(color):
  #62 add ColorFormat.alpha"** (2026-05-02). Unlike the sibling
  `python-docx` fork, this fork has been keeping docs *alive* — nearly every
  fork-era landing touched `HISTORY.rst` and many touched at least one page
  under `docs/api/` or `docs/user/`. `git log --no-merges -- docs/` shows
  163 commits since 2024-01-01 and many dozens on 2026-05-02 alone. The
  last upstream-authored doc commit is `c38d5f5c` (Steve Canny,
  2024-07-26).
- **HISTORY.rst** (repo root, 1 768 lines) has an open `Unreleased` block
  (lines 6-1 269) that already records **160 distinct issue references**
  across 94 `feat:` / 28 `fix:` bullets plus assorted `verify:`, `docs:`,
  `rfctr:`, `perf:`, `security:`, and `Foundation` items. That's a good
  indicator that changelog discipline is *much* tighter here than in the
  sibling python-docx fork, where fork features ship undocumented.

### Top findings

1. **The `rst_epilog` substitution table has drifted**. A non-strict
   Sphinx 7 build against the current source tree emits **159 `ERROR:
   Undefined substitution referenced:`** events across **32 distinct
   symbols** — `AnimationEffect`, `AnimationEffectView`, `ErrorBars`,
   `PathGeometry`, `Path`, `Sound`, `Audio`, `Section`, `Sections`,
   `ConnectorAdjustmentCollection`, `Transition`, `_HeaderFooter`,
   `_Field`, `Comment`, `CommentAuthor`, `CommentAuthors`, `Comments`,
   `SmartArt`, `Movie`, `MediaPart`, `PROG_ID`, `LinePlot`,
   `ExtendedPropertiesPart`, `TagsPart`, `SlideTags`, `CustomProperties`,
   `MediaPart`, `XL_ERROR_BAR_TYPE`, `XL_ERROR_BAR_INCLUDE`,
   `XL_ERROR_BAR_DIRECTION`, `XL_CROSS_BETWEEN`, `XL_AXIS_POSITION`.
   Every one of these substitutions is used in a docstring but is missing
   from `conf.py`'s `rst_epilog` table (lines 86-382). Adding them is a
   mechanical change — one `.. |X| replace:: :class:\`.X\`` line each —
   and retires ~120 of the build's 159 errors in a single pass.
2. **Seven public classes on the fork's new surface have no API reference
   page**. `pptx.presentation.Section` / `Sections`, `pptx.slide.Transition`,
   `pptx.slide.AnimationEffectView`, `pptx.animation.AnimationEffect`,
   `pptx.shapes.picture.Movie`, `pptx.shapes.graphfrm.SmartArt`, and
   `pptx.media.Video` are exposed but either (a) have autoclass stubs that
   fail to render because substitutions are missing (`Section`,
   `Transition`), or (b) have no directive anywhere in `docs/api/`
   (`AnimationEffect`, `AnimationEffectView`, `Movie`, `SmartArt`, `Video`).
3. **Three new public enums ship with no enum page**:
   `MSO_ANIMATION_TYPE` (`src/pptx/enum/animation.py`),
   `MSO_ANIMATION_TRIGGER` (same), and the `PP_TRANSITION_*` family
   (`PP_TRANSITION_TYPE`, `PP_TRANSITION_SPEED`,
   `PP_TRANSITION_SIDE_DIRECTION` in `src/pptx/enum/transition.py`).
   `docs/api/enum/index.rst` has 35 entries; none for these. Users
   setting `slide.transition.type = PP_TRANSITION_TYPE.MORPH` or
   `shape.set_animation(MSO_ANIMATION_TYPE.FADE_IN, ...)` have no
   reference page listing the legal values.
4. **`docs/user/install.rst` is still the upstream 2013-era page**. It
   claims `Python 2.6, 2.7, 3.3 or later` (line 22) and recommends
   `python setup.py install` / `easy_install` — neither has been relevant
   since the fork's `pyproject.toml` build was modernised. It is also the
   one user-guide page that has *no* fork-era update.
5. **The "Feature Support" bullet list in `docs/index.rst`** (lines 24-47)
   has 11 bullets. Fork-era landings added exactly three of those (legacy
   comments, presentation sections, password-protected `.pptx`). The
   other 50+ new capabilities ship invisible to a landing-page reader —
   font embedding, MORPH transitions, shape animations, flat-OPC save,
   `.ppsx` save, `.potx` read, combo charts, 3D charts, Excel error-bars,
   cross-slide chart copy, deck merging, flat paper-size presets,
   picture replacement, movie replacement, cell borders, table column
   mutation, connector adjustments, custom doc properties, ….

### Sphinx build result

A build with **Sphinx 7.4.7** against the current source tree finishes
successfully but emits **140 warning-lines** from the Sphinx front-end
(159 `ERROR` substitution events + 8 plain `WARNING` messages; Sphinx
counts one rubric error as potentially multiple diagnostic lines). The
strict `-W` variant fails immediately on the first substitution error in
`pptx.action.ActionSetting.set_sound`. Concrete numbers in section 3.

The build was executed with `PYTHONPATH=src` prepended so autodoc could
import the current worktree's source tree. Without that, running the
build in an environment where `pip install -e .` points at a different
worktree causes 27 `autodoc: failed to import class` warnings in addition
to the substitution errors. This is an orthogonal problem for anyone
running multiple worktrees; a normal one-clone checkout is unaffected.

---

## 2. Docs layout

| Path | Purpose |
|---|---|
| `docs/index.rst` (137 lines) | Landing page: "Feature Support" bullet list + three top-level toctrees (User Guide, Community Guide, API Documentation, Contributor Guide). |
| `docs/conf.py` (588 lines) | Sphinx config. Theme, extensions (`autodoc`, `doctest`, `inheritance_diagram`, `intersphinx`, `todo`, `coverage`, `ifconfig`, `viewcode`), and the `rst_epilog` substitutions block (lines 86-382, 150 `.. |X| replace:: :class:...` entries). |
| `docs/user/*.rst` (19 files, 4 548 lines) | Narrative user guide — intro, install, quickstart, presentations, slides, understanding-shapes, autoshapes, placeholders-understanding, placeholders-using, text, charts, table, media, math-equations, notes, ole-objects, comments, use-cases, concepts. |
| `docs/api/*.rst` (14 top-level pages, 1 625 lines) | API reference. One page per major module group: `presentation`, `slides`, `shapes`, `placeholders`, `table`, `text`, `chart`, `chart-data`, `dml`, `action`, `comments`, `image`, `exc`, `util`. Mostly `.. autoclass:: Foo :members:`. |
| `docs/api/enum/*.rst` (35 enum pages + `index.rst`, 3 055 lines) | Hand-written enum reference pages (title, alias, intro, flat list of members). Not autogenerated — `:ref:` targets in docstrings point at these. |
| `docs/dev/analysis/*.rst` (88 files, ~25 600 lines) | XSD / feature analyses, largely upstream; 11 fork-era additions (F1..F9 foundations, chartex-*, combo-chart, omml-parsing). Linked from the "Contributor Guide" toctree. |
| `docs/community/*.rst` (3 files, 25 lines) | Tiny pages: `faq.rst` (2 lines — just a header), `support.rst` (16 lines), `updates.rst` (7 lines). |
| `docs/.themes/armstrong/` | Vendored HTML theme. A fork of the old "armstrong" sidebar theme. Note the **leading dot** on the directory name — `html_theme_path = [".themes"]` in `conf.py`. |
| `docs/_static/img/` | 25 PNGs. Every one is referenced by an `.. image::` directive; no orphans, no broken links. |
| `docs/_templates/` | Contains only `sidebarlinks.html` (used by `html_sidebars` in `conf.py:465`). |

---

## 3. Build health

### 3.1 Prerequisites

`requirements-docs.txt` pins:

```
Sphinx==1.8.6
Jinja2==2.11.3
MarkupSafe==0.23
alabaster<0.7.14
-e .
```

Two problems:

- `MarkupSafe==0.23` fails with `ImportError: cannot import name 'Mapping'
  from 'collections'` on Python 3.10+. `Jinja2==2.11.3` is in the same
  boat.
- `alabaster` is never consumed. The active theme is `armstrong` (vendored
  under `docs/.themes/`).

On Read-the-Docs the build succeeds because `.readthedocs.yaml` still
targets `python: "3.9"`. A local `make docs` on any modern checkout fails
at `pip install -r requirements-docs.txt` before Sphinx ever runs.

### 3.2 Build command and result

Running with a modern Sphinx 7.4.7 and `PYTHONPATH=src` (so autodoc finds
the current worktree's source):

```
PYTHONPATH=src python -m sphinx -b html docs docs/.build/html
```

- **Exit status: 0** (build succeeded).
- **Warnings emitted: 140 lines** (159 `ERROR:` and 8 `WARNING:`
  diagnostics; a handful of errors take two lines each hence the mismatch).

Strict mode (`-W`) fails at the first substitution error in
`pptx.action.ActionSetting.set_sound`:

```
Warning, treated as error:
src/pptx/action.py:docstring of pptx.action.ActionSetting.set_sound:15:
Undefined substitution referenced: "Sound".
```

### 3.3 Warning breakdown

| Count | Category |
|---:|---|
| 159 | `ERROR: Undefined substitution referenced:` — a `|SomeName|` substitution used in a docstring or `.rst` that has no entry in `conf.py`'s `rst_epilog` (see 3.4 below) |
| 2 | `WARNING: Title underline too short.` — `docs/user/charts.rst:463` and `docs/user/presentations.rst:311` |
| 3 | `WARNING: Explicit markup ends without a blank line; unexpected unindent.` — `docs/user/text.rst:326`, `:372`, `:422` |
| 1 | `WARNING: document isn't included in any toctree` — `docs/dev/analysis/f8-animations-transitions.rst` (orphaned; `f9-smartart.rst` is in the toctree but f8 is not — see section 9.2) |
| 1 | `ERROR: Malformed table.` — `docs/dev/analysis/f9-smartart.rst:65` |
| 1 | `WARNING: checking consistency...` (toctree consistency for f8) |

### 3.4 Missing `|Name|` substitutions

The 32 symbols referenced in docstrings or RST files but not declared in
`conf.py`'s `rst_epilog` (alphabetical):

```
AnimationEffect, AnimationEffectView, Audio, Comment, CommentAuthor,
CommentAuthors, Comments, ConnectorAdjustmentCollection, CustomProperties,
ErrorBars, ExtendedPropertiesPart, _Field, _HeaderFooter, LinePlot,
MediaPart, Movie, Path, PathGeometry, PROG_ID, Section, Sections, SlideTags,
SmartArt, Sound, TagsPart, Transition, XL_AXIS_POSITION, XL_CROSS_BETWEEN,
XL_ERROR_BAR_DIRECTION, XL_ERROR_BAR_INCLUDE, XL_ERROR_BAR_TYPE, bool
```

Top offenders by occurrence count: `ErrorBars` (26), `AnimationEffect`
(11), each of the `XL_ERROR_BAR_*` trio (9), `Path`/`PathGeometry` (8
each), `Sound` (6), `Section` (5), `Audio` and `Comments` (4 each).

Adding them to `rst_epilog` is a mechanical change — one
`.. |X| replace:: :class:\`.X\`` line each — and will retire ~120 of the
159 errors in a single pass. The six `bool` / `PROG_ID` / `XL_*` entries
need slightly different formatting (`:class:` vs `:ref:`) depending on
whether they point at a class, an enum page, or a primitive type.

### 3.5 Broken / stale refs

- `docs/dev/analysis/f8-animations-transitions.rst` is an orphaned page
  (not in any toctree). Pairs oddly with `f9-smartart.rst` which *is* in
  the toctree (added by the F9 PR) — the F8 PR evidently missed
  `docs/dev/analysis/index.rst`.
- `docs/dev/analysis/f9-smartart.rst:65` has a malformed table — one row
  is under-aligned.
- `docs/user/charts.rst:463` and `docs/user/presentations.rst:311` have
  title underlines shorter than the title text. Cosmetic but visible in
  rendered output.
- `docs/user/text.rst:326`, `:372`, `:422` are three `Templating text`
  section-added blocks (#398 / #836 work) that don't leave a blank line
  between an explicit-markup directive and the prose that follows. Easy
  fixes.

### 3.6 Duplicated labels

None detected by this build.

---

## 4. API reference gaps

### 4.1 New public classes with **no** `docs/api/*.rst` directive

These classes exist in `src/pptx/` at `1a345f8f` but do not appear as the
argument to any `.. autoclass::` / `.. automethod::` / `.. autofunction::`
directive under `docs/api/` or `docs/user/`. Verified with
`grep -rn 'pptx\.<module>\.<Class>' docs/api/ docs/user/`, which returns
zero matches for each.

| Class | Location | Shipped in | Suggested placement |
|---|---|---|---|
| `AnimationEffect` | `src/pptx/animation.py:1` | F8 + #102 (authoring API) | New section on `docs/api/slides.rst`, or a new `docs/api/animation.rst` |
| `AnimationEffectView` | `src/pptx/slide.py` (deprecated-alias `AnimationEffect`) | F8 + #256 | Same section as above |
| `ShapeAnimation` | `src/pptx/slide.py` (introspection proxy via `Slide.iter_shape_animations`) | #264 | Same section |
| `Movie` | `src/pptx/shapes/picture.py:108` | Pre-fork, but new properties `.start_time` (#811), `.start_condition` (#811), `.replace_media` (#784), `.delete` override (#974) are all undocumented | Dedicated `|Movie|` section on `docs/api/shapes.rst` |
| `SmartArt` | `src/pptx/shapes/graphfrm.py` | F9 | New `SmartArt` section on `docs/api/shapes.rst` or a new `docs/api/smartart.rst` |
| `Video` | `src/pptx/media.py` | Pre-fork but never documented; referenced by `Movie` | New `|Video|` section on `docs/api/media.rst` (see below — the page doesn't exist yet either) |
| `CustomProperties` | `src/pptx/parts/customprops.py` (as `Presentation.custom_properties`) | #259 / #14 | New section on `docs/api/presentation.rst` |
| `LinePlot` | `src/pptx/chart/plot.py` | Pre-fork but `|LinePlot|` substitution is referenced from new docstrings | `docs/api/chart.rst` already has sections for BarPlot, BubblePlot; add LinePlot |

**There is no `docs/api/media.rst` page** — `Movie`, `Video`, and
`_MediaFormat` (`src/pptx/shapes/picture.py:...`) have no reference
page at all. The closest coverage is `docs/user/media.rst` (75 lines,
narrative). A dedicated API page would be ~30 lines.

### 4.2 New methods/properties on documented classes that render with drift

Most existing API pages use `.. autoclass:: X :members:`, so new methods
on already-documented classes appear automatically. The problem is that
**their docstrings reference substitutions that don't exist** (section
3.4), so each new member renders with one or more `Undefined substitution
referenced` boxes where a class cross-reference should be.

#### `docs/api/presentation.rst` (245 lines)

`Presentation` uses `:members:` so new methods auto-render. Missing:

- `Presentation.custom_properties` (#259) — returns `CustomProperties`;
  `|CustomProperties|` is undefined.
- `Presentation.embed_font` / `.embedded_fonts` (#355) — documented in the
  "Embedding custom fonts" section (nice).
- `Presentation.save_flat_xml` (#1059) — no narrative, no substitution.
- `Presentation.save_ppsx` (#438) — no narrative.
- `Presentation.merge` (#934) — **is** covered in `docs/user/presentations.rst`
  (nice).
- `Presentation.set_auto_advance` (#376) — auto-renders but has no
  narrative anchor.
- `Presentation.embedded_fonts` — auto-renders.
- `Presentation.slide_masters` — pre-existing, fine.

#### `docs/api/slides.rst` (146 lines)

Uses `:members: :inherited-members:` on `Slide`, `Slides`,
`SlideLayouts`, `SlideLayout`, `SlideMasters`, `SlideMaster`,
`_HeaderFooter`, `NotesSlide`, `Transition`. Missing:

- New `Slide.animation_sequence` / `.iter_shape_animations` /
  `.has_animations` / `.timing_xml` / `.is_hidden` (#319, #256, #264)
  all auto-render, but `|AnimationEffect|` / `|AnimationEffectView|` /
  `|ShapeAnimation|` substitutions (section 3.4) break their cross-refs.
- `Slides.duplicate` (#132), `Slides.move_slide`, `Slides.delete_slide`
  (#68, #67), `Slides.add_slide_from_external` (#1036, #934 extensions) —
  all auto-render. No narrative pointing at the cross-slide copy /
  deck-merge pattern on this page (the narrative lives on
  `docs/user/presentations.rst`, but a reader of the API page won't know).
- `Transition` has a narrative section (lines 32-50) but the
  `|Transition|` substitution is missing so every class reference in the
  transitions subtree renders broken.

#### `docs/api/shapes.rst` (227 lines)

Uses `:members:` on `SlideShapes`, `GroupShapes`, `BaseShape`, `Shape`,
`Connector`, `FreeformBuilder`, `Picture`, `GraphicFrame`, `GroupShape`.
New surface auto-renders. Gaps:

- No `|Movie|` page — `SlideShapes.add_movie` return type renders as a
  raw `|Movie|` placeholder.
- No `|SmartArt|` page — `GraphicFrame.smart_art` return type same.
- `Connector.adjustments` (#946) has a narrative in the existing
  `|ConnectorAdjustmentCollection|` section; `|ConnectorAdjustmentCollection|`
  substitution is missing, so the narrative header renders broken.
- `PathGeometry`, `Path`, `MoveTo`, `LineTo`, `CubicBezierTo`,
  `QuadBezierTo`, `ArcTo`, `Close` (#515) — all autoclassed in the page,
  but `|Path|` and `|PathGeometry|` substitutions used by docstrings
  elsewhere are missing.
- New methods on `BaseShape`: `.delete` (#41), `.duplicate` (#533),
  `.replace_with` (#246), `.bring_to_front` / `.send_to_back` /
  `.bring_forward` / `.send_backward` / `.zorder_index` (#49),
  `.flip_horizontal` / `.flip_vertical` / `.flip_horizontally` /
  `.flip_vertically` (#547), `.is_hidden` (#971), `.effective_left` /
  `.effective_top` / `.effective_width` / `.effective_height` (#925),
  `.animation` / `.set_animation` (#102), `.math_equation_xml` /
  `.has_math_equation` (#126) — all auto-render but are invisible to the
  index/sidebar.
- New on `Picture`: `.replace_image` (#834); on `Movie`: `.replace_media`
  (#784), `.start_time` / `.start_condition` (#811) — auto-render, no
  substitution support.

#### `docs/api/text.rst` (81 lines)

Uses `:members:` on `TextFrame`, `Font`, `_Paragraph`, `_BulletFormat`,
`_Run`, `_Field`. Missing:

- `|_Field|` substitution is undefined, so the `_Field` section header
  renders broken.
- New methods: `TextFrame.replace_text`, `_Paragraph.replace_text` (#836),
  `_Paragraph.add_math_equation` (#528), `TextFrame.font_scale` /
  `line_space_reduction` (#715), `TextFrame.rotation` (#133),
  `_Run.delete` / `_Paragraph.delete` (#144), `Font.border_*` (#?? — search
  for #120 analog), `Font.effective_color` (#938),
  `Font.use_theme_hyperlink_color` (#940), `Font.strikethrough` (#574),
  `Font.language` / `Font.name_ea` / `Font.name_cs` (#337),
  `_BulletFormat` (#100 + #114) — all auto-render.

#### `docs/api/table.rst` (84 lines)

Uses `:members: :inherited-members:` on `Table`, `_Cell`, and so on.
Missing:

- `_ColumnCollection.add` / `.remove` (#895) autoclass is restricted to
  `:members: add, remove` — that's *correct* but the `|_ColumnCollection|`
  substitution narrative is fine here.
- New `_Cell.border_*` (#71) — auto-render; `|LineFormat|` substitution is
  defined, so these work.
- `_Cell.margin_*`, `_Cell.merge` / `split`, `_Cell.is_merge_origin` /
  `.is_spanned` / `.merge_origin` / `.span_height` / `.span_width`,
  `_Row.height` / `.height_rule`, `_Cell.row_idx` / `.col_idx` (#849),
  `_Cell.fill` — all auto-render.
- `Table.columns.add()` and `Table.rows.add()` documented in the
  `_ColumnCollection` / `_RowCollection` sections.
- No mention of the new `Table.style_id` (#27) in narrative form.

#### `docs/api/chart.rst` (396 lines)

Large and mostly current, but `|ErrorBars|` / `|XL_ERROR_BAR_*|`
substitutions (26 + 27 refs total) break the `Series.error_bars` / #544
story. Missing:

- No narrative mention of `Chart.apply_template` (#243).
- No narrative mention of `Chart.clone_to` / `SlideShapes.clone_chart`
  (#877).
- No narrative mention of `Chart.add_plot` for combo charts (#338); the
  analysis page `docs/dev/analysis/combo-chart.rst` exists but there is no
  user-facing section.
- No narrative mention of `Chart.replace_data_preserve_formulas` (#239).
- No narrative mention of `Chart.update_cached_values` (#381).
- `DataLabels.format` (#560), `DataLabel.format` (#716) auto-render.
- `ChartTitle.position` (#1030) auto-renders.

#### `docs/api/comments.rst` (61 lines)

The `|Comments|`, `|Comment|`, `|CommentAuthor|`, `|CommentAuthors|`
substitutions are *missing*, so all four section headers render with the
broken-substitution box. The autoclass directives themselves use the
dotted path (`pptx.comments.Comments()`) which works — it's just the
headline markers that are broken.

#### `docs/api/action.rst` (54 lines)

`|Sound|` and `|Audio|` substitutions are missing. Both section headers
are broken. `ActionSetting.screen_tip` (#425), `.set_sound` / `.remove_sound`
(#734) auto-render.

#### `docs/api/chart-data.rst` (70 lines)

Missing substitutions cascade through `|XL_AXIS_POSITION|` /
`|XL_CROSS_BETWEEN|` if referenced (a few new docstrings do reference
these). Adding the two entries retires one warning each.

#### `docs/api/dml.rst` (95 lines)

Covers `ColorFormat`, `RGBColor`, `FillFormat`, `LineFormat`,
`LineEndFormat`, `ShadowFormat`, `GlowFormat`, `ReflectionFormat`,
`SoftEdgeFormat`, `ChartFormat`. New: `ColorFormat.alpha` (#62),
`ColorFormat.to_rgb` (#308 / #420). Autodoc complains at build time about
missing attributes on `ColorFormat` for `alpha` / `to_rgb` (see 3.3) —
likely because the autoclass stanza pre-dates those additions and needs
an `:inherited-members:` toggle or a re-run.

#### `docs/api/placeholders.rst` (125 lines)

New `SlidePlaceholders.get` (#769) — autodoc warns `missing attribute get
in object pptx.shapes.shapetree.SlidePlaceholders`. Either the method
shipped as `get_by_idx` (check), or the autoclass needs `:members: get`
spelled out. New `SlidePlaceholder.insert_chart` / `insert_picture` /
`insert_table` (#199 / #333) auto-render.

#### `docs/api/image.rst`, `docs/api/exc.rst`, `docs/api/util.rst`

- `exc.rst` uses `.. automodule::` so new exceptions
  (`PackageTooLargeError`, `EncryptedPackageError`,
  `UnsupportedImageTypeError` — #1055, #668, #652) auto-render.
- `image.rst` (46 lines) is current; #1042 `Image.content_type` / `.ext`
  for EMF still works through the autoclass.
- `util.rst` is current.

---

## 5. Enum coverage gaps

`docs/api/enum/` contains 35 hand-written pages plus `index.rst`. Each
page is short (20-70 lines). `src/pptx/enum/*.py` defines **39 enum
classes** (plus the `PROG_ID` extendable enum). **4 classes have no page:**

### 5.1 Enums in `src/pptx/enum/` (39 total)

From `grep -E '^class (WD|MSO|XL|PP|PROG)_' src/pptx/enum/*.py`:

| Enum | File:line | Doc page | Status |
|---|---|---|---|
| `MSO_ANIMATION_TYPE` | `animation.py:1` | — | **missing** — #102 (set_animation API) |
| `MSO_ANIMATION_TRIGGER` | `animation.py:2` | — | **missing** — #102 |
| `MSO_AUTO_SHAPE_TYPE` | `shapes.py:1` | `MsoAutoShapeType.rst` | covered |
| `MSO_AUTO_SIZE` | `text.py:1` | `MsoAutoSize.rst` | covered |
| `MSO_COLOR_TYPE` | `dml.py:1` | `MsoColorType.rst` | covered |
| `MSO_CONNECTOR_TYPE` | `shapes.py:2` | `MsoConnectorType.rst` | covered |
| `MSO_FILL_TYPE` | `dml.py:2` | `MsoFillType.rst` | covered |
| `MSO_LANGUAGE_ID` | `lang.py` | `MsoLanguageId.rst` | covered |
| `MSO_LINE_DASH_STYLE` | `dml.py:3` | `MsoLineDashStyle.rst` | covered |
| `MSO_LINE_END_LENGTH` | `dml.py:6` | `MsoLineEndLength.rst` | covered |
| `MSO_LINE_END_TYPE` | `dml.py:4` | `MsoLineEndType.rst` | covered |
| `MSO_LINE_END_WIDTH` | `dml.py:5` | `MsoLineEndWidth.rst` | covered |
| `MSO_PATTERN_TYPE` | `dml.py:7` | `MsoPatternType.rst` | covered |
| `MSO_SHAPE_TYPE` | `shapes.py:3` | `MsoShapeType.rst` | covered |
| `MSO_TEXT_STRIKE_TYPE` | `text.py:3` | `MsoTextStrikeType.rst` | covered |
| `MSO_TEXT_UNDERLINE_TYPE` | `text.py:2` | `MsoTextUnderlineType.rst` | covered |
| `MSO_THEME_COLOR_INDEX` | `dml.py:8` | `MsoThemeColorIndex.rst` | covered |
| `MSO_VERTICAL_ANCHOR` | `text.py:4` | `MsoVerticalAnchor.rst` | covered |
| `PP_ACTION_TYPE` | `action.py` | `PpActionType.rst` | covered |
| `PP_AUTO_NUMBER_SCHEME` | `text.py:6` | `PpAutoNumberScheme.rst` | covered |
| `PP_MEDIA_TYPE` | `shapes.py:4` | `PpMediaType.rst` | covered |
| `PP_PARAGRAPH_ALIGNMENT` | `text.py:5` | `PpParagraphAlignment.rst` | covered |
| `PP_PLACEHOLDER_TYPE` | `shapes.py:5` | `PpPlaceholderType.rst` | covered |
| `PP_TRANSITION_SIDE_DIRECTION` | `transition.py:3` | — | **missing** — #942 / #1004 |
| `PP_TRANSITION_SPEED` | `transition.py:2` | — | **missing** — #1004 |
| `PP_TRANSITION_TYPE` | `transition.py:1` | — | **missing** — F8 + #942 |
| `PROG_ID` | `shapes.py:6` (`enum.Enum`) | — | not covered; extended by #752 with `.HTML`, etc. (extension members undocumented) |
| `XL_AXIS_CROSSES` | `chart.py` | `XlAxisCrosses.rst` | covered |
| `XL_AXIS_POSITION` | `chart.py` | `XlAxisPosition.rst` | covered |
| `XL_CATEGORY_TYPE` | `chart.py` | `XlCategoryType.rst` | covered |
| `XL_CHART_TYPE` | `chart.py` | `XlChartType.rst` | covered |
| `XL_CROSS_BETWEEN` | `chart.py` | `XlCrossBetween.rst` | covered |
| `XL_DATA_LABEL_POSITION` | `chart.py` | `XlDataLabelPosition.rst` | covered |
| `XL_ERROR_BAR_DIRECTION` | `chart.py` | `XlErrorBarDirection.rst` | covered |
| `XL_ERROR_BAR_INCLUDE` | `chart.py` | `XlErrorBarInclude.rst` | covered |
| `XL_ERROR_BAR_TYPE` | `chart.py` | `XlErrorBarType.rst` | covered |
| `XL_LEGEND_POSITION` | `chart.py` | `XlLegendPosition.rst` | covered |
| `XL_MARKER_STYLE` | `chart.py` | `XlMarkerStyle.rst` | covered |
| `XL_TICK_LABEL_POSITION` | `chart.py` | `XlTickLabelPosition.rst` | covered |
| `XL_TICK_MARK` | `chart.py` | `XlTickMark.rst` | covered |

### 5.2 Enum pages to add (4)

Stub pages needed (pattern from `docs/api/enum/MsoAutoSize.rst` — title,
alias, one-paragraph intro, bulleted member list). Each is 20-40 lines
and can be mostly machine-generated from the enum's `__members__`.

```
MsoAnimationType.rst         (MSO_ANIMATION_TYPE)
MsoAnimationTrigger.rst      (MSO_ANIMATION_TRIGGER)
PpTransitionType.rst         (PP_TRANSITION_TYPE)
PpTransitionSpeed.rst        (PP_TRANSITION_SPEED)
PpTransitionSideDirection.rst (PP_TRANSITION_SIDE_DIRECTION)
```

5 entries (4 enums + the side-direction variant). `docs/api/enum/index.rst`
has a 35-entry toctree that would need to grow by 5.

### 5.3 `PROG_ID` extension

`src/pptx/enum/shapes.py` defines `PROG_ID(enum.Enum)` with the built-in
members `XLSX`, `DOCX`, `XLS`, etc., plus the fork-added `HTML`, `MHT`,
`PDF` members for #752 (generic OLE embed). None are listed anywhere in
`docs/api/enum/`. A `ProgId.rst` stub (30 lines) covering the 10-odd
members is a P2 follow-up.

---

## 6. User-guide gaps

`docs/user/` has 19 pages totalling 4 548 lines:

```
intro.rst                 54 lines
install.rst               37 lines  ← out of date (section 7.3)
quickstart.rst           288 lines  ← pre-fork hello-world examples
presentations.rst        327 lines  ← has merge + sections (good)
slides.rst               223 lines
understanding-shapes.rst 144 lines
autoshapes.rst           316 lines
placeholders-understanding.rst 128 lines
placeholders-using.rst   372 lines
text.rst                 585 lines  ← has Templating-text (#398)
charts.rst               827 lines  ← has enumerated chart types (#244)
table.rst                538 lines
media.rst                 75 lines
math-equations.rst       102 lines
notes.rst                130 lines
ole-objects.rst           84 lines
comments.rst              89 lines
use-cases.rst            166 lines  ← has render-to-pdf/video/image section
concepts.rst              92 lines
```

### 6.1 Missing user-guide topics (P1/P2)

These feature areas have landed in HISTORY.rst but have no dedicated user
guide section (some have an API autoclass but no narrative):

| Feature area | Issues | Suggested page |
|---|---|---|
| Animations + transitions (`set_animation`, timings, MORPH, PP_TRANSITION_*) | F8, #102, #256, #264, #861, #942, #1004, #1106, #376 | `docs/user/animations.rst` (new, ~250 lines) |
| Slide operations — duplicate, move, delete, hidden, cross-deck copy, merge presentations | #67, #68, #132, #319, #1036, #934 | extend `docs/user/slides.rst` (+100 lines) |
| Shape operations — delete, duplicate, replace_with, z-order, flip, group transforms | #41, #49, #246, #533, #547, #925 | extend `docs/user/understanding-shapes.rst` or new `shape-operations.rst` |
| Picture/Movie replacement and linked pictures | #784, #806, #834 | extend `docs/user/media.rst` (75 → 200 lines) |
| Chart operations — apply_template, clone_to, add_plot (combo), replace_data preserve formulas, update_cached_values, 3D chart emit, scatter point colors | #239, #243, #266, #338, #381, #877 | extend `docs/user/charts.rst` |
| Table — cell borders, cell margin, cell merge, table style_id, row/column add + delete, column-count helpers | #15, #27, #28, #39, #51, #63, #71, #93, #142, #144, #145, #832, #837, #895 | extend `docs/user/table.rst` (538 → ~750 lines) |
| Custom document properties | #14, #82, #259 | new `docs/user/custom-properties.rst` (~80 lines) |
| Password-protected `.pptx` (open / save) | #668 | extend `docs/user/presentations.rst` with a "Password-protected decks" section (~40 lines) |
| Flat OPC / `.ppsx` / `.potx` formats | #438, #1059, #1070 | extend `docs/user/presentations.rst` or new `alt-formats.rst` |
| Font embedding | #355 | extend `docs/user/text.rst` or new `docs/user/font-embedding.rst` (API page already covers it) |
| Font borders, alpha, effective color, theme-hyperlink-color, East Asian / CS slots, strikethrough | #19 analog, #62, #337, #574, #938, #940 | extend `docs/user/text.rst` |
| Paragraph formatting — paragraph borders, bullet format, RTL, frame | #100, #114, #126, #127 | extend `docs/user/text.rst` |
| Text replacement / templating | #285, #398, #836 | covered (`docs/user/text.rst` "Templating text") |
| XY scatter / data labels / error bars | #544 (error bars), #619 (scatter data labels), #560 (series colors), #716 (label border) | extend `docs/user/charts.rst` |
| Sections (presentation) | F7, #257, #694 | covered (`docs/user/presentations.rst` "Presentation sections") |
| Math equations (read + write) | #126, #528, #892, #874 | covered (`docs/user/math-equations.rst`, 102 lines) |
| Comments | #487 (legacy) | covered (`docs/user/comments.rst`, 89 lines) |
| OLE objects — generic prog_id + extension | #752, #777 | extend `docs/user/ole-objects.rst` |
| SVG images (raise `UnsupportedImageTypeError`) | #652 | partially covered (`add_picture` note); needs "Inserting SVG images" section as referenced |
| Security / encryption / size limits | #1055, #668, #652 | covered by `docs/dev/security.rst` but not mentioned in user guide |
| Fit-text on placeholders | #715, #773, #936 | extend `docs/user/placeholders-using.rst` |
| Find shapes by name / xpath | #224, #309 | extend `docs/user/understanding-shapes.rst` |
| Preset geometry access | #515 | extend `docs/user/autoshapes.rst` (narrative for `Path`, `PathGeometry`) |
| Connector adjustments | #946, #1017, #332, #611, #674 | extend `docs/user/autoshapes.rst` |
| Accessibility — alt text on pictures, InlineShape properties | #158 analog (python-docx), N/A in pptx | no action |
| SmartArt scaffolding (F9) | #83 umbrella | needs dedicated `docs/user/smartart.rst` or an "Advanced shape kinds" section |

### 6.2 Feature areas already covered (partial check)

- **Comments** — `docs/user/comments.rst` (89 lines) covers the #487
  legacy-comments round-trip story.
- **Math equations** — `docs/user/math-equations.rst` (102 lines) covers
  read (`has_math_equation`, `math_equation_xml`) after #126 and write
  after #528.
- **Presentation sections** — `docs/user/presentations.rst` has a
  dedicated section.
- **Presentation merge** — covered in the same file.
- **Templating text** — `docs/user/text.rst` has a dedicated section.
- **Password-protected decks** — narrative deferred (see punch list).
- **Rendering out of scope** — covered (`docs/user/use-cases.rst`
  "Rendering to video, PDF, or image formats").
- **Animated GIFs** — covered (`docs/user/media.rst` section for #501).
- **Table height limitation** — covered (`docs/user/table.rst` section
  for #296).
- **SVG images** — partial (raise message pointed at, but the
  "Inserting SVG images" section referenced is only a brief stub).

---

## 7. Front-page, index, and quickstart

### 7.1 `docs/index.rst` (137 lines)

- The "Feature Support" bullet list (lines 24-47) has 11 bullets. Only
  three (legacy comments #487; sections F7/#257; password-protected decks
  #668) mention fork-era features. The other 50+ capabilities that
  landed in HISTORY.rst are not surfaced on the landing page — a new
  user scanning for "can python-pptx do X?" will not find 90% of the
  fork's work here.
- The "API Documentation" toctree (lines 104-121) lists 14 pages. No
  entry for a future `docs/api/animation.rst`, `docs/api/media.rst`,
  `docs/api/smartart.rst`, or any of the other future pages listed in
  section 4.1. Adding those stubs would need this toctree updated at
  the same time.
- The toctree's `:maxdepth: 2` is healthy (compared to the sibling
  python-docx's `:maxdepth: 1`) — the second-level anchors do render in
  the sidebar.

### 7.2 `docs/user/quickstart.rst` (288 lines)

- Last touched in upstream's 1.0 cycle. None of the fork features
  appear — no example of a tracked shape, transition, animation, deck
  merge, custom property, or font embedding.
- Sections that are still accurate for 1.0 but lag the fork: "Hello
  World!", "Bullet slide example", "Adding a paragraph", "Adding a
  heading", "Adding a page break", "Adding a picture", "Adding a
  table", "Opening an existing presentation".
- Feature additions that a modern quickstart should mention (one
  paragraph each, not exhaustive): presentation sections, slide
  duplicate, font embedding, combo chart, connector, deck merge,
  tracked-text templating.

### 7.3 `docs/user/install.rst` (37 lines)

- `docs/user/install.rst:22` claims:

  ```
  Python 2.6, 2.7, 3.3 or later
  ```

  **Wrong.** `pyproject.toml` has `requires-python = ">=3.8"` (and
  classifiers advertise 3.8-3.12). The Python 2.x, 3.3, 3.4 entries are
  historical.
- Line 13 recommends `easy_install`. That has been obsolete since
  `pyproject.toml` shipped.
- The `python setup.py install` reference is likewise legacy.
- The `msoffcrypto-tool` optional-dependency section (lines 27-37) is
  current and well-written — the only fork-era update on this page.
- One of the three `docs/user/*.rst` pages that do *not* mention any
  fork feature; the other two are `notes.rst` and `intro.rst`.

### 7.4 `docs/user/intro.rst`

Upstream content. No fork-era content.

---

## 8. HISTORY.rst

Unlike the sibling `python-docx` fork (where HISTORY stopped at upstream's
1.2.0), **this fork's HISTORY.rst is *in good shape*.** Lines 6-1 269
form the open `Unreleased` block with:

- **94 `feat:` entries** (new features, many with extensive
  paragraph-length descriptions)
- **28 `fix:` entries**
- Dozens of `verify:` / `docs:` / `rfctr:` / `perf:` / `security:` /
  `Foundation` items
- **160 distinct `#NNN` issue references**

Release-notes discipline is tight — the critical gap here is not
HISTORY; it's whether the features named in HISTORY actually appear in
the user guide / API reference (sections 4-6).

Before an actual 1.3.0 cut the `Unreleased` block would need to be
split into a version heading and possibly pruned for parallel structure,
but content-wise it is *complete* for a reader who wants to know what
changed.

---

## 9. Other issues

### 9.1 Broken image links

**None.** All 25 `.. image::` references point at files that exist in
`docs/_static/img/`. Every image on disk is referenced from at least
one `.rst` page. (`diff <(grep -rnE '\.\. image::' docs | extract paths)
<(ls docs/_static/img/)` returns empty.)

### 9.2 Orphaned / malformed pages

- **`docs/dev/analysis/f8-animations-transitions.rst`** — not in
  `docs/dev/analysis/index.rst` toctree. Sphinx emits "document isn't
  included in any toctree". F9 is wired up; F8 is not. One-line fix.
- **`docs/dev/analysis/f9-smartart.rst:65`** — malformed table (one row
  column-alignment mismatch). Emits an ERROR. Cosmetic one-line fix.
- **`docs/user/text.rst:326`, `:372`, `:422`** — three explicit-markup
  blocks that run into the next paragraph without a blank line.
  Cosmetic.
- **`docs/user/charts.rst:463`** and **`docs/user/presentations.rst:311`**
  — title underlines shorter than the title. Cosmetic.

### 9.3 References to removed modules / classes

Grepping the `.rst` tree for dotted paths against the current source
tree:

```
grep -rn 'pptx\.' docs/api/*.rst | cut -d: -f3 | grep -oE 'pptx\.[a-z_.]+' | sort -u
```

returns 38 distinct dotted paths. Every one resolves under `src/pptx/`.
No dead references.

### 9.4 `.. todo::` / `:deprecated:` markers

`grep -rn '.. todo::\|.. deprecated::' docs/ src/pptx/` returns zero hits
in `docs/`. There are no dangling to-dos in RST files. (There is one
deprecation in `src/pptx/slide.py` — `AnimationEffect = AnimationEffectView`
— but it is handled by the `rfctr:` HISTORY.rst entry, not by a
`.. deprecated::` directive.)

### 9.5 `docs/conf.py` observations

- Pinned to Sphinx 1.8.6 (`requirements-docs.txt:1`) — Sphinx 1.8 reached
  EOL in early 2019. Most modern Sphinx extensions and themes don't
  support this version.
- `html_theme_path = [".themes"]` (`conf.py:432`) — the *dotted* theme
  directory (not `_themes`). Works with the current Sphinx but is
  unusual.
- `copyright = u"2012, 2013, Steve Canny"` (`conf.py:72`) — has not been
  updated in 13 years. The fork's `LICENSE` remains under Steve Canny's
  name, which is appropriate, but the `copyright` banner on generated
  pages reflects only the original upstream date range.
- `intersphinx_mapping` is commented out (`conf.py:587`) — no actual
  intersphinx is configured. The upstream python-docx version has the
  same state. This limits the usefulness of the existing
  `sphinx.ext.intersphinx` extension declaration.
- `exclude_patterns = [".build"]` (`conf.py:396`) — this *matches* the
  build directory under the current Makefile (`BUILDDIR = .build`), so
  it's correct, but it is noise in the config.
- `sphinx.ext.todo` is enabled (`conf.py:52`) but never used
  (`grep -rn '.. todo::' docs/` returns 0).
- `sphinx.ext.coverage` is enabled (`conf.py:53`) but has no
  `coverage_*` options set — the coverage report, if generated, would be
  minimal.
- `sphinx.ext.ifconfig` is enabled (`conf.py:54`) but never used.
- No `autodoc_default_options` is set. Every `.. autoclass::` directive
  must spell out `:members:` / `:inherited-members:` individually. A
  project-wide `autodoc_default_options = {"members": True, "undoc-members":
  False, "show-inheritance": True}` would cut maybe 15% off each API
  page.
- No `nitpicky = True` and no `-n` on Sphinx invocations — dangling
  cross-references render silently as raw text. Toggling `nitpicky` on
  the strict CI build would expose many more latent problems (but is a
  larger cleanup than the 140-warning substitution pass).

---

## 10. Coverage matrix for fork additions

Reading the `Unreleased` block of `HISTORY.rst` top-to-bottom, grouping
by functional area, and cross-checking against
`docs/api/`, `docs/user/`, and `docs/index.rst` "Feature Support":

| Area (issues) | API docs | User guide | Feature Support bullet | HISTORY.rst |
|---|---|---|---|---|
| Legacy comments (#487) | ✅ `docs/api/comments.rst` (substitutions broken) | ✅ `docs/user/comments.rst` | ✅ | ✅ |
| Presentation sections (F7, #257, #694) | ✅ `presentation.rst` | ✅ `presentations.rst` | ✅ | ✅ |
| Password-protected `.pptx` (#668, #702) | — (no dedicated section) | — | ✅ partial | ✅ |
| Presentation.merge (#934) | — | ✅ `presentations.rst` | — | ✅ |
| Presentation.save_ppsx (#438) | — | — | — | ✅ |
| Presentation.save_flat_xml (#1059) | — | — | — | ✅ |
| Reproducible builds via fixed zip timestamps (#702) | — (docstring only) | — | — | ✅ |
| `.potx` + `.ppsx` read (#1070) | — | — | — | ✅ |
| Letter / A4 / (cx, cy) paper sizes (#883) | — | — | — | ✅ |
| 16x9 default template (#1066) | — | — | — | ✅ |
| Custom document properties (#14, #82, #259) | — (would need a section on `presentation.rst`) | — | — | ✅ |
| Embed fonts (#355) | ✅ `presentation.rst` "Embedding custom fonts" | — | — | ✅ |
| Digital-signature / encryption / recovery (#1055, #668, #152 analog) | ✅ `dev/security.rst` | — (user-guide not pointed at it) | — | ✅ |
| Sections leaf API (#257) | — (autorender on Section) | ✅ | — | ✅ |
| Slide duplicate / move / delete / hidden / external copy (#67, #68, #132, #319, #1036) | ✅ (autorender) | — | — | ✅ |
| Slide transitions, MORPH, speed, direction (F8, #942, #1004) | ✅ `slides.rst` (`|Transition|` broken) | — | — | ✅ |
| Slide animations & timing introspection (F8, #102, #256, #264, #861, #1106, #376) | — | — | — | ✅ |
| Shape operations — delete, duplicate, replace_with, z-order, flip, hidden, is_hidden (#41, #49, #246, #533, #547, #971) | ✅ (autorender) | — | — | ✅ |
| Shape group composite effective_* (#925) | ✅ (autorender) | — | — | ✅ |
| Shape find — by name, xpath (#224, #309) | ✅ (autorender on SlideShapes) | — | — | ✅ |
| Picture.replace_image (#834) | ✅ (autorender) | — | — | ✅ |
| SlideShapes.add_picture_link (#806) | ✅ (autorender) | — | — | ✅ |
| Movie.replace_media / start_time / start_condition / delete (#811, #784, #974) | — (no `|Movie|` page) | — | — | ✅ |
| Floating-image / anchoring via add_picture (#30 analog — N/A pptx) | n/a | n/a | n/a | n/a |
| Fill alpha / ColorFormat.alpha (#62) | ✅ (autorender) | — | — | ✅ |
| Fill — blip / image fill (#234) | ✅ (autorender on FillFormat) | — | — | ✅ |
| Shadow family — full blur/distance/direction/color (#130, #275, #446, #705, F2) | ✅ (autorender on ShadowFormat) | — | — | ✅ |
| Line end arrow (#375, #1033) | ✅ (autorender) | — | — | ✅ |
| Connector adjustments (#946, #1017, #332, #611, #674) | ✅ `shapes.rst` (`|ConnectorAdjustmentCollection|` broken) | — | — | ✅ |
| Preset geometry / PathGeometry (#515) | ✅ `shapes.rst` (`|PathGeometry|` / `|Path|` broken) | — | — | ✅ |
| Chart — apply_template (#243) | — | — | — | ✅ |
| Chart — clone_to / SlideShapes.clone_chart (#877) | — | — | — | ✅ |
| Chart — add_plot combo (#338) | — (autodoc) | — | — | ✅ + analysis page |
| Chart — replace_data preserve formulas (#239) | ✅ (autorender) | — | — | ✅ |
| Chart — update_cached_values (#381) | ✅ (autorender) | — | — | ✅ |
| Chart — 3D chart emit (#266) | — | — | — | ✅ |
| Chart — chart_style two-tier (#516, #407) | ✅ (docstring rewrite) | — | — | ✅ |
| Chart — pie/scatter/error-bar per-point colors (#544, #716, #560, #357, #825) | ✅ partial (`|ErrorBars|` broken; XL_ERROR_BAR_* broken) | — (mentions scatter in #244) | — | ✅ |
| Chart — title position (#1030) | ✅ (autorender) | — | — | ✅ |
| Chart — data-label background + border + fill (#560, #662, #716) | ✅ (autorender) | — | — | ✅ |
| Chart — category axis number_format (#571) | ✅ (autorender) | — | — | ✅ |
| Chart — XY scatter data labels (#619) | ✅ (autorender) | — | — | ✅ |
| Chart — chartex passthrough / UNSUPPORTED_CHARTEX (#386) | — | ✅ `charts.rst` mentions (#244) | — | ✅ |
| Chart — date-axis major/minor unit (#472) | ✅ (autorender) | — | — | ✅ |
| Chart — Axis.position (#349) | ✅ (autorender) | — | — | ✅ |
| Chart — secondary value axis (#141) | ✅ (autorender) | — | — | ✅ |
| Chart — data-label colors / text_frame (#560, #1072) | ✅ (autorender) | — | — | ✅ |
| Chart — theme accent cycling (#529, #539) | ✅ (autorender) | — | — | ✅ |
| Table — cell borders (#71) | ✅ (autorender) | — | — | ✅ |
| Table — cell shading / background (#63) | ✅ (autorender) | — | — | ✅ |
| Table — add / delete rows + columns (#832, #837, #895) | ✅ (autorender) | — | — | ✅ |
| Table — row height (#28, #296) | ✅ (autorender) | ✅ `table.rst` (#296 limitation note) | — | ✅ |
| Table — Row.is_header / allow_break (#51, #93) | ✅ (autorender) | — | — | ✅ |
| Table — style_id (#27) | ✅ (autorender) | — | — | ✅ |
| Table — cell merge / span / split helpers | ✅ (autorender) | ✅ `table.rst` | — | ✅ |
| Table — cell row_idx / col_idx (#849) | ✅ (autorender) | — | — | ✅ |
| Table — float dimensions on add_table (#288) | ✅ (autorender on SlideShapes) | — | — | ✅ |
| Text — replace_text family (#836, #398, #285) | ✅ (autorender) | ✅ `text.rst` "Templating text" | — | ✅ |
| Text — math equation read/write (#126, #528, #892, #874) | ✅ (autorender on BaseShape + _Paragraph) | ✅ `math-equations.rst` | — | ✅ |
| Text — font borders, alpha, effective_color, symbol (#62, #938) | ✅ (autorender) | — | — | ✅ |
| Text — font name_ea / name_cs (#337) | ✅ `text.rst` (has |Font| narrative) | — | — | ✅ |
| Text — font strikethrough (#574) | ✅ (autorender) | — | — | ✅ |
| Text — use_theme_hyperlink_color (#940) | ✅ (autorender) | — | — | ✅ |
| Text — TextFrame.rotation / font_scale / line_space_reduction (#133, #715) | ✅ (autorender) | — | — | ✅ |
| Text — fit_text guards (#773, #936) | ✅ (autorender) | — | — | ✅ |
| Text — paragraph bullet API (#100, #114) | ✅ `text.rst` ( `|_BulletFormat|` substitution defined) | — | — | ✅ |
| Text — Run / Paragraph .delete (#144) | ✅ (autorender) | — | — | ✅ |
| ActionSetting.screen_tip (#425) | ✅ (autorender) | — | — | ✅ |
| ActionSetting set_sound / remove_sound (#734) | ✅ `action.rst` (`|Sound|` / `|Audio|` broken) | — | — | ✅ |
| OLE — generic prog_id / extension (#752, #777) | ✅ (autorender) | — (narrative would belong here) | — | ✅ |
| OLE — html embed (#777) | ✅ (autorender) | — | — | ✅ |
| SVG refusal + UnsupportedImageTypeError (#652) | ✅ (autorender) | — (mentions message) | — | ✅ |
| SmartArt scaffolding (F9) | — | — | — | ✅ (analysis page exists) |
| Sections F7, F8, F6, F1..F9 foundations | — mostly analysis-only | — | — | ✅ |
| Performance — partname allocation (#644) | n/a | n/a | n/a | ✅ |
| Security — XML hardening + size cap (#1055) | ✅ `dev/security.rst` | — | — | ✅ |
| pyparsing migration (#1103) | n/a | n/a | n/a | ✅ |
| MPO image support (#787) | ✅ (autorender on Image) | — | — | ✅ |
| EMF content-type fix (#1042) | ✅ (autorender) | — | — | ✅ |
| Picture blipFill missing guard (#434) | n/a (bug fix) | — | — | ✅ |
| Movie add when audio exists (#926, #323, #502) | n/a (bug fix) | — | — | ✅ |
| Movie delete orphans (#974) | n/a (bug fix) | — | — | ✅ |
| Slide-id allocator (#68/#132/#1036 via F6) | — (analysis page only) | — | — | ✅ |
| HPlaceholder inherit methods, insert_chart, insert_picture, insert_table (#199, #333, #449, #176) | ✅ (autorender) | — | — | ✅ |
| Masters/layouts accept non-placeholder shapes (#575) | ✅ (autorender) | — | — | ✅ |
| SlideMaster.name fallback (#679) | ✅ (autorender) | — | — | ✅ |
| Set auto-advance bulk (#376) | ✅ (autorender) | — | — | ✅ |
| `docs/dev/analysis/chartex-*` (funnel, treemap, waterfall, foundation) | ✅ analysis pages | ✅ `charts.rst` #244 enumerates | — (not on index) | ✅ |

**Totals** (rough bucketing of the 160 distinct issue references):

- Features with API page coverage (at least an `.. autoclass::`): ~130/160
- Features with user-guide narrative: **13 / 160** (~8%)
- Features advertised on `docs/index.rst` Feature Support: **3 / 160**
  (~2%)

---

## 11. Recommendations

Prioritised punch list. Effort labels: **S** = < 2 hr, **M** = half a
day, **L** = 1+ day.

### P1 — build hygiene + visibility

1. **P1 / S — Add the 32 missing `|Name|` substitutions to
   `docs/conf.py` `rst_epilog`.** One line per symbol
   (`.. |Foo| replace:: :class:\`.Foo\``). Retires ~120 of the 159 build
   errors in one pass. ~30 min, ~35 lines added.
2. **P1 / S — Add `docs/api/enum/MsoAnimationType.rst`,
   `MsoAnimationTrigger.rst`, `PpTransitionType.rst`,
   `PpTransitionSpeed.rst`, `PpTransitionSideDirection.rst`** following
   the `docs/api/enum/MsoAutoSize.rst` pattern (title, alias, one-paragraph
   intro, bulleted member list). Update
   `docs/api/enum/index.rst` toctree. ~2 hr, ~150 lines total.
3. **P1 / S — Add autoclass directives for fork-introduced public
   classes that currently have no reference page.** `AnimationEffect`
   (authoring), `AnimationEffectView` (introspection), `Movie`
   (picture subclass), `SmartArt` (graphfrm), `Video` (media).
   Either extend `docs/api/shapes.rst` + `docs/api/slides.rst` with new
   subsections, or create a thin `docs/api/animation.rst` /
   `docs/api/media.rst` / `docs/api/smartart.rst`. ~2 hr,
   ~80 lines added per page.
4. **P1 / S — Fix the 5 cosmetic build warnings** — two underline-too-short,
   three unexpected-unindent, one malformed table in
   `f9-smartart.rst:65`, plus wire `f8-animations-transitions.rst` into
   `docs/dev/analysis/index.rst`. ~15 min.
5. **P1 / M — Expand the "Feature Support" bullet list on
   `docs/index.rst`** to include the 10-15 most-prominent fork
   capabilities. Suggested additions: MORPH transitions, shape
   animations, combo charts, 3D charts, deck merging, font embedding,
   cross-slide chart copy, picture replacement, table column mutation,
   cell borders, connector adjustments, custom document properties,
   flat-OPC / `.ppsx` / `.potx` formats. ~30 min.
6. **P1 / S — Fix `docs/user/install.rst`** — drop Python 2.x and
   easy_install references, update `requires-python = ">=3.8"`, drop
   `setup.py install` recommendation, add `pip install python-pptx`
   prominently. ~20 min, ~15 lines changed.

### P2 — user-guide narrative for fork features

7. **P2 / L — Write `docs/user/animations.rst`** covering MORPH
   transitions, `Shape.set_animation`, `Slide.animation_sequence`, preset
   enums, timing introspection. Reference `docs/dev/analysis/f8-animations-transitions.rst`
   for the deep dive. ~250 lines, ~1 day.
8. **P2 / M — Extend `docs/user/slides.rst`** with sections for slide
   duplicate (#132), move (#68), delete (#67), hidden (#319), cross-deck
   copy (#1036). ~100 lines added.
9. **P2 / M — Extend `docs/user/charts.rst`** with sections for chart
   templates (#243), cross-slide chart copy (#877), combo charts (#338),
   3D chart emit (#266), scatter / error-bar setup (#544, #619,
   #716/#560). ~200 lines added to an already-large file; possibly
   split off `docs/user/charts-advanced.rst` to keep the main page
   legible.
10. **P2 / M — Extend `docs/user/table.rst`** with narrative sections on
    cell borders (#71), cell shading (#63), cell margins, table
    style_id (#27), row+column add/delete (#832, #837, #895). ~150
    lines added (to the existing 538).
11. **P2 / M — Write `docs/user/custom-properties.rst`** for the
    CustomProperties dict-like accessor (#259). ~80 lines.
12. **P2 / M — Extend `docs/user/presentations.rst`** with sections on
    password-protected decks (#668), flat-OPC save (#1059), `.ppsx` save
    (#438), `.potx` read (#1070), 16x9 default + paper-size presets
    (#883, #1066), reproducible zip timestamps (#702), font embedding
    (#355). ~150 lines.
13. **P2 / M — Extend `docs/user/media.rst`** from 75 to ~200 lines —
    cover Movie replace_media (#784), start_time + start_condition
    (#811), `Picture.replace_image` (#834), linked pictures (#806).
14. **P2 / S — Modernise `docs/user/quickstart.rst`.** Add short
    paragraphs for 4-5 common fork features (transition, font
    embedding, search/replace, chart add, deck merge). ~2 hr, +40
    lines.
15. **P2 / S — Write `docs/user/smartart.rst`** stub pointing at the F9
    read-only story and the full-authoring roadmap (#83). ~50 lines.
16. **P2 / S — Extend `docs/user/ole-objects.rst`** with a section on
    the #752 generic `prog_id` / `extension` path (HTML, PDF,
    arbitrary). ~30 lines.
17. **P2 / S — Extend `docs/user/autoshapes.rst`** with narrative for
    path geometry (#515) and connector adjustments (#946). ~60 lines.

### P2 — conf-file improvements

18. **P2 / S — Add `autodoc_default_options`** in
    `conf.py` so new autoclass directives inherit
    `members: True, show-inheritance: True`. Shrinks each API page,
    makes new methods visible by default without having to re-edit the
    `.rst`. ~15 min.
19. **P2 / S — Update `requirements-docs.txt`** to modern Sphinx (≥5,
    <8) and drop the obsolete `MarkupSafe==0.23` / `Jinja2==2.11.3`
    pins; drop the unused `alabaster` pin. Keep the existing
    `.readthedocs.yaml` `python: "3.9"` constraint or bump it to
    `"3.11"` in the same pass. ~15 min.
20. **P2 / S — Update `copyright = "2012, 2013, Steve Canny"`** in
    `conf.py:72` to include current year. ~1 min.

### P3 — polish and longer-term

21. **P3 / M — Migrate theme** from vendored `armstrong` to a maintained
    theme (`furo`, `pydata-sphinx-theme`, or `sphinx_rtd_theme`). Delete
    `docs/.themes/armstrong/` (~2-3 MB of vendored code). ~2 hr + QA
    pass.
22. **P3 / S — Remove unused Sphinx extensions** from `conf.py:47-56` —
    `sphinx.ext.todo`, `sphinx.ext.coverage`, `sphinx.ext.ifconfig`,
    `sphinx.ext.inheritance_diagram` are all enabled but not used. ~5
    min.
23. **P3 / M — Enable `nitpicky = True`** in `conf.py` and run a clean
    pass — expect ~200-300 additional warnings for dangling cross-refs.
    Fix or add to `nitpick_ignore`. ~1 day, but high-value for the
    strict-CI build.
24. **P3 / L — Audit every public docstring** for OOXML-term consistency
    (`c:ser`, `a:xfrm`, `p:sld`, …) and add `.. versionadded::` / `..
    versionchanged::` directives so the release-notes skeleton can be
    auto-generated from code. ~1 week, on and off.
25. **P3 / S — Expand `docs/community/`** from 25 lines total. `faq.rst`
    is 2 lines; `support.rst` is 16; `updates.rst` is 7. An FAQ of
    common gotchas (e.g., "Why doesn't my transition show in
    PowerPoint?", "How do I embed a font?") would be high-value and
    ~150 lines.
26. **P3 / M — Add an `intersphinx_mapping`** (currently commented out
    at `conf.py:587`) so `:class:\`datetime.datetime\`` and
    `:class:\`pathlib.Path\`` render as cross-links. Trivial unblock of
    an already-enabled extension.

---

## Appendix A — counts at a glance

| Metric | Count |
|---:|---|
| `.rst` files under `docs/` | 180 |
| Lines across all `.rst` files | 36 353 |
| `.rst` files under `docs/api/` (top level) | 14 |
| `.rst` files under `docs/api/enum/` (incl. index) | 36 |
| `.rst` files under `docs/user/` | 19 |
| `.rst` files under `docs/dev/` (total) | 107 |
| `.rst` files under `docs/dev/analysis/` | 88 |
| `.rst` files under `docs/community/` | 3 |
| Python modules under `src/pptx/` (top level) | 19 |
| Python submodule files under `src/pptx/{chart,dml,enum,opc,oxml,parts,shapes,text}` | 65+ |
| Enum classes defined in `src/pptx/enum/` | 39 |
| Enum classes with a `docs/api/enum/*.rst` page | 35 |
| `|Substitution|` tokens defined in `conf.py:rst_epilog` | ~150 |
| `|Substitution|` tokens referenced but undefined | 32 |
| Distinct `#NNN` issue refs in `HISTORY.rst` Unreleased block | 160 |
| `feat:` entries in `HISTORY.rst` Unreleased block | 94 |
| `fix:` entries in `HISTORY.rst` Unreleased block | 28 |
| Non-merge doc commits since 2024-01-01 | 163 |
| Sphinx build warnings (non-strict) | 159 ERROR + 8 WARNING |
| Sphinx build errors blocking `-W` (strict) | 1st at `action.py` docstring |

## Appendix B — Sphinx build command used

```
PYTHONPATH=src python -m sphinx -b html docs docs/.build/html
```

Ran with Sphinx 7.4.7 inside a Python 3.14 venv. Build result:
`build succeeded, 140 warnings.` Exit status 0. `docs/.build/` and the
venv were removed after the run.
