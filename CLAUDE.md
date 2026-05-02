# CLAUDE.md

Guidance for Claude Code (and other AI assistants) working in this repository.

---

## 1. Project summary

**python-pptx** is a mature, MIT-licensed Python library for creating, reading, and updating PowerPoint (`.pptx`) files. It parses and emits Office Open XML via `lxml`, does not require PowerPoint to be installed, and is intended to be **industrial-grade** — suitable for commercial use, which demands correctness and round-trip fidelity.

- Upstream: `scanny/python-pptx` (tracked by the `upstream` git remote)
- Active release line: v1.0.x
- Python support: ≥3.8 (classifiers list 3.8–3.12)
- Runtime deps: `Pillow`, `XlsxWriter`, `lxml`, `typing_extensions`

### Sibling projects

python-pptx is part of a family of Python libraries for reading/writing Office Open XML formats. Each targets a different Office application but shares the same design philosophy (lxml-backed, no Office install required, round-trip fidelity, src-layout + strict tooling):

- **python-docx** — Microsoft Word (`.docx`) documents. `scanny/python-docx`.
- **python-pptx** — Microsoft PowerPoint (`.pptx`) presentations. This repo.
- **python-xlsx** — Microsoft Excel (`.xlsx`) workbooks. `scanny/python-xlsx`.

Patterns and idioms found in this repo often have direct analogues in the siblings — e.g., the `xmlchemy` descriptor layer, part/relationship graph, content-type registration, `get_or_add_*` oxml helpers, and the `Describe*`/`it_*` test-naming convention. Cross-pollinate from sibling repos when a problem here has already been solved there.

---

## 2. Repository layout

```
src/pptx/           # library source (src-layout)
  api.py              top-level Presentation() factory
  presentation.py     Presentation class
  slide.py            Slide, SlideLayout, SlideMaster, Slides, …
  package.py          OPC package assembly
  util.py             Length, Emu, Pt, Inches, Cm, Mm, …
  chart/              charts (axis, series, plot, data, xmlwriter, …)
  dml/                DrawingML (fill, line, color, effect)
  enum/               all public enumerations (XL_CHART_TYPE, MSO_SHAPE, …)
  opc/                Open Packaging Conventions (parts, relationships, zip)
  oxml/               lxml custom element classes (the XML layer)
  parts/              package parts (slide, chart, image, media, …)
  shapes/             shape tree, autoshape, picture, connector, table, group
  text/               text frames, runs, paragraphs, fonts
  templates/          default.pptx and friends shipped with the package
  py.typed            PEP 561 marker (don't delete)

tests/              pytest unit tests (mirrors src/pptx/ layout)
features/           behave acceptance tests (.feature + steps/)
docs/               Sphinx documentation (user, api, dev, community)
spec/               ad-hoc OOXML discovery notes (excluded from lint)
lab/                experimental / throwaway scripts (excluded from lint)
typings/            custom type stubs (mainly for lxml)
HISTORY.rst         release-history changelog (user-visible)
pyproject.toml      build + tool config
Makefile            convenience targets: accept, docs, coverage, build
tox.ini             py38–py312 test envs
```

The `lab/` and `spec/` directories are **intentionally undisciplined**. Ruff and pyright exclude them. Do not "clean them up" or reformat their contents; they're scratch space and that's a feature.

---

## 3. Remotes and branching

- `upstream` → `https://github.com/scanny/python-pptx.git` — do **not** push here.
- `origin` → the developer's fork — push all work here.
- Default branch is `master`.

When contributing a PR upstream, branch from `master`, push to `origin`, and open the PR against `upstream/master`.

---

## 4. Tooling and style

All three tools are strict. Respect them — don't disable or silence lints/types to make a patch land.

| Tool | Config | Notes |
|---|---|---|
| **Black** | `pyproject.toml` → `[tool.black]` | line length 100 |
| **Ruff** | `pyproject.toml` → `[tool.ruff]` | line length 100; lint rule set includes `C4`, `COM`, `E/F/W`, `I`, `PT`, `SIM`, `TCH001`, select `UP` rules; `isort` first-party is `pptx`; `lab/`, `spec/`, `docs/`, `ref/` excluded |
| **Pyright** | `pyproject.toml` → `[tool.pyright]` | `typeCheckingMode = "strict"`, `pythonVersion = "3.9"`, `reportUnnecessaryTypeIgnoreComment = true`, `reportUnnecessaryCast = true`, custom stubs under `typings/` |

**Style rules to follow when editing:**
- Keep docstrings concise; this project uses short imperative-voice docstrings. No multi-paragraph essays.
- Prefer editing existing modules over adding new ones. If a new module is genuinely needed, mirror the existing layout.
- Public API additions need to be exported through the relevant `__init__.py` and referenced from the API docs page (see §7).
- If `pyright --strict` flags something, **fix the types**. Don't reach for `# type: ignore` as a first resort. If you must, suppress the specific rule (e.g., `# pyright: ignore[reportPrivateUsage]`) and include a one-line reason.
- `from __future__ import annotations` is used throughout — keep it.
- Imports follow isort ordering (ruff `I` rule). Keep first-party `pptx` imports in the dedicated group.

---

## 5. Testing

There are **two** test suites and both must pass for any feature/fix:

### Unit tests (pytest) — `tests/`

- Layout mirrors `src/pptx/` (e.g., `src/pptx/chart/axis.py` ↔ `tests/chart/test_axis.py`).
- **Test-function naming** (configured in `pyproject.toml`):
  - classes: `Test…` or `Describe…`
  - functions: `test_…`, `it_…`, `they_…`, `but_…`, `and_…`
- **Warnings are errors** (`filterwarnings = ["error"]`). A new `DeprecationWarning` anywhere will fail the suite — fix the cause.
- Test fixtures for XML are commonly built with the helpers in `tests/unitdata.py` and per-subpackage `unitdata` modules.
- Private-access suppression: some test files start with `# pyright: reportPrivateUsage=false` — fine, since unit tests exercise internals.

Run:
```bash
pytest                # entire suite
pytest tests/chart    # subtree
pytest -k <pattern>   # filter by test name
make coverage         # pytest with coverage
```

### Acceptance tests (behave) — `features/`

- Each capability area has a `*.feature` file in Gherkin, plus a matching step module under `features/steps/`.
- Naming convention uses a short area prefix: `cht-*` (charts), `dml-*` (DrawingML fill/line/effect/color), `shp-*` (shapes), `txt-*` (text), `tbl-*` (tables), `ph-*` (placeholders), `prs-*` (presentation-level), `act-*` (actions/hyperlinks), etc.
- `features/environment.py` creates a `_scratch/` dir for generated pptx output.

Run:
```bash
make accept           # behave --stop
behave features/cht-chart.feature    # single file
behave --tags=-wip                   # skip work-in-progress
```

### Full CI-equivalent local run

```bash
tox              # py38..py312 in parallel; runs pytest + behave per env
```

---

## 6. The strict feature-work rule

**Every user-visible feature change requires updates in lockstep across four places:**

1. **Implementation** under `src/pptx/`.
2. **Unit tests** under `tests/` — new tests exercising the new code paths, covering both the happy path and failure modes. The test file path should mirror the source file path.
3. **Acceptance tests** under `features/` — a new `Scenario` (or a new `.feature` file for a new area) plus its step implementation. BDD scenarios document the capability from the user's perspective; they must execute the real API, not mocks.
4. **Documentation** under `docs/`:
   - **User guide** (`docs/user/…`) if the feature is part of the public API — add a section or code example showing how to use it.
   - **API reference** (`docs/api/…`) for every new class, method, or property — at minimum a `.. autoclass::` or `.. automethod::` directive in the right page.
   - **Feature Support** bullet in `docs/index.rst` for a capability-level addition (new shape kind, new chart type, new round-trip area, etc.).
5. **`HISTORY.rst`** — add a one-line entry for any user-visible change (new feature, fix of a user-reported bug, API change) under the pending release heading. Use the existing style (`- fix: #NNN …` for bugs, `- Add #NNN …` for features, or a short description for other items).

"N/A" is acceptable for items genuinely not applicable — e.g., a pure internal refactor with no public-surface change may legitimately skip docs and HISTORY — but the default assumption is **all five apply**. If you skip an item, state why in the PR description.

For pure bug fixes, items 3 (acceptance) and the user-guide portion of 4 may be skipped if the bug is purely internal; a unit test and a HISTORY line are still required.

---

## 7. Documentation build

```bash
make docs        # Sphinx HTML under docs/.build/html/
make cleandocs   # nuke build cache
make opendocs    # open built docs in browser
```

`docs/` layout:
- `docs/user/` — tutorial/conceptual guide pages
- `docs/api/` — API reference (one file per module/area)
- `docs/dev/` — contributor docs (`development_practices.rst`, `philosophy.rst`, `xmlchemy.rst`, etc.)
- `docs/community/` — support, FAQ, updates
- `docs/index.rst` — top-level page; includes the "Feature Support" bullet list that must be updated when capability is added

Sphinx config is in `docs/conf.py`.

---

## 8. Common workflows

### Adding a new public method on an existing class
1. Implement in the appropriate `src/pptx/…` module.
2. Add unit tests in the mirrored test file.
3. Add a behave scenario to the relevant `*.feature` file and its step function.
4. Add `.. automethod::` to the corresponding `docs/api/…` page.
5. Update `docs/user/…` if end users should know about it.
6. Add a `HISTORY.rst` line.

### Adding a new enum value
- Enums live in `src/pptx/enum/`. They use a custom metaclass; read a neighboring enum first to see the pattern (in particular, the "return value" XML mapping).
- Update the enum's doc in `docs/api/enum/` if present.

### Adding a new XML element class
- Custom element classes live in `src/pptx/oxml/…`. Read `docs/dev/xmlchemy.rst` first — it explains `ZeroOrOne`, `OneAndOnlyOne`, `ZeroOrMore`, `RequiredAttribute`, etc.
- Register the new element with the `ns` registry (see top of `src/pptx/oxml/__init__.py`).

### Release prep (reference only — ask before running)
- See `docs/dev/development_practices.rst`.
- Update `src/pptx/__init__.py` version, `HISTORY.rst`, confirm docs compile, full tox run clean, build (`make build`), upload (`make upload`).

---

## 9. OOXML / XML tips

- `oxml/` uses a home-grown descriptor layer (`xmlchemy`) on top of `lxml.etree`. Don't bypass it with raw `etree` access in production code — the descriptors carry namespace/type/default semantics.
- The `ns` helpers (`qn("a:solidFill")`, `nsmap`, `nsdecls`) are how namespaces are handled throughout — use them consistently.

### Consulting `spec/` before implementing a new OOXML element

`spec/` is an offline archive of the ECMA-376 / ISO/IEC 29500 specification. It is **not** imported or executed by the library; it's the reference-reading room. The subdirectories:

- `spec/ISO-IEC-29500-1/` — Part 1 ("Fundamentals & Markup Language Reference"). The `schemas/xsd/` tree is the authoritative grammar for every DrawingML (`dml-*.xsd`), PresentationML (`pml.xsd`), SpreadsheetML (`sml.xsd`), and shared-type (`shared-*.xsd`) element the library models. `schemas/dml-geometries/` is the preset-geometry reference for autoshapes.
- `spec/ISO-IEC-29500-2/` — Part 2 (Open Packaging Conventions). Read `opc-xsd/` for relationship / content-type / package grammar.
- `spec/ISO-IEC-29500-3/` — Part 3 (Markup Compatibility). Defines `mc:AlternateContent` / `mc:Choice` / `mc:Fallback`, which wrap most Office 2010+ extended features.
- `spec/ISO-IEC-29500-4/` — Part 4 (Transitional conformance). Relevant when round-trip fidelity diverges between the "strict" and "transitional" variants.
- `spec/gen_spec/` — one-off maintainer tooling used to seed `src/pptx/spec.py`. Do not try to run it; treat it as historical.

**Workflow before writing code for a new element or feature:**

1. **Produce a real PowerPoint sample.** Create a minimal `.pptx` in Microsoft PowerPoint that exercises the feature, unzip it, and read the XML. **This is ground truth** — what the library must emit to round-trip cleanly.
2. **Look up the grammar in `spec/ISO-IEC-29500-1/schemas/xsd/`** (or Part 2's `opc-xsd/` for packaging work). This gives the formal parent/child relationships, attribute types, defaults, and cardinality — which map directly onto `ZeroOrOne` / `OneAndOnlyOne` / `ZeroOrMore` / `RequiredAttribute` descriptors.
3. **Reconcile steps 1 and 2.** They will diverge. PowerPoint routinely emits XML the strict schema wouldn't validate, omits optional attributes inconsistently, and uses undocumented conventions. **python-pptx's commitment is round-trip fidelity with what PowerPoint actually writes, not strict ECMA-376 conformance.** Prefer the real sample when the two conflict; use the XSD to understand structure and type, not to police PowerPoint's output.
4. **Check `src/pptx/oxml/…` for similar elements already modeled.** Copy the established pattern from a neighboring element — this keeps the descriptor layer consistent.

**The XSDs do not cover everything.** Microsoft extension namespaces — `cx:` (Office 2016+ chart types: funnel, treemap, sunburst, waterfall, histogram, box-whisker), `p14:` / `a14:` / `p15:` (2010/2013/2015-era PowerPoint additions), and the "modern comments" format — are not in the Part 1 XSDs. For these, you must rely on the real-sample approach in step 1 and/or Microsoft's extension documentation (MS-PPTX, MS-OE376, MS-ODRAWXML). Note this in the PR when relevant.

When you find a useful sample file or annotated fragment, drop it under `lab/` (scratch) rather than `spec/` (which is intentionally an immutable reference archive).

---

## 10. What NOT to do

- Don't amend or force-push to `master`, and never force-push to `upstream` under any circumstance.
- Don't commit secrets, API tokens, local `_scratch/` output, or generated docs (`.build/`).
- Don't add runtime dependencies lightly — every new dep affects a large user base. If you must, raise it first.
- Don't introduce backwards-incompatible API changes without a HISTORY note and a transition plan (deprecation warning where possible).
- Don't silence warnings with broad `filterwarnings` ignores — they exist to catch real problems.
- Don't delete `py.typed`; removing it silently breaks downstream type-checking.
- Don't "fix" code inside `lab/` or `spec/` just because lint would catch it elsewhere.

---

## 11. Quick command reference

| Task | Command |
|---|---|
| Unit tests | `pytest` |
| Unit test subtree | `pytest tests/chart` |
| Acceptance tests | `make accept` |
| Type check | `pyright src/pptx tests` |
| Lint | `ruff check .` |
| Format | `black src tests` |
| Coverage | `make coverage` |
| Build docs | `make docs` |
| All envs | `tox` |
| Build dist | `make build` |

---

_Last updated: 2026-05-02. Update this file when the layout, conventions, or workflows change._
