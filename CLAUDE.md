# CLAUDE.md

python-pptx fork (loadfix/python-pptx) — a mature, MIT-licensed Python library for creating, reading, and updating PowerPoint (`.pptx`) files. It parses and emits Office Open XML via `lxml`, does not require PowerPoint to be installed, and is intended to be industrial-grade — suitable for commercial use, which demands correctness and round-trip fidelity.

This project is one of a sibling series of OOXML libraries under the loadfix org:

- **loadfix/python-docx** — Word `.docx`
- **loadfix/python-pptx** — PowerPoint `.pptx` (this repo)
- **loadfix/python-xlsx** — Excel `.xlsx`

The three libraries share an architectural lineage (three-layer proxy/part/oxml pattern over lxml) and OOXML spec conventions. When implementing a feature that exists across the trio, consult the sibling repos for naming and API-shape precedent.

## Architecture

Three-layer pattern:

```
Presentation API  (src/pptx/api.py, src/pptx/presentation.py, src/pptx/slide.py, …)
    |  Proxy objects wrapping oxml elements (Shape, Slide, Chart, Table, …)
Parts Layer       (src/pptx/parts/*.py)
    |  XmlPart subclasses owning XML trees, managing relationships
oxml Layer        (src/pptx/oxml/*.py)
    |  CT_* element classes extending lxml.etree.ElementBase via `xmlchemy`
lxml              (XML parsing/serialization)
```

Cross-cutting concerns:
- `src/pptx/opc/` — Open Packaging Conventions (zip/rel/content-type machinery under the parts layer)
- `src/pptx/enum/` — public enumerations consumed from all layers

## Source Layout

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
typings/            custom type stubs (mainly for lxml)
HISTORY.rst         release-history changelog (user-visible)
pyproject.toml      build + tool config
Makefile            convenience targets: accept, docs, coverage, build
tox.ini             py38–py312 test envs
```

The `spec/` directory is **intentionally undisciplined**. Ruff and pyright exclude it. Do not "clean it up" or reformat its contents; it's a reference archive and that's a feature.

## Key Patterns

### CT_ Element Classes (oxml layer)

Define in `src/pptx/oxml/…`, register via the `ns` registry at the top of `src/pptx/oxml/__init__.py`. The descriptor vocabulary (`ZeroOrOne`, `OneAndOnlyOne`, `ZeroOrMore`, `RequiredAttribute`, `OptionalAttribute`, …) is the `xmlchemy` layer on top of `lxml.etree`. Don't bypass it with raw `etree` access in production code — the descriptors carry namespace, type, and default semantics.

- `successors` tuple must match XSD schema ordering exactly — consult `spec/ISO-IEC-29500-1/schemas/xsd/` for authoritative grammar.
- Read `docs/dev/xmlchemy.rst` first if you're new to the descriptor layer.
- Read a neighboring element class in the same subpackage (`src/pptx/oxml/…`) before adding a new one — it keeps the descriptor layer consistent.

### Part Classes

Parts live under `src/pptx/parts/…` and extend `XmlPart`. They own the XML tree for a single package part, manage relationships to other parts, and plug into the package assembly in `src/pptx/package.py`. The OPC machinery (zip / rel / content-type) lives under `src/pptx/opc/`.

### Proxy Objects

Proxy objects wrap `CT_*` oxml elements and form the public API (`Slide`, `Shape`, `Chart`, `Table`, `TextFrame`, `Run`, …). They live alongside their subpackage (`src/pptx/slide.py`, `src/pptx/shapes/…`, `src/pptx/chart/…`, `src/pptx/text/…`). Public API additions need to be exported through the relevant `__init__.py` and referenced from the API docs page.

### Constants

- Content types and relationship types: `src/pptx/opc/constants.py`.
- Namespaces: `src/pptx/oxml/ns.py` — `qn("a:solidFill")`, `nsmap`, `nsdecls`. Use these helpers consistently; don't hand-assemble Clark-notation strings.
- Units: `src/pptx/util.py` — `Emu`, `Pt`, `Inches`, `Cm`, `Mm`.

## OOXML spec vs Microsoft PowerPoint reality

Microsoft PowerPoint does NOT strictly implement ISO/IEC 29500 / ECMA-376. Treat the spec as a starting point, not ground truth.

- PowerPoint writes the **Transitional** flavor, not **Strict**. The 4th/5th/6th editions of ISO 29500-1 tightened the spec toward Strict; PowerPoint still emits Transitional namespaces that trace back to the original 1st edition / ECMA-376 2006.
- PowerPoint emits Microsoft extensions in the `p14:`, `p15:`, `a14:`, `c14:`, `cx:`, and related namespaces (PowerPoint 2010/2013/2016+), gated by `mc:AlternateContent` / `mc:Ignorable`. These are documented in the `[MS-PPTX]` / `[MS-OE376]` / `[MS-ODRAWXML]` extension series, not in the ISO PDFs under `spec/`.
- PowerPoint's reader tolerates out-of-order, extra, and missing elements that the spec forbids. PowerPoint's writer emits shapes the spec doesn't mandate. A spec-valid file is not automatically a file PowerPoint will open cleanly.
- **When the spec and PowerPoint disagree, match PowerPoint.** The canonical way to resolve ambiguity is: save a minimal `.pptx` from PowerPoint, unzip it, and inspect the XML. `spec/…/xsd/*.xsd` tells you what is *allowed*; PowerPoint tells you what is *interoperable*.

Workflow before writing code for a new element or feature:

1. **Produce a real PowerPoint sample.** Create a minimal `.pptx` in Microsoft PowerPoint that exercises the feature, unzip it, and read the XML. This is ground truth.
2. **Look up the grammar in `spec/ISO-IEC-29500-1/schemas/xsd/`** (or Part 2's `opc-xsd/` for packaging work). This gives the formal parent/child relationships, attribute types, defaults, and cardinality.
3. **Reconcile steps 1 and 2.** They will diverge. python-pptx's commitment is round-trip fidelity with what PowerPoint actually writes, not strict ECMA-376 conformance. Prefer the real sample when the two conflict; use the XSD to understand structure and type.
4. **Check `src/pptx/oxml/…` for similar elements already modeled.** Copy the established pattern from a neighboring element.

The XSDs do not cover Microsoft extensions (`cx:` 2016+ chart types, `p14:` / `a14:` / `p15:` additions, the modern-comments format). For these, rely on the real-sample approach and/or Microsoft's extension documentation (MS-PPTX, MS-OE376, MS-ODRAWXML).

## Test Conventions

- Framework: pytest with BDD-style naming.
- Test-function naming (configured in `pyproject.toml`):
  - classes: `Test…` or `Describe…`
  - functions: `test_…`, `it_…`, `they_…`, `but_…`, `and_…`
- Test layout mirrors `src/pptx/` (e.g., `src/pptx/chart/axis.py` ↔ `tests/chart/test_axis.py`).
- **Warnings are errors** (`filterwarnings = ["error"]`). A new `DeprecationWarning` anywhere will fail the suite — fix the cause.
- Test fixtures for XML are commonly built with the helpers in `tests/unitdata.py` and per-subpackage `unitdata` modules.
- Private-access suppression: some test files start with `# pyright: reportPrivateUsage=false` — fine, since unit tests exercise internals.
- Acceptance tests live under `features/` (behave + Gherkin). Naming uses a short area prefix: `cht-*` (charts), `dml-*` (DrawingML fill/line/effect/color), `shp-*` (shapes), `txt-*` (text), `tbl-*` (tables), `ph-*` (placeholders), `prs-*` (presentation-level), `act-*` (actions/hyperlinks). `features/environment.py` creates a `_scratch/` dir for generated pptx output.

## Commands

```bash
# Run tests
pytest                             # entire unit suite
pytest tests/chart                 # subtree

# Run a specific test
pytest -k <pattern>

# Run acceptance tests
make accept                        # behave --stop
behave features/cht-chart.feature  # single file
behave --tags=-wip                 # skip work-in-progress

# Type check
pyright src/pptx tests

# Lint / format
ruff check .
black src tests

# Coverage / full CI-equivalent / build
make coverage
tox                                # py38..py312 in parallel; pytest + behave per env
make build

# Install in dev mode
pip install -e ".[dev]"
```

## What NOT to do

- Don't amend or force-push to `master`, and never force-push to an upstream remote under any circumstance.
- Don't commit secrets, API tokens, local scratch output (`_scratch/`), or generated docs (`.build/`).
- Don't add runtime dependencies lightly — every new dep affects a large user base. If you must, raise it first.
- Don't introduce backwards-incompatible API changes without a HISTORY/FEATURES note and a transition plan (deprecation warning where possible).
- Don't silence warnings with broad `filterwarnings` ignores — they exist to catch real problems.
- Don't delete `py.typed`; removing it silently breaks downstream type-checking.
- Don't "fix" code inside `spec/` just because lint would catch it elsewhere — it's an intentionally undisciplined reference archive.
- Don't bypass the xmlchemy descriptor layer with raw `lxml.etree` access in production code — the descriptors carry namespace, type, and default semantics.
- Don't move unit tests out of their current location or rename test methods away from the `Describe*` / `it_*` BDD convention — test discovery relies on it.
- Don't reach for `# type: ignore` as a first resort when `pyright --strict` flags something — fix the types. If you must suppress, target the specific rule (e.g. `# pyright: ignore[reportPrivateUsage]`) and include a one-line reason.

## Common workflows

### Adding a new public method on an existing class
1. Implement in the appropriate `src/pptx/…` module.
2. Add unit tests in the mirrored test file under `tests/`.
3. Add a behave scenario to the relevant `*.feature` file and its step function.
4. Add `.. automethod::` to the corresponding `docs/api/…` page.
5. Update `docs/user/…` if end users should know about it.
6. Add a `HISTORY.rst` line and refresh the `FEATURES.md` entry.

### Adding a new enum value
- Enums live in `src/pptx/enum/`. They use a custom metaclass; read a neighboring enum first to see the pattern (in particular, the "return value" XML mapping).
- Update the enum's doc in `docs/api/enum/` if present.

### Adding a new XML element class
- Custom element classes live in `src/pptx/oxml/…`. Read `docs/dev/xmlchemy.rst` first — it explains `ZeroOrOne`, `OneAndOnlyOne`, `ZeroOrMore`, `RequiredAttribute`, etc.
- Consult `spec/ISO-IEC-29500-1/schemas/xsd/` for authoritative element ordering before declaring `successors`.
- Register the new element with the `ns` registry (see top of `src/pptx/oxml/__init__.py`).
- Save a minimal `.pptx` from PowerPoint that exercises the element, unzip it, and compare — **when the spec and PowerPoint disagree, match PowerPoint**.

## Important

- Before implementing a new feature or element class, consult `spec/` for authoritative schema information. The subdirectories: `spec/ISO-IEC-29500-1/schemas/xsd/` (Part 1 grammars for DrawingML, PresentationML, SpreadsheetML, shared types), `spec/ISO-IEC-29500-2/opc-xsd/` (Open Packaging Conventions), `spec/ISO-IEC-29500-3/` (Markup Compatibility — `mc:AlternateContent` / `mc:Choice` / `mc:Fallback`), `spec/ISO-IEC-29500-4/` (Transitional conformance). These are not runtime dependencies — they are the canonical sources for element ordering, attribute types, and cardinality. When you find a useful sample file or annotated fragment during investigation, keep it local to your worktree rather than committing it; `spec/` is intentionally an immutable reference archive.
- Keep `FEATURES.md` and `HISTORY.rst` current when adding, modifying, or deleting public API. `FEATURES.md` is a single-page catalogue of every public capability; fork-era additions are marked `[Added in <version>.dev0]`. For each change: add the new entry (or update/remove the existing one) under the relevant section, refresh the snippet if the API surface shifted, and verify the snippet runs against a fresh `Presentation()`.
- Always run tests after changes: `pytest` and `make accept`.
- Use `src/` layout — all code is under `src/pptx/`, not `pptx/`.
- Follow existing code style: short imperative-voice docstrings; `from __future__ import annotations` is used throughout; isort ordering (ruff `I` rule) keeps first-party `pptx` imports in a dedicated group.
- Public API additions need to be exported through the relevant `__init__.py` and referenced from the API docs page.
- Tooling is strict (Black line length 100, Ruff with a targeted rule set, Pyright in strict mode with custom stubs under `typings/`). Respect it — don't disable or silence lints/types to make a patch land.

## Documentation build

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
- `docs/index.rst` — top-level page; includes the "Feature Support" bullet list that must be updated when capability is added.

Sphinx config is in `docs/conf.py`.

## Release prep

Reference only — ask before running.

- See `docs/dev/development_practices.rst`.
- Update `src/pptx/__init__.py` version, `HISTORY.rst`, confirm docs compile, full tox run clean, build (`make build`), upload (`make upload`).
