# python-pptx

A Python library for creating, reading, and updating Microsoft PowerPoint
2007+ (`.pptx`) files.

This repository is a fork of [python-pptx](https://github.com/scanny/python-pptx)
by Steve Canny. It builds on their original work with additional
PowerPoint capabilities beyond upstream `1.0.2` — full-fidelity slide
duplication and cross-presentation merging, animations and transitions,
presentation sections, modern chart types (3D, combo, secondary axis,
error bars, chartEx passthrough), shape effects (shadow, glow,
reflection, soft-edge), password-protected saves, slide comments, math
equations, SmartArt scaffolding, custom document properties, slide-level
tags, and accessibility alt text. Credit for the foundational library
goes to the original author.

## Installation

```
pip install git+https://github.com/loadfix/python-pptx.git
```

Requires Python 3.8+.

## Usage

```python
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "It was a dark and stormy night."
prs.save("dark-and-stormy.pptx")

prs = Presentation("dark-and-stormy.pptx")
print(prs.slides[0].shapes.title.text)
# It was a dark and stormy night.
```

The package is imported as `pptx`, matching upstream. Existing upstream
code runs unchanged against this fork.

## API

Full API and user-guide documentation lives under `docs/` and builds
with Sphinx (theme: Furo).

```
pip install Sphinx furo
python -m sphinx -b html docs docs/_build/html
```

A single-page catalogue of every public capability (with fork-era
additions marked `[Added in <version>.dev0]`) lives in `FEATURES.md`.

## Status

**Actively maintained** on `loadfix/python-pptx`. This fork is under
continuous development: upstream issues from `scanny/python-pptx` are
triaged and addressed here, and the fork ships 340+ audited-issue
capabilities layered on top of upstream `1.0.2`. See `FEATURES.md` for
the single-page capability catalogue (fork-era additions are marked
`[Added in <version>.dev0]`) and `HISTORY.rst` for the user-visible
changelog. The remaining upstream issues that do not represent a code
gap in this fork (usage questions, out-of-scope requests, environment
bugs, corrupt source files, duplicates) are catalogued in
`docs/community/issue-triage.rst` with a one-line disposition each.

Unstable. Not yet published to PyPI — install from source only. Current
version: `2026.05.0`; versioning is CalVer (`YYYY.MM.patch`). The public
API follows the upstream `pptx` surface, so upstream code runs
unchanged; fork-era additions may still shift before a tagged release.

## Contributing

Issues and pull requests are tracked at
<https://github.com/loadfix/python-pptx/issues>. Please file issues
against this fork; upstream's tracker is for upstream-shared concerns
only.

Run the test suite with `pytest` (unit) and `make accept` (behave
acceptance tests) before opening a PR. See `CLAUDE.md` for the broader
development workflow.

## License

MIT. See `LICENSE`. Inherited from upstream `scanny/python-pptx`.

## Related projects

Part of a family of document-rendering libraries:

- [docxjs](https://github.com/loadfix/docxjs) — browser-side DOCX → HTML renderer (TypeScript)
- [pptxjs](https://github.com/loadfix/pptxjs) — browser-side PPTX → HTML renderer (TypeScript)
- [xlsxjs](https://github.com/loadfix/xlsxjs) — browser-side XLSX → HTML renderer (TypeScript)
- [python-docx](https://github.com/loadfix/python-docx) — Python DOCX parser/generator
- [python-xlsx](https://github.com/loadfix/python-xlsx) — Python XLSX parser/generator
- [ooxml-validate](https://github.com/loadfix/ooxml-validate) — Python/.NET OOXML validator (wraps Microsoft Open XML SDK + LibreOffice)
