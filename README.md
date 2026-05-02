# python-pptx

A Python library for creating, reading, and updating Microsoft PowerPoint
2007+ (`.pptx`) files.

Based on [scanny/python-pptx](https://github.com/scanny/python-pptx) by
Steve Canny and contributors. Forked at upstream `1.0.2` (2024-08-07)
and extended with additional OOXML capabilities.

## Status

Unstable. Not yet published to PyPI. Install from source only.

Current version: `2026.05.0`. Versioning is CalVer (`YYYY.MM.patch`).

## Installation

```
pip install git+https://github.com/loadfix/python-pptx.git
```

Requires Python 3.8+.

## Example

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

## Features

This fork adds a range of PowerPoint capabilities beyond upstream
`1.0.2`, including full-fidelity slide duplication and cross-presentation
merging, animations and transitions, presentation sections, modern chart
types (3D, combo, secondary axis, error bars, chartEx passthrough),
shape effects (shadow, glow, reflection, soft-edge), password-protected
saves, slide comments, math equations, SmartArt scaffolding, custom
document properties, slide-level tags, and accessibility alt text.

## Documentation

API and user-guide documentation lives under `docs/` and builds with
Sphinx. The theme is Furo.

```
pip install Sphinx furo
python -m sphinx -b html docs docs/_build/html
```

## Contributing

Issues and pull requests are tracked at
<https://github.com/loadfix/python-pptx/issues>. Please file issues
against this fork; upstream's tracker is for upstream-shared concerns
only.

## License

MIT. See `LICENSE`. Inherited from upstream `scanny/python-pptx`.

## Related projects

This project is part of a series of OOXML libraries under the loadfix
org:

- [loadfix/python-docx](https://github.com/loadfix/python-docx) — Word
- [loadfix/python-pptx](https://github.com/loadfix/python-pptx) — PowerPoint
- [loadfix/python-xlsx](https://github.com/loadfix/python-xlsx) — Excel
