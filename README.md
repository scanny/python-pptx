# python-pptx

A Python library for creating, reading, and updating Microsoft PowerPoint
2007+ (`.pptx`) files.

Based on [python-openxml/python-pptx](https://github.com/scanny/python-pptx)
by Steve Canny and contributors. Forked at upstream `1.0.2` (2024-08-07)
and extended with 190+ additional OOXML features — full slide
duplication and cross-presentation merge, animations and transitions,
presentation sections, modern chart types (3D, combo, secondary axis,
error bars, chart templates, chartex passthrough for funnel/treemap/
waterfall), full shape effects (shadow/glow/reflection/soft-edge),
password-protected saves, comments, math equations, SmartArt scaffold,
custom document properties, slide-level tags, accessibility alt text,
and many more capabilities that were previously out of reach.

## Status

Unstable. Not yet published to PyPI. Install from source only.

Current version: `2026.05.0` (first release as an independent fork).
Versioning is CalVer (`YYYY.MM.patch`).

## Installation

```
pip install git+https://github.com/loadfix/python-pptx.git
```

Requires Python 3.8+.

## Example

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only layout
slide.shapes.title.text = "Hello, PowerPoint"

textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
textbox.text_frame.paragraphs[0].text = "It was a dark and stormy night."
textbox.text_frame.paragraphs[0].font.size = Pt(28)

prs.save("dark-and-stormy.pptx")

prs = Presentation("dark-and-stormy.pptx")
print(prs.slides[0].shapes.title.text)
# Hello, PowerPoint
```

The package is imported as `pptx`, matching upstream. Existing
upstream code runs unchanged against this fork.

## Features

See [`FEATURES.md`](FEATURES.md) for the full catalogue — 27 sections
covering every public capability, with fork additions marked
`[Added in 2026.05.0]`.

Summary of areas extended beyond upstream `1.0.2`:

- Slides (duplicate, delete, reorder, hide, section assignment,
  cross-presentation copy, full-fidelity deck merge)
- Presentations (`.ppsx` PowerPoint Show save, Flat OPC XML save,
  password-protected saves via Agile Encryption, reproducible zip
  timestamps, paper-size presets: `screen4x3` / `screen16x9` /
  `letter` / `A4` / arbitrary EMU tuples, blank-from-template,
  `Presentation.merge()`)
- Sections (full `p14:sectionLst` subsystem — create, rename, remove,
  reorder, assign slides, author-supplied GUIDs)
- Shapes (delete, duplicate, flip horizontal/vertical, hide, z-order,
  replace-with, find-by-xpath, get-by-name, accessibility alt text and
  title, effective geometry through group transforms, preset path
  geometry extraction)
- Autoshapes and connectors (line end arrows, connector bend
  adjustments, `MSO_SHAPE.LINE` support)
- Pictures (MPO images, EMF detection, linked pictures via external
  rels, fit-without-crop mode, `Picture.replace_image`, SVG error
  guidance)
- Media (audio MIME registration, `add_movie` autoplay kwarg, movie
  start time and condition, click-sound on hyperlinks, `Movie.blob` /
  `.ext` / `.content_type`, `Movie.replace_media`)
- Tables (per-cell borders including diagonals, `Table.style_id`,
  `_RowCollection.add` / `remove`, `_Row.delete`, `_ColumnCollection.add`
  / `remove`, `_Column.delete`, `_Cell.row_idx` / `col_idx`)
- Text frames, paragraphs, runs (rotation, `replace_text` across runs
  preserving origin-run formatting, full bullet API with
  font/color/size, strikethrough, East-Asian and complex-script font
  slots, `Font.effective_color` resolving through inheritance chain,
  `_Paragraph.delete` / `_Run.delete`, OMML math-equation read and
  write)
- Charts (2D + 3D, combo charts with `Chart.add_plot`, secondary value
  axis, error bars, chart templates via `Chart.apply_template`, chart
  duplication across slides, data-label borders/fills/font/text-frame
  fix, chartex sentinel passthrough for funnel/treemap/waterfall,
  theme-accent color cycling on replaced data, number-format
  preservation on `replace_data`, `Chart.update_cached_values`,
  `Chart.workbook` bytes access, `replace_data_preserve_formulas`,
  user-shapes part scaffold, manual chart-title position, axis
  position/cross_between/major-unit/category-num-format/tick-label
  rotation)
- Fills, colors, effects (full `a:effectLst` family — shadow, glow,
  reflection, soft-edge; universal `ColorFormat.to_rgb()` with
  scheme/preset/HSL/sysColor resolution; tint/shade via lumMod/lumOff;
  `ColorFormat.alpha` transparency; `FillFormat.blip_fill()` picture
  fills)
- Animations and transitions (F8 XML-layer MVP + leaf APIs:
  `Slide.transition` with 23 transition types including MORPH wrapped
  in `mc:AlternateContent`, duration, `advance_on_click`,
  `advance_after_time`, `PP_TRANSITION_SPEED`, `wipe_direction`;
  `Slide.iter_shape_animations` introspection;
  `Slide.animation_sequence`; `BaseShape.set_animation(type, trigger,
  delay)` with APPEAR/FADE_IN/FLY_IN/PULSE/FADE_OUT presets;
  `Presentation.set_auto_advance(seconds)` bulk auto-play)
- Comments (legacy `p:cmLst` read and write, per-slide `Slide.comments`
  collection with authors registered at presentation-part scope)
- Math equations (`Shape.has_math_equation` and `.math_equation_xml`
  read; `_Paragraph.add_math_equation(omml_xml)` write with
  `mc:AlternateContent` wrap and text fallback)
- SmartArt scaffold (`GraphicFrame.has_smart_art` /
  `GraphicFrame.smart_art`; raw XML passthrough via
  `SmartArt.data_xml` / `.layout_xml` / `.colors_xml` /
  `.quick_style_xml`; authoring deferred — see
  `docs/dev/analysis/f9-smartart.rst`)
- OLE embedding (arbitrary `prog_id` + extension support; built-in
  `PROG_ID.ZIP` / `PDF` / `DOC` / `HTML` / `XLSX` / `PPTX` / `DOCX`)
- Hyperlinks and click actions (hyperlink color override per run,
  `ActionSetting.screen_tip` tooltip, click-sound `<a:snd>`,
  run-level slide-jump target)
- Headers, footers, slide numbers (`p:hf` toggles on master / layout /
  notes master; `_Paragraph.add_field` for auto-refresh slide-number /
  date / footer fields)
- Placeholders (generic content placeholder supports
  `insert_picture` / `insert_table` / `insert_chart`; fit-without-crop
  mode; placeholder name preservation when slides are added from
  layouts)
- Font embedding (`.ttf` and `.otf` custom fonts via
  `Presentation.embed_font`)
- Document properties (`Presentation.extended_properties` for full
  app.xml metadata; `Presentation.custom_properties` dict-like access
  to `docProps/custom.xml` for DOCPROPERTY field codes; slide count
  auto-sync on save)
- Slide-level tags (`Slide.tags` dict-like access to
  `ppt/tags/*.xml` VBA-style custom tags)
- Security (XXE hardening on every XML parse, zip-bomb guard with
  configurable `PackageTooLargeError` threshold, `MAX_UNCOMPRESSED_PACKAGE_SIZE`
  environment override)
- Rendering (deliberately out of scope — python-pptx emits
  Open XML, it does not rasterize. See `docs/user/use-cases.rst` for
  LibreOffice / PowerPoint COM / Aspose integration paths to PDF /
  video / image output)

## Documentation

API and user-guide documentation lives under `docs/` and builds with
Sphinx (using `sphinx-rtd-theme`).

```
pip install -r requirements-docs.txt
make -C docs html
```

Open `docs/.build/html/index.html` in a browser.

## Contributing

Issues and pull requests are tracked at
<https://github.com/loadfix/python-pptx/issues>. Please file issues
against this fork; upstream's tracker is for upstream-shared concerns
only.

When contributing:

- Run the tests: `PYTHONPATH=src pytest tests/ -q` and `behave
  features/`.
- Keep `FEATURES.md` current when adding, modifying, or removing public
  API (see `CLAUDE.md` for contributor conventions).
- Consult `spec/` (XSD schemas and the ISO/IEC 29500 PDFs) for
  authoritative element ordering and cardinality when implementing new
  `CT_*` classes — but remember that PowerPoint's emitted XML is the
  authoritative interop target when it disagrees with the spec
  (see `CLAUDE.md` §10).

## License

MIT. See `LICENSE`. Inherited from upstream
`python-openxml/python-pptx`.

## Related projects

This project is part of a series of OOXML libraries under the loadfix
org:

- [loadfix/python-docx](https://github.com/loadfix/python-docx) — Word
- [loadfix/python-pptx](https://github.com/loadfix/python-pptx) — PowerPoint (this repo)
- [loadfix/python-xlsx](https://github.com/loadfix/python-xlsx) — Excel
