# Features

`loadfix/python-pptx` is a fork of
[python-pptx](https://github.com/scanny/python-pptx) that extends the library
with full-fidelity presentation merge, slide duplication / reorder / delete,
hidden slides, section management, shape lifecycle (delete / duplicate /
replace / flip / z-order), effective (group-transform composited) geometry,
XPath and by-name shape lookup, picture / movie replacement, linked pictures,
combo and 3-D charts, secondary-value-axis charts, chart templates, cached-
value refresh, chart cloning across slides / packages, `replace_data`
variants that preserve formulas and number formats, full `ShadowFormat`,
alpha and `to_rgb` across every color type, blip (picture) fills, line arrow
ends, slide transitions (23 variants including MORPH), shape animations,
timing introspection, comments read/write, OMML equation read/write, SmartArt
scaffolding, OLE embedding for arbitrary progIds (ZIP/PDF/HTML/...), auto-
play, headers/footers/date/slide-numbers, Content placeholder chart/picture/
table insertion, font embedding (TTF + OTF), custom and extended document
properties, slide-level tags, XXE hardening + zip-bomb guard, paper-size
presets (letter / A4 / 16x9 / 4x3 / tuple), `.ppsx` read + write, Flat-OPC
XML save, reproducible zip timestamps, password-to-open (Agile Encryption),
and a long tail of smaller OOXML capabilities that were previously out of
reach.

This document is the full, single-page catalogue of what the library can do
today. Each section covers one feature area, opens with a short overview,
shows a copy-pasteable snippet against a fresh `Presentation()`, and then
lists the public methods, properties, and classes that make up that surface.
Items marked `[Added in 2026.05.0]` are additions from this fork — every
other item is inherited from the upstream base.

**Table of contents**

- [Opening and saving presentations](#opening-and-saving-presentations)
- [Slides](#slides)
- [Slide masters and layouts](#slide-masters-and-layouts)
- [Sections](#sections)
- [Shapes](#shapes)
- [Placeholders](#placeholders)
- [Auto shapes and connectors](#auto-shapes-and-connectors)
- [Group shapes](#group-shapes)
- [Pictures](#pictures)
- [Movies and audio](#movies-and-audio)
- [OLE objects and embedded files](#ole-objects-and-embedded-files)
- [Text frames and paragraphs](#text-frames-and-paragraphs)
- [Fonts and text formatting](#fonts-and-text-formatting)
- [Bullets, numbering, and fields](#bullets-numbering-and-fields)
- [Hyperlinks](#hyperlinks)
- [Fill and line format (DrawingML)](#fill-and-line-format-drawingml)
- [Shadow and effect format](#shadow-and-effect-format)
- [Tables](#tables)
- [Charts](#charts)
- [Chart data](#chart-data)
- [Chart axes and legend](#chart-axes-and-legend)
- [Chart series, points, and data labels](#chart-series-points-and-data-labels)
- [Chart templates and cloning](#chart-templates-and-cloning)
- [SmartArt and equations](#smartart-and-equations)
- [Action settings and click actions](#action-settings-and-click-actions)
- [Animations and timing](#animations-and-timing)
- [Transitions](#transitions)
- [Comments (legacy and threaded)](#comments-legacy-and-threaded)
- [Notes slides and notes master](#notes-slides-and-notes-master)
- [Header, footer, slide number, and date](#header-footer-slide-number-and-date)
- [Theme and colors](#theme-and-colors)
- [Font embedding](#font-embedding)
- [View props and presentation metadata](#view-props-and-presentation-metadata)
- [Document properties](#document-properties)
- [Encryption and protection](#encryption-and-protection)
- [Packaging and I/O options](#packaging-and-io-options)
- [Cross-presentation operations](#cross-presentation-operations)
- [Units and helpers](#units-and-helpers)
- [API concepts](#api-concepts)

---

## Opening and saving presentations

The top-level `pptx.Presentation()` factory opens a `.pptx`, `.pptm`,
`.potx`, `.potm`, or `.ppsx` package, or — when called with no argument —
creates a fresh 16:9 presentation from the bundled default template.
Flat-OPC (`<pkg:package>`) single-XML input and ECMA-376 Agile-Encryption
password-protected packages are auto-detected. `Presentation.save()`
serialises back to a path or stream; companion helpers emit slideshow
(`.ppsx`) and Flat-OPC variants.

```python
from pptx import Presentation

# create a fresh 16:9 deck
prs = Presentation()
slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(slide_layout)
slide.shapes.title.text = "Hello"
prs.save("out.pptx")

# open an existing file (path may be str, pathlib.Path, or file-like)
prs = Presentation("deck.pptx")

# open a .potx / .ppsx template package
prs = Presentation("brand.potx")

# password-to-open (Agile Encryption)
prs = Presentation("encrypted.pptx", password="s3cret")

# paper-size / aspect-ratio preset
wide = Presentation(pptx_format="16:9")
```

- `pptx.Presentation(pptx=None, password=None, pptx_format=None)` — Factory returning a `pptx.presentation.Presentation`. `password`, `pptx_format`, `.potx` / `.potm` / `.ppsx` templates, and Flat-OPC input are `[Added in 2026.05.0]`.
- `Presentation.save(file, password=None)` — Write the presentation to a path or file-like. `password=` re-encrypts with Agile Encryption. `[Added in 2026.05.0]`
- `Presentation.save_ppsx(file)` — Save as a PowerPoint slideshow (`.ppsx`) by flipping the main-part content-type. `[Added in 2026.05.0]`
- `Presentation.save_flat_xml(file)` — Save as single-file Flat-OPC XML. `[Added in 2026.05.0]`
- `Presentation.slide_width` / `Presentation.slide_height` — Read/write slide dimensions as `Emu` lengths.
- `Presentation.slides` / `Presentation.slide_layouts` / `Presentation.slide_master` / `Presentation.slide_masters` — Top-level collections.
- `Presentation.notes_master` — The notes master. `[Added in 2026.05.0]`
- `Presentation.sections` — `Sections` collection for presentation sections. `[Added in 2026.05.0]`
- `Presentation.first_slide_num` — Starting slide number (affects slide-number placeholders). `[Added in 2026.05.0]`
- `Presentation.strip_slides()` — Remove every slide (and any section definitions) in place, returning `self`; used to produce a blank deck carrying only a template's masters / layouts / theme / embedded fonts. `[Added in 2026.05.0]`
- `Presentation.merge(other_presentation)` — Append every slide in `other_presentation` to this presentation, cloning pictures, media, charts, OLE objects, and external hyperlinks into the target package. Returns the list of newly added `Slide` instances. `[Added in 2026.05.0]`
- `Presentation.embed_font(font_file, typeface, ...)` / `Presentation.embedded_fonts` — Embed TrueType / OpenType fonts into `fontEmbed*` parts. `[Added in 2026.05.0]`
- `Presentation.set_auto_advance(seconds, apply_to="all")` — Set slide auto-advance on every (or the current) slide in one call. `[Added in 2026.05.0]`
- `pptx.exc.PackageNotFoundError` / `pptx.exc.InvalidXmlError` / `pptx.exc.PythonPptxError` — Library-specific exceptions. `EncryptedPackageError` for password-protected packages is `[Added in 2026.05.0]`.

---

## Slides

`Presentation.slides` is the main collection: add a slide from any layout,
duplicate or move a slide, delete it, or import a slide from another
presentation. Each `Slide` carries shapes, placeholders, comments, tags,
transitions, animations, and a flag indicating whether the slide is hidden.

A `Presentation.slides` collection supports add, duplicate, delete, reorder
(move), and external-slide-copy. Each `Slide` exposes `shapes`,
`placeholders`, `slide_layout`, `slide_id`, `notes_slide`, `has_notes_slide`,
`background`, `follow_master_background`, `name`, and the fork-era
`is_hidden`, `show_master_shapes`, `comments`, `has_comments`, `tags`,
`has_tags`, `transition`, `animation_sequence`, `iter_shape_animations`,
`has_animations`, `timing_xml`, and `find_shapes_by_xpath`.

```python
from pptx import Presentation

prs = Presentation()
title_layout = prs.slide_layouts[0]
blank_layout = prs.slide_layouts[5]

slide1 = prs.slides.add_slide(title_layout)
slide1.shapes.title.text = "First"

slide2 = prs.slides.add_slide(blank_layout)
slide2.is_hidden = True

# insert a slide at a specific position (index= added in 2026.05.0)
intro = prs.slides.add_slide(title_layout, index=0)

# duplicate and move
dup = prs.slides.duplicate(slide1)
prs.slides.move_slide(dup, 1)

# delete
prs.slides.delete(slide2)

prs.save("out.pptx")
```

- `Presentation.slides` — `Slides` sequence.
- `Slides.add_slide(slide_layout, index=None)` — Add a slide bound to `slide_layout`; append by default, or insert at zero-based `index` (negative counts from end, out-of-range clamps). The `index` keyword was added in 2026.05.0 (#194).
- `Slides.duplicate(slide, index=None)` — Clone a slide within the presentation, optionally at `index`. `[Added in 2026.05.0]`
- `Slides.delete(slide)` — Remove a slide from the deck and drop its part. Parts uniquely referenced by the deleted slide (its notes slide, images, charts, embedded workbooks, media) are garbage-collected at save time via rels-graph reachability, so they do not linger in the saved ``.pptx`` zip. Parts shared with another slide (e.g. an image on a duplicate) are preserved. Verified by `tests/test_issue_956_delete_slide_cleanup.py`. `[Added in 2026.05.0]`
- `Slides.move_slide(slide, new_idx)` — Reorder a slide to `new_idx`. `[Added in 2026.05.0]`
- `Slides.add_slide_from_external(source_slide, slide_layout)` — Copy a slide from another presentation, rewiring images, media, and other part-level dependencies. `[Added in 2026.05.0]`
- `Slides.get(slide_id, default=None)` / `Slides.index(slide)` — Slide-id and position lookups.
- `Slide.slide_id` — Stable integer identifier assigned by PowerPoint.
- `Slide.slide_layout` — The `SlideLayout` this slide inherits from.
- `Slide.shapes` / `Slide.placeholders` — `SlideShapes` and `SlidePlaceholders`.
- `Slide.name` — Read/write slide name (back-fills to `sld{n}` if unset).
- `Slide.is_hidden` — Read/write boolean for hidden-slide flag. `[Added in 2026.05.0]`
- `Slide.follow_master_background` — True when the slide inherits its background from the master / layout.
- `Slide.background` — `_Background` proxy; exposes `fill` for solid / gradient / picture backgrounds.
- `Slide.has_notes_slide` / `Slide.notes_slide` — Notes access.
- `Slide.has_comments` / `Slide.comments` — Legacy and threaded comments. `[Added in 2026.05.0]`
- `Slide.has_tags` / `Slide.tags` — Slide-level custom tags. `[Added in 2026.05.0]`
- `Slide.transition` — Slide transition proxy. `[Added in 2026.05.0]`
- `Slide.has_animations` / `Slide.timing_xml` / `Slide.animation_sequence` / `Slide.iter_shape_animations()` — Animation introspection. `[Added in 2026.05.0]`
- `Slide.find_shapes_by_xpath(xpath_expr)` — Evaluate an XPath against this slide's shape tree and return matching `BaseShape` proxies. `[Added in 2026.05.0]`

- `Presentation.slides` — `Slides` collection (sequence).
- `Slides.add_slide(slide_layout, index=None)` — Append a new slide (or insert it at zero-based `index`; ``index`` keyword added in 2026.05.0, #194).
- `Slides.duplicate(slide, index=None)` — Deep-clone an existing slide within the presentation. `[Added in 1.0.2.dev0]`
- `Slides.delete(slide)` — Remove a slide. `[Added in 1.0.2.dev0]`
- `Slides.move_slide(slide, new_idx)` — Reorder. `[Added in 1.0.2.dev0]`
- `Slides.add_slide_from_external(source_slide, slide_layout)` — Full-fidelity copy from another presentation (images / charts / OLE rewritten into this package). `[Added in 1.0.2.dev0]`
- `Slides.get(slide_id, default=None)` — Look up by slide ID.
- `Slides.index(slide)` — Positional lookup.
- `Slide.slide_id` / `Slide.slide_layout` / `Slide.shapes` / `Slide.placeholders` / `Slide.name` / `Slide.element` / `Slide.part`.
- `Slide.is_hidden` (read/write `bool`) — `p:sld/@show="0"` for hidden slides. `[Added in 1.0.2.dev0]`
- `Slide.show_master_shapes` (read/write `bool`) — `p:sld/@showMasterSp="0"` hides the master's non-placeholder shapes (e.g. a company logo) from this slide; mirrors PowerPoint's *Hide Background Graphics* checkbox. `[Added in 1.0.2.dev0]`
- `Slide.background` / `Slide.follow_master_background()` — Per-slide background.
- `Slide.has_notes_slide` / `Slide.notes_slide` — Lazy notes page.
- `Slide.find_shapes_by_xpath(xpath_expr)` — Evaluate a namespaced XPath against `p:spTree` and return matching shapes. `[Added in 1.0.2.dev0]`
- `Slide.comments` / `Slide.has_comments` — Legacy-format comments (see [Comments](#comments)). `[Added in 1.0.2.dev0]`
- `Slide.tags` / `Slide.has_tags` — Slide-level tag dict (see [Slide-level tags](#slide-level-tags)). `[Added in 1.0.2.dev0]`
- `Slide.transition` — `Transition` proxy (see [Animations and transitions](#animations-and-transitions)). `[Added in 1.0.2.dev0]`
- `Slide.has_animations` / `Slide.timing_xml` — Animation introspection. `[Added in 1.0.2.dev0]`
- `Slide.animation_sequence` — Tuple of `AnimationEffectView` for the main sequence. `[Added in 1.0.2.dev0]`
- `Slide.iter_shape_animations()` — Iterator of `ShapeAnimation` proxies for every shape-targeted effect. `[Added in 1.0.2.dev0]`

---

## Slide masters and layouts

Every slide inherits text styles, placeholders, color theme, and background
from its layout, which in turn inherits from its master. The fork adds
`SlideLayouts.remove()`, `.index()`, `.get_by_name()`,
`SlideMaster.add_layout_from()`, and the ability to add non-placeholder
shapes onto masters and layouts.

```python
from pptx import Presentation

prs = Presentation()
master = prs.slide_master
for layout in master.slide_layouts:
    print(layout.name)

# look up a layout by user-visible name
blank = master.slide_layouts.get_by_name("Blank")

# or by presentation-stable id (robust against layout reordering)
layout_id = master.slide_layouts[1].slide_layout_id
same_layout = master.get_layout(layout_id)

# theme colors available to this master
for key, rgb in master.theme_colors.items():
    print(key, rgb)

# import a layout from another presentation's master
brand = Presentation("branded.pptx").slide_masters[0]
imported = master.add_layout_from(brand.slide_layouts.get_by_name("Callout"))
```

- `Presentation.slide_master` / `Presentation.slide_masters` — Default master plus every master in the deck.
- `SlideMaster.name` / `SlideMaster.slide_layouts` / `SlideMaster.placeholders` / `SlideMaster.shapes`.
- `SlideMaster.theme_colors` — Mapping of theme-color role names to resolved `RGBColor`. `[Added in 2026.05.0]`
- `SlideMaster.get_layout(layout_id, default=None)` — Layout lookup by presentation-stable `p:sldLayoutId/@id` (robust against reordering). `[Added in 2026.05.0]`
- `SlideMaster.header_footer` — `_HeaderFooter` proxy.
- `SlideMaster.add_layout_from(source_layout)` — Clone a slide layout from any other master (same or different presentation) into this master. Returns the new `SlideLayout`; raises `ValueError` on a name collision. `[Added in 2026.05.0]`
- `SlideLayouts.__getitem__` / `__iter__` / `__len__` — Index, iterate, count layouts.
- `SlideLayouts.get_by_name(name, default=None)` — Layout lookup by name. `[Added in 2026.05.0]`
- `SlideLayouts.get_by_id(layout_id, default=None)` — Layout lookup by presentation-stable id. `[Added in 2026.05.0]`
- `SlideLayouts.index(slide_layout)` — Position of a layout within the master. `[Added in 2026.05.0]`
- `SlideLayouts.remove(slide_layout)` — Delete an unused slide layout. `[Added in 2026.05.0]`
- `SlideLayout.name` / `.placeholders` / `.shapes` / `.slide_master` / `.header_footer`.
- `SlideLayout.slide_layout_id` — Presentation-stable integer id from `p:sldLayoutId/@id`. `[Added in 2026.05.0]`
- `SlideLayout.iter_cloneable_placeholders()` — Iterator over placeholders that will be cloned onto new slides.
- `SlideLayout.used_by_slides` — Tuple of slides currently using this layout. `[Added in 2026.05.0]`
- `NotesMaster.placeholders` / `NotesMaster.shapes` / `NotesMaster.header_footer` — Notes-master access.

---

## Sections

Presentation sections — PowerPoint's "Section" feature — group slides
into named, reorderable chunks. Add sections, rename them, move them,
move individual slides between sections, and read section membership.
`[Added in 2026.05.0]`.

```python
from pptx import Presentation

prs = Presentation()
layout = prs.slide_layouts[5]

intro = prs.slides.add_slide(layout)
body = prs.slides.add_slide(layout)

prs.sections.add_section(name="Intro", first_slide=intro)
prs.sections.add_section(name="Body", first_slide=body)

for sec in prs.sections:
    print(sec.id, sec.name, [s.slide_id for s in sec.slides])

prs.save("out.pptx")
```

- `Presentation.sections` — `Sections` collection. `[Added in 2026.05.0]`
- `Sections.add_section(name, first_slide)` / `Sections.remove(section)` / `Sections.index(section)` — CRUD. `[Added in 2026.05.0]`
- `Sections.find_containing(slide)` — The `Section` a slide belongs to, or `None`. `[Added in 2026.05.0]`
- `Sections.get_by_id(id)` / `Sections.get_by_name(name)` — Lookups. `[Added in 2026.05.0]`
- `Section.id` / `.index` / `.name` / `.slides` — Core metadata and membership. `[Added in 2026.05.0]`
- `Section.add_slide(slide)` / `Section.move_slide(slide)` — Move a slide into a section. `[Added in 2026.05.0]`
- `Section.move_before(other)` / `Section.move_after(other)` — Reorder sections. `[Added in 2026.05.0]`

---

## Shapes

`BaseShape` is the root of the shape hierarchy (auto shapes, pictures,
graphic frames, group shapes, placeholders, connectors, text boxes). The
fork adds deletion, duplication, swap-replace, hide/show, flip helpers,
z-order mutation, alt-text, animations, and shape look-up by name.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
shapes = slide.shapes

# add a rounded rectangle
rect = shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(1), Inches(1), Inches(4), Inches(1),
)
rect.text_frame.text = "Hello"
rect.alt_text = "Greeting box"

# duplicate, flip, bring to front
dup = rect.duplicate()
dup.flip_horizontal = True
dup.bring_to_front()

# find later by name
hit = shapes.get_by_name(rect.name)

prs.save("out.pptx")
```

- `Slide.shapes` — `SlideShapes` ordered collection.
- `SlideShapes.add_shape(shape_type, left, top, width, height)` — Append a preset auto shape.
- `SlideShapes.add_textbox(left, top, width, height)` — Append a standalone text box.
- `SlideShapes.add_picture(path_or_stream, left, top, width=None, height=None)` — Inline picture.
- `SlideShapes.add_picture_link(url, left, top, width=None, height=None)` — Linked (external URL) picture. `[Added in 2026.05.0]`
- `SlideShapes.add_movie(path_or_stream, left, top, width, height, poster_frame_image=None, mime_type=None, autoplay=False)` — Add a video/audio shape. `autoplay` is `[Added in 2026.05.0]`.
- `SlideShapes.add_chart(chart_type, x, y, cx, cy, chart_data)` — Append a chart graphic frame.
- `SlideShapes.clone_chart(chart, x, y, cx=None, cy=None)` — Copy a chart from another slide/presentation into this slide. `[Added in 2026.05.0]`
- `SlideShapes.add_connector(connector_type, begin_x, begin_y, end_x, end_y)` — Connector line between two points.
- `SlideShapes.add_group_shape(shapes=())` — Wrap shapes in a group. `[Added in 2026.05.0]`
- `SlideShapes.add_ole_object(object_file, prog_id, left, top, width=None, height=None, icon_file=None, icon_width=None, icon_height=None)` — Embed an OLE payload. Generic `prog_id` + `extension` support is `[Added in 2026.05.0]`.
- `SlideShapes.add_table(rows, cols, left, top, width, height)` — Append a table graphic frame. Accepts float dimensions. `[Added in 2026.05.0]` for the float overload.
- `SlideShapes.build_freeform(start_x, start_y, scale=EMU_PER_INCH)` — Freeform-geometry builder.
- `SlideShapes.find_all_by_name(name)` / `SlideShapes.get_by_name(name)` — Look up shapes by `cNvPr/@name`. `[Added in 2026.05.0]`
- `SlideShapes.title` — Title placeholder shape, if any.
- `SlideShapes.placeholders` — `SlidePlaceholders` view.
- `SlideShapes.turbo_add_enabled` — Bulk-add performance mode for programmatic authoring.
- `BaseShape.shape_type` / `.shape_id` / `.name` / `.left` / `.top` / `.width` / `.height` / `.rotation` — Core geometry and identity.
- `BaseShape.effective_left` / `.effective_top` / `.effective_width` / `.effective_height` — Group-aware absolute geometry. `[Added in 2026.05.0]`
- `BaseShape.flip_horizontal` / `.flip_vertical` / `.flip_horizontally()` / `.flip_vertically()` — Mirroring. `[Added in 2026.05.0]`
- `BaseShape.alt_text` / `.title` — Accessibility fields. `[Added in 2026.05.0]`
- `BaseShape.is_hidden` — Show/hide a shape without deleting it. `[Added in 2026.05.0]`
- `BaseShape.shadow` — `ShadowFormat` proxy (with `.inherit`, `.blur_radius`, `.distance`, `.angle`, etc.). `[Added in 2026.05.0]`
- `BaseShape.click_action` — `ActionSetting` for on-click behaviour.
- `BaseShape.animation` / `BaseShape.set_animation(...)` — Read or author an entrance / exit / emphasis effect on this shape. `[Added in 2026.05.0]`
- `BaseShape.delete()` — Remove the shape element from its tree and drop orphan relationships. `[Added in 2026.05.0]`
- `BaseShape.duplicate()` — Clone this shape in place. `[Added in 2026.05.0]`
- `BaseShape.replace_with(other_shape)` — Swap this shape for another (geometry + z-order preserved). `[Added in 2026.05.0]`
- `BaseShape.bring_forward()` / `.send_backward()` / `.bring_to_front()` / `.send_to_back()` / `.zorder_index` — Z-order mutation. `[Added in 2026.05.0]`
- `BaseShape.has_chart` / `.has_table` / `.has_text_frame` / `.is_placeholder` — Feature predicates.
- `BaseShape.has_math_equation` / `.math_equation_xml` — OMML inspection. `[Added in 2026.05.0]`
- `BaseShape.placeholder_format` — `_PlaceholderFormat` with `.idx` and `.type`.

- `Presentation.slide_masters` — `SlideMasters` sequence.
- `Presentation.slide_layouts` — `SlideLayouts` sequence on the primary master.
- `SlideMaster.shapes` / `SlideMaster.placeholders` — Shape tree (write-enabled `[Added in 1.0.2.dev0]`).
- `SlideMaster.name` (read/write) — `[Added in 1.0.2.dev0]`. Returns `"Master N"` positional fallback when `p:cSld/@name` is empty.
- `SlideMaster.slide_layouts` — Child layouts.
- `SlideMaster.background` / `SlideMaster.theme_colors` — Master-level theming.
- `SlideMaster.header_footer` — Default header/footer overrides.
- `SlideLayouts.get_by_name(name, default=None)` — Name lookup.
- `SlideLayouts.get_by_type(layout_type, default=None)` — Lookup by `ST_SlideLayoutType` token (`"title"`, `"blank"`, `"cust"`, ...). Useful for Google-Slides-origin decks where names are unreliable. `[Added in 1.0.2.dev0]`
- `SlideLayouts.index(slide_layout)` — Positional lookup.
- `SlideLayouts.remove(slide_layout)` — Delete an unused layout.
- `SlideLayout.shapes` / `SlideLayout.placeholders` — Shape tree (write-enabled `[Added in 1.0.2.dev0]`).
- `SlideLayout.name` (read/write) — `[Updated in 1.0.2.dev0]`. Returns `"Layout N"` positional fallback when `p:cSld/@name` is empty (common for Google Slides exports — issue #864).
- `SlideLayout.slide_layout_type` — `ST_SlideLayoutType` token (`"title"`, `"obj"`, `"blank"`, `"cust"`, ...); defaults to `"cust"` when the attribute is absent. `[Added in 1.0.2.dev0]`
- `SlideLayout.used_by_slides` / `SlideLayout.slide_master` / `SlideLayout.header_footer`.
- `SlideLayout.iter_cloneable_placeholders()` — Placeholders new slides inherit.
- `MasterShapes` / `LayoutShapes` — Shape collections; inherit `add_shape`, `add_picture`, `add_textbox`, `add_connector`, `add_group_shape`, `build_freeform`. `[Added in 1.0.2.dev0]`

---

## Placeholders

Layout and slide placeholders inherit text, geometry, and theming from the
slide master. `SlidePlaceholder` subclasses expose type-specific ergonomics
(picture, chart, table, content).

```python
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])

# enumerate placeholders with idx and type
for ph in slide.placeholders:
    print(ph.placeholder_format.idx, ph.placeholder_format.type, ph.name)

title = slide.placeholders[0]
title.text = "Welcome"

prs.save("out.pptx")
```

- `Slide.placeholders` — `SlidePlaceholders` collection (indexed by `idx`).
- `SlidePlaceholders.__getitem__(idx)` / `__iter__` / `get(idx, default)` — Access by placeholder idx.
- `SlidePlaceholder` / `LayoutPlaceholder` / `MasterPlaceholder` / `NotesSlidePlaceholder` — Type-specific subclasses.
- `SlidePlaceholder.insert_picture(image_file)` — Replace a content / picture placeholder with an image, auto-fitting. `[Added in 2026.05.0]` for the content-placeholder variant.
- `SlidePlaceholder.insert_chart(chart_type, chart_data)` — Insert a chart into a content placeholder. `[Added in 2026.05.0]`
- `PicturePlaceholder.insert_picture(image_file)` — Classic picture-placeholder swap.
- `TablePlaceholder.insert_table(rows, cols)` — Replace with a table.
- `LayoutPlaceholders.get(idx, default=None)` — Master / layout placeholder lookup by idx.
- `BaseShape.is_placeholder` / `.placeholder_format.idx` / `.placeholder_format.type` — Placeholder metadata on any shape.
- Enum: `PP_PLACEHOLDER` (title / body / subtitle / date / slide-number / footer / header / object / chart / table / picture / etc.).

---

## Auto shapes and connectors

Auto shapes are parameterised preset geometries (`MSO_SHAPE`). Connectors
are lines that snap between two shapes and remember their bend /
adjustment state.

```python
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

a = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                           Inches(1), Inches(1), Inches(2), Inches(1))
b = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                           Inches(5), Inches(3), Inches(2), Inches(1))
c = slide.shapes.add_connector(
    MSO_CONNECTOR_TYPE.ELBOW, a.left + a.width, a.top,
    b.left, b.top + b.height // 2,
)
c.line.width = Inches(0.05)

prs.save("out.pptx")
```

- `SlideShapes.add_shape(shape_type, left, top, width, height)` — Preset auto shape.
- `SlideShapes.add_connector(connector_type, begin_x, begin_y, end_x, end_y)` — Straight / elbow / curved connector.
- `Shape.auto_shape_type` — Read/write preset type of an auto shape (fixed to handle `line` is `[Added in 2026.05.0]`).
- `Shape.adjustments` — Tuple of adjustment handles (fractions in 0..1) that tweak the preset geometry. Elbow/curved-connector adjustment access is `[Added in 2026.05.0]`.
- `Shape.line` — `LineFormat` for outline colour, width, dash, and arrow heads.
- `Shape.fill` — `FillFormat` for solid/gradient/pattern/picture fill.
- `Shape.click_action` — `ActionSetting` for on-click behaviour.
- `Connector.begin_connect(shape, connection_point)` / `Connector.end_connect(shape, connection_point)` — Attach connector ends to specific connection points.
- `Connector.begin_x` / `.begin_y` / `.end_x` / `.end_y` — Read/write connector endpoint coordinates.
- `FreeformBuilder` — Returned by `SlideShapes.build_freeform()`; `.add_line_segments(vertices, close=True)` / `.convert_to_shape(origin_x=0, origin_y=0)`.
- Preset-shape-path geometry (`Shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE`) is exposed read-only via `BaseShape.element`. `[Added in 2026.05.0]`
- Enums: `MSO_SHAPE`, `MSO_CONNECTOR_TYPE`, `MSO_SHAPE_TYPE`, `MSO_LINE_DASH_STYLE` (corrected `ROUND_DOT` and `UP_DOWN_ARROW` mappings are `[Added in 2026.05.0]`).

---

## Group shapes

Group shapes nest other shapes. The fork adds group creation, deletion,
duplication, and group-aware effective coordinates.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
shapes = slide.shapes

a = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
b = shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(3), Inches(2), Inches(1))

# group both shapes together
group = shapes.add_group_shape([a, b])
print(group.width, group.height)

# duplicate the entire group
dup = group.duplicate()

prs.save("out.pptx")
```

- `SlideShapes.add_group_shape(shapes=())` — Wrap an iterable of shapes in a new `p:grpSp`. `[Added in 2026.05.0]`
- `GroupShape.shapes` — `GroupShapes` collection of direct children.
- `GroupShape.duplicate()` — Clone the entire group, including every descendant. `[Added in 2026.05.0]`
- `GroupShape.delete()` — Remove the group and drop orphan part references. `[Added in 2026.05.0]`
- `GroupShapes.add_shape(...)` / `.add_textbox(...)` / `.add_picture(...)` — Author new shapes inside a group.
- `BaseShape.effective_left` / `effective_top` / `effective_width` / `effective_height` — Group-transform-aware absolute geometry in slide space. `[Added in 2026.05.0]`

---

## Pictures

`SlideShapes.add_picture()` adds an inline picture; `add_picture_link()`
adds an external-linked picture. Pictures support outline, crop, alt-text,
and (fork-era) in-place image replacement and mask-shape cropping.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

pic = slide.shapes.add_picture("logo.png", Inches(1), Inches(1),
                               width=Inches(3))
pic.alt_text = "Company logo"

# replace image bytes, keep the shape identity
pic.replace_image("logo-new.png")

# crop edges (fractions of the original image)

# accessibility metadata (issue #508)
clone.alt_text = "decorative rounded rectangle"
clone.title = "Decoration"

# swap clone for a freshly-added oval in the same z-slot + geometry
oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, 0, 0, Inches(1), Inches(1))
clone.replace_with(oval)

# delete outright
rect.delete()

# lookup by name
found = slide.shapes.get_by_name("Oval 4")
matches = slide.shapes.find_all_by_name("Rectangle 1")

prs.save("out.pptx")
```

- `BaseShape.shape_id` / `BaseShape.name` / `BaseShape.shape_type` — Core identity.
- `BaseShape.left` / `.top` / `.width` / `.height` — Raw geometry in the enclosing group's coordinate system.
- `BaseShape.effective_left` / `.effective_top` / `.effective_width` / `.effective_height` — Slide-relative geometry after group-transform cascade. `[Added in 1.0.2.dev0]`
- `BaseShape.rotation` — Read/write float degrees.
- `BaseShape.flip_horizontal` / `BaseShape.flip_vertical` — Read/write `bool`. `[Added in 1.0.2.dev0]`
- `BaseShape.flip_horizontally()` / `BaseShape.flip_vertically()` — Toggle helpers. `[Added in 1.0.2.dev0]`
- `BaseShape.is_hidden` — Read/write `bool` (`cNvPr/@hidden`). `[Added in 1.0.2.dev0]`
- `BaseShape.alt_text` — Read/write `str` accessibility description (`cNvPr/@descr`); empty string when unset. `[Added in 1.0.2.dev0]`
- `BaseShape.title` — Read/write `str` accessibility title (`cNvPr/@title`); empty string when unset. `[Added in 1.0.2.dev0]`
- `BaseShape.delete()` — Remove from parent shape tree, tidying relationships. `[Added in 1.0.2.dev0]`
- `BaseShape.duplicate()` — Deep-clone in place (subclass-aware; `GroupShape.duplicate()` preserves nested relationships). `[Added in 1.0.2.dev0]`
- `BaseShape.replace_with(other_shape)` — Copy this shape's geometry onto `other_shape`, move it into this shape's z-order, then delete this. `[Added in 1.0.2.dev0]`
- `BaseShape.bring_to_front()` / `.bring_forward()` / `.send_backward()` / `.send_to_back()` / `.zorder_index` — Z-order. `[Added in 1.0.2.dev0]`
- `BaseShape.has_text_frame` / `.text_frame` / `.has_chart` / `.has_table` / `.is_placeholder` — Feature probes.
- `BaseShape.has_math_equation` / `.math_equation_xml` — OMML equation read (see [Math equations](#math-equations)). `[Added in 1.0.2.dev0]`
- `BaseShape.shadow` — `ShadowFormat` with full read/write `.blur_radius` / `.distance` / `.direction` / `.color` (see [Fills, colors, and effects](#fills-colors-and-effects)).
- `BaseShape.click_action` — `ActionSetting` (hyperlink / click action).
- `BaseShape.hover_action` — `ActionSetting` bound to `a:hlinkMouseOver`; parallel to `click_action` for mouse-over behaviors. `[Added in 1.0.2.dev0]`
- `BaseShape.placeholder_format` — `_PlaceholderFormat` or `None`.
- `BaseShape.animation` — Read-only `AnimationEffect` proxy for this shape's main-sequence animation. `[Added in 1.0.2.dev0]`
- `BaseShape.set_animation(effect_type, trigger="onClick", delay=0)` — Author an entrance/exit/emphasis preset (see [Animations and transitions](#animations-and-transitions)). `[Added in 1.0.2.dev0]`
- `BaseShape.element` — Underlying oxml element.
- `SlideShapes.add_shape(shape_type, left, top, width, height)` — Add an autoshape.
- `SlideShapes.add_textbox(left, top, width, height)` — Add a plain text box.
- `SlideShapes.add_picture(image_file, left, top, width=None, height=None)` — Add an embedded picture (see [Pictures](#pictures)).
- `SlideShapes.add_picture_link(url, left, top, width=None, height=None)` — Add a linked (external URL) picture. `[Added in 1.0.2.dev0]`
- `SlideShapes.add_connector(connector_type, begin_x, begin_y, end_x, end_y)` — Add a straight connector.
- `SlideShapes.add_group_shape(shapes=())` — Group existing shapes.
- `SlideShapes.add_chart(chart_type, x, y, cx, cy, chart_data)` — Add a chart (see [Charts](#charts)).
- `SlideShapes.add_table(rows, cols, left, top, width, height)` — Add a table.
- `SlideShapes.add_movie(...)` — Add a movie / audio (see [Media](#media-audio-and-video)).
- `SlideShapes.add_ole_object(...)` — Embed an OLE object (see [OLE embedding](#ole-embedding)).
- `SlideShapes.build_freeform(start_x=0, start_y=0, scale=1.0)` — Freeform builder.
- `SlideShapes.clone_chart(source_chart, x, y, cx, cy)` — Cross-slide / cross-deck chart duplicate. `[Added in 1.0.2.dev0]`
- `SlideShapes.get_by_name(name)` / `SlideShapes.find_all_by_name(name)` — Name-based lookup. `[Added in 1.0.2.dev0]`
- `SlideShapes.title` — The title placeholder, or `None`.
- `Slide.find_shapes_by_xpath(xpath_expr)` — XPath lookup (Open-XML namespace map pre-bound). `[Added in 1.0.2.dev0]`
- `GroupShape.shapes` — Child shapes.
- `GroupShape.duplicate()` — Override placing the clone at the source's slide-relative rectangle, with fresh unique IDs. `[Added in 1.0.2.dev0]`
- `GroupShape.ungroup()` — Dissolve the group, hoisting each direct child onto the slide's top-level shape tree at its slide-relative effective rectangle (Wave 3 #925 cascade); returns the freed shapes in original z-order. Handles nested groups; a freed child-group keeps its internal `chOff`/`chExt` intact so its descendants still render at their original slide positions. `[Added in 1.0.2.dev0]`

---

## Autoshapes and connectors

Autoshapes are added via `SlideShapes.add_shape(MSO_SHAPE.*)`. The fork
adds `MSO_SHAPE.LINE` (so `.auto_shape_type` on a straight-line `prst="line"`
resolves cleanly), read/write `Connector.adjustments` for elbow and curved
connectors, and line-end arrow configuration via `LineFormat.begin_arrow` /
`.end_arrow`.

PowerPoint's Insert Shapes lists *Connector with Arrow*, *Connector with
Double Arrow*, *Curved with Arrow*, and *Elbow with Arrow* as separate
gallery tiles, but internally those are just an `MSO_CONNECTOR` geometry
(`STRAIGHT` / `ELBOW` / `CURVE`) whose `a:ln` carries an `a:headEnd` and/or
`a:tailEnd` decoration. To author any of the arrow variants, set
`connector.line.begin_arrow.type` / `.end_arrow.type` after
`add_connector(...)` — there is no separate connector-type enum for arrow
variants (#657).

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_END_TYPE, MSO_LINE_END_WIDTH, MSO_LINE_END_LENGTH

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# autoshape with adjustments (rounded rectangle corner radius)
rect = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(1), Inches(1), Inches(3), Inches(1),
)
rect.adjustments[0] = 0.25

# line autoshape (fork addition: MSO_SHAPE.LINE is recognized)
line = slide.shapes.add_shape(MSO_SHAPE.LINE,
                              Inches(1), Inches(3), Inches(2), Inches(0))

# elbow connector with an adjustment
conn = slide.shapes.add_connector(
    MSO_CONNECTOR.ELBOW,
    Inches(1), Inches(4), Inches(5), Inches(5),
)
conn.adjustments[0] = 0.25  # bend position

# arrow heads on the connector line
conn.line.end_arrow.type = MSO_LINE_END_TYPE.TRIANGLE
conn.line.end_arrow.width = MSO_LINE_END_WIDTH.LARGE
conn.line.end_arrow.length = MSO_LINE_END_LENGTH.LARGE

prs.save("out.pptx")
```

- `Shape.auto_shape_type` / `Shape.adjustments` — Autoshape type and parameter adjustments (returns `None` for unknown `prst` `[Added in 1.0.2.dev0]`).
- `MSO_SHAPE.LINE` — Enum member for `prst="line"`. `[Added in 1.0.2.dev0]`
- `Connector.adjustments` — `ConnectorAdjustmentCollection` (indexable `float`). `[Added in 1.0.2.dev0]`
- `Connector.begin_x` / `.begin_y` / `.end_x` / `.end_y` — Endpoint geometry (read/write).
- `Connector.begin_connect(shape, cxn_pt_idx)` / `Connector.end_connect(shape, cxn_pt_idx)` — Snap to a connection point.
- `Connector.line` — `LineFormat` proxy.
- `LineFormat.begin_arrow` / `LineFormat.end_arrow` — `LineEndFormat` with `.type`, `.width`, `.length`. `[Added in 1.0.2.dev0]`
- `LineFormat.color` / `.fill` / `.dash_style` / `.width` — Line formatting.
- `LineEndFormat.type` — `MSO_LINE_END_TYPE` (NONE / TRIANGLE / STEALTH / DIAMOND / OVAL / OPEN). `[Added in 1.0.2.dev0]`
- `LineEndFormat.width` — `MSO_LINE_END_WIDTH` (SMALL / MEDIUM / LARGE). `[Added in 1.0.2.dev0]`
- `LineEndFormat.length` — `MSO_LINE_END_LENGTH` (SMALL / MEDIUM / LARGE). `[Added in 1.0.2.dev0]`

---

## Pictures

`SlideShapes.add_picture()` embeds a new picture; the fork adds
`add_picture_link()` for external-URL references, `Picture.replace_image()`
for swapping the image bytes while keeping position / size / crop, MPO image
detection, and SVG input that raises the dedicated
`UnsupportedImageTypeError` instead of Pillow's opaque exception.

```python
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# embed (use any small PNG bytes)
png = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\rIDATx\x9cc\xf8\xcf"
    b"\xc0\x00\x00\x00\x03\x00\x01\xde\xfb\xc4f\x00\x00\x00\x00IEND\xaeB`\x82"
)
pic = slide.shapes.add_picture(BytesIO(png), Inches(1), Inches(1),
                               Inches(1), Inches(1))

# crop
pic.crop_left = 0.05
pic.crop_right = 0.05

# mask (crop) into an oval
pic.auto_shape_type = MSO_SHAPE.OVAL

prs.save("out.pptx")
```

- `SlideShapes.add_picture(image_file, left, top, width=None, height=None)` — Inline picture (PNG, JPEG, GIF, BMP, TIFF, WMF, EMF, SVG supported; `.mpo` support is `[Added in 2026.05.0]`).
- `SlideShapes.add_picture_link(url, left, top, width=None, height=None)` — Linked (external URL) picture. `[Added in 2026.05.0]`
- `Picture.image` — `Image` wrapper (`.blob`, `.content_type`, `.ext`, `.filename`, `.size`). EMF content-type correction is `[Added in 2026.05.0]`.
- `Picture.replace_image(image_file)` — Swap the picture's bytes while preserving position, size, cropping, masking shape, outline, alt-text, and name. `[Added in 2026.05.0]`
- `Picture.auto_shape_type` — Cropping / masking shape (e.g. `MSO_SHAPE.OVAL`). `[Added in 2026.05.0]`
- `Picture.crop_left` / `.crop_top` / `.crop_right` / `.crop_bottom` — Fractional crop edges.
- `Picture.line` — Outline `LineFormat`.
- `Picture.click_action` — `ActionSetting`.
- SVG images loaded into picture placeholders are supported (error-free SVG loading is `[Added in 2026.05.0]`).

---

## Movies and audio

Movies and audio are added by `SlideShapes.add_movie()` — the media part
is embedded into `ppt/media/`, a `p:pic` is authored with a poster frame,
and a `p:timing` entry is appended (merged in-place if a
`mc:AlternateContent` wrapper already exists).

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

movie = slide.shapes.add_movie(
    "clip.mp4", Inches(1), Inches(1), Inches(4), Inches(3),
    poster_frame_image="poster.png",
    autoplay=True,
)
print(movie.media_type, movie.start_time)

# swap the media bytes without dropping the shape
movie.replace_media("clip-v2.mp4", mime_type="video/mp4")

prs.save("out.pptx")
```

- `SlideShapes.add_movie(movie_file, left, top, width, height, poster_frame_image=None, mime_type=None, autoplay=False)` — Add video or audio shape. `autoplay` is `[Added in 2026.05.0]`.
- `Movie.media_type` — `PP_MEDIA_TYPE.VIDEO` or `PP_MEDIA_TYPE.AUDIO`.
- `Movie.media_format` — Backing `MediaPart` wrapper.
- `Movie.blob` / `Movie.ext` / `Movie.content_type` — Raw media data and metadata. `[Added in 2026.05.0]`
- `Movie.start_time` — Auto-play start time in seconds. `[Added in 2026.05.0]`
- `Movie.replace_media(new_path_or_file, mime_type=None)` — Swap the audio/video binary behind the shape while preserving position, size, cropping, poster frame, hyperlink, and `p:timing` entries. `[Added in 2026.05.0]`
- `Movie.delete()` — Remove the movie and drop its three media-related part relationships. `[Added in 2026.05.0]`
- `Movie.poster_frame_image` — Access the poster-frame image part (read-only).
- `add_movie` tolerates a pre-existing `p:timing` wrapped in `mc:AlternateContent` (merges new `p:video` into the existing `p:childTnLst`). `[Added in 2026.05.0]`

---

## OLE objects and embedded files

Arbitrary binary files can be embedded as OLE objects on a slide. The fork
adds generic `prog_id` / `extension` support (so Word / PDF / ZIP / HTML /
any file can be embedded, not just Excel).

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

ole = slide.shapes.add_ole_object(
    "model.xlsx", prog_id="Excel.Sheet.12",
    left=Inches(1), top=Inches(1),
    width=Inches(2), height=Inches(1),
    icon_file="xls-icon.emf",
)
print(ole.ole_format.prog_id)

prs.save("out.pptx")
```

- `SlideShapes.add_ole_object(object_file, prog_id, left, top, width=None, height=None, icon_file=None, icon_width=None, icon_height=None)` — Embed an OLE payload. Arbitrary `prog_id` + `extension` overloads are `[Added in 2026.05.0]`; `icon_width` / `icon_height` overloads are `[Added in 2026.05.0]`.
- `OleFormat.prog_id` / `OleFormat.show_as_icon` / `OleFormat.blob` — OLE-object metadata and payload.
- `GraphicFrame.has_chart` / `GraphicFrame.has_table` / `GraphicFrame.ole_format` — Graphic-frame introspection.

---

## Text frames and paragraphs

Every shape that can hold text exposes a `text_frame` with paragraphs and
runs. The fork adds `TextFrame.replace_text()`, paragraph and run
`delete()`, `rotation`, math-equation authoring, and per-paragraph
replacement.

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(3))
tf = tb.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.1)

p = tf.paragraphs[0]
p.text = "Hello "
run = p.add_run()
run.text = "world"
run.font.bold = True
run.font.size = Pt(24)

# second paragraph
p2 = tf.add_paragraph()
p2.text = "Second line"
p2.level = 1

# bulk replace across paragraphs
n = tf.replace_text("world", "everyone")

# slide-relative rectangle PowerPoint allocates for rendering text
# (shape bounding-box minus the four TextFrame.margin_* insets)
rect = tb.text_frame_rect
rect.width.inches   # 5.8 == 6.0 - 2 * 0.1" (default l/r inset)

prs.save("out.pptx")
```

- `BaseShape.has_text_frame` / `BaseShape.text_frame` — Text-frame access.
- `BaseShape.text_frame_rect` — `TextFrameRect` namedtuple `(left, top, width, height)` in EMU for the rectangle PowerPoint allocates for rendering text, shape bounding-box shrunk by the four `TextFrame.margin_*` insets. Raises `ValueError` on a shape without a text frame. `[Added in 2026.05.0]`
- `TextFrame.text` — Read/write plain text. `\v` round-trips as a soft line-break. `[Added in 2026.05.0]` for `\v` normalisation.
- `TextFrame.paragraphs` — Tuple of `_Paragraph`.
- `TextFrame.add_paragraph()` — Append a paragraph.
- `TextFrame.clear()` — Remove all paragraphs but keep the first (empty) one.
- `TextFrame.word_wrap` / `TextFrame.auto_size` / `TextFrame.vertical_anchor` / `TextFrame.rotation` — Layout. `rotation` is `[Added in 2026.05.0]`.
- `TextFrame.margin_top` / `margin_bottom` / `margin_left` / `margin_right` — Internal padding.
- `TextFrame.fit_text(font_family, max_size=18, bold=False, italic=False, font_file=None)` — Shrink text until it fits. Guarded against no-fit scenarios. `[Added in 2026.05.0]`
- `TextFrame.font_scale` / `TextFrame.line_space_reduction` — Readbacks for the scaling fit_text applied. `[Added in 2026.05.0]`
- `TextFrame.replace_text(find, replace)` — Bulk search/replace preserving per-run formatting. `[Added in 2026.05.0]`
- `_Paragraph.text` / `.runs` / `.alignment` / `.level` / `.space_before` / `.space_after` / `.line_spacing` — Paragraph formatting.
- `_Paragraph.add_run()` / `.add_line_break()` — Authoring.
- `_Paragraph.add_field(field_type, text="")` — Append a `a:fld` (slide number, date, etc.). `[Added in 2026.05.0]`
- `_Paragraph.add_math_equation(omml_xml)` — Append an OMML equation. `[Added in 2026.05.0]`
- `_Paragraph.delete()` — Remove this paragraph. `[Added in 2026.05.0]`
- `_Paragraph.replace_text(find, replace)` — Per-paragraph search/replace. `[Added in 2026.05.0]`
- `_Paragraph.font` / `.bullet` — Paragraph-level font and bullet-format proxies.
- `_Run.text` / `_Run.font` / `_Run.hyperlink` — Run identity.
- `_Run.delete()` — Remove this run. `[Added in 2026.05.0]`

---

## Fonts and text formatting

`Font` maps the DrawingML character-properties element. The fork adds
effective-color resolution (walks master / theme inheritance),
strikethrough, East-Asian / complex-script font slots, and a hyperlink-
style toggle.

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.lang import MSO_LANGUAGE_ID

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

gf = slide.shapes.add_table(3, 4, Inches(1), Inches(1), Inches(6), Inches(2))
tbl = gf.table

# banding flags
tbl.first_row = True
tbl.horz_banding = True

# apply a built-in style by its GUID ("Medium Style 2 - Accent 1")
tbl.style_id = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

# cell content
tbl.cell(0, 0).text = "Region"
tbl.cell(1, 0).text = "North"

# per-cell row / column index
print(tbl.cell(2, 3).row_idx, tbl.cell(2, 3).col_idx)  # 2, 3

# merge the top-left 2x2 block
tbl.cell(0, 0).merge(tbl.cell(1, 1))

# cell borders
border = tbl.cell(2, 0).border_top
border.width = Pt(1.5)
border.color.rgb = RGBColor(0x44, 0x44, 0x44)

# add a row, drop a column
tbl.rows.add()
tbl.columns[3].delete()

prs.save("out.pptx")
```

- `SlideShapes.add_table(rows, cols, left, top, width, height)` — Returns a `GraphicFrame`; its `.table` attribute is the `Table`.
- `Table.cell(row_idx, col_idx)` / `Table.iter_cells()` — Cell access.
- `Table.rows` / `Table.columns` — `_RowCollection` / `_ColumnCollection`.
- `Table.first_row` / `.first_col` / `.last_row` / `.last_col` / `.horz_banding` / `.vert_banding` — Style-flag toggles.
- `Table.style_id` — Read/write GUID of the built-in PowerPoint table style. `[Added in 1.0.2.dev0]`
- `Table.notify_height_changed()` / `Table.notify_width_changed()` — Size-change notifications.
- `_RowCollection.add()` — Append a row. `[Added in 1.0.2.dev0]`
- `_RowCollection.remove(row)` — Delete a row. `[Added in 1.0.2.dev0]`
- `_ColumnCollection.add(width=None)` — Append a column. `[Added in 1.0.2.dev0]`
- `_ColumnCollection.remove(column)` — Delete a column. `[Added in 1.0.2.dev0]`
- `_Row.height` (read/write) / `_Row.cells` / `_Row.delete()` — Row API. `_Row.delete()` is `[Added in 1.0.2.dev0]`.
- `_Column.width` (read/write) / `_Column.delete()` — Column API. `[Added in 1.0.2.dev0]`
- `_Cell.text` (read/write) / `_Cell.text_frame` — Content.
- `_Cell.row_idx` / `_Cell.col_idx` — Grid coordinates. `[Added in 1.0.2.dev0]`
- `_Cell.fill` — `FillFormat` proxy.
- `_Cell.border_left` / `.border_right` / `.border_top` / `.border_bottom` / `.border_diagonal_down` / `.border_diagonal_up` — Per-cell edge and diagonal borders. `[Added in 1.0.2.dev0]`
- `_Cell.margin_left` / `.margin_right` / `.margin_top` / `.margin_bottom` — Cell padding.
- `_Cell.merge(other_cell)` / `_Cell.split()` / `_Cell.is_merge_origin` / `_Cell.is_spanned` / `_Cell.span_height` / `_Cell.span_width` — Merge handling. `_Cell.merge()` now collapses a full-table-span range (every column or every row) by deleting the redundant rows / columns rather than emitting a merge PowerPoint would silently drop; the surviving row / column absorbs their height / width, and any residual in-axis merge is preserved. `[#636 fix Added in 1.0.2.dev0]`
- `_Cell.vertical_anchor` — `MSO_VERTICAL_ANCHOR`.

---

## Text frames, paragraphs, and runs

Text frames expose paragraphs, each with runs that carry a `Font`. The fork
adds `TextFrame.rotation`, `TextFrame.replace_text()` (matches across runs;
preserves origin-run formatting), `TextFrame.font_scale` and
`line_space_reduction` (the autofit knobs), `_Paragraph.add_math_equation()`,
`_Paragraph.replace_text()`, `_Paragraph.delete()`, `_Run.delete()`,
a full `Font.effective_color` resolver that walks style inheritance (para
`a:defRPr` → body `a:lstStyle` → master `p:txStyles` → theme
`a:clrScheme`), sibling `Font.effective_size` / `.effective_bold` /
`.effective_italic` / `.effective_name` resolvers that walk the same chain
for size / bold / italic / Latin typeface, `Font.strikethrough`,
`Font.highlight_color`,
`Font.use_theme_hyperlink_color`,
`Font.name_ea` / `name_cs` (East-Asian and complex-script slots), and a
complete bullet-format API via `_Paragraph.bullet`.

```python
from pptx import Presentation
from pptx.util import Emu, Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_TEXT_STRIKE_TYPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
run = tb.text_frame.paragraphs[0].add_run()
run.text = "Styled"

font = run.font
font.name = "Calibri"
font.name_ea = "MS Mincho"         # East-Asian slot
font.name_cs = "Arial"             # complex-script slot
font.size = Pt(18)
font.bold = True
font.italic = True
font.strikethrough = MSO_TEXT_STRIKE_TYPE.SINGLE_LINE
font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
font.language_id = MSO_LANGUAGE_ID.ENGLISH_US

# effective color walking the inheritance chain (master / theme)
print(font.effective_color)

# resolved size / bold / italic / typeface via the same chain (issue #378)
print(font.effective_size, font.effective_name,
      font.effective_bold, font.effective_italic)

# build a paragraph
p = tf.paragraphs[0]
r = p.add_run()
r.text = "Hello {NAME}"
r.font.name = "Calibri"
r.font.size = Pt(18)
r.font.bold = True
r.font.strikethrough = MSO_TEXT_STRIKE_TYPE.SINGLE_LINE
r.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

# text shadow (issue #546)
r.font.shadow.blur_radius = Emu(50800)
r.font.shadow.distance = Emu(38100)
r.font.shadow.direction = 45.0
r.font.shadow.color.rgb = RGBColor(0x80, 0x80, 0x80)

# text highlight / background color (issue #675)
r.font.highlight_color.rgb = RGBColor(0xFF, 0xFF, 0x00)
# r.font.clear_highlight_color()  # to remove an explicit highlight

# template-style replacement (cross-run safe, origin-run formatting wins)
tf.replace_text("{NAME}", "World")

# strip autofit hints for a placeholder
# tf.font_scale = 100.0
# tf.line_space_reduction = 0.0

# per-paragraph bullet
p2 = tf.add_paragraph()
p2.text = "Bulleted"
p2.bullet.character(u"•")

# delete a run
# r.delete()

prs.save("out.pptx")
```

- `Font.name` / `Font.size` / `Font.bold` / `Font.italic` / `Font.underline` — Core typography.
- `Font.strikethrough` — `True` / `False` / `MSO_TEXT_STRIKE_TYPE`. `[Added in 2026.05.0]`
- `Font.name_ea` / `Font.name_cs` — East-Asian and complex-script typeface overrides. `[Added in 2026.05.0]`
- `Font.color` — `ColorFormat` proxy.
- `Font.fill` — `FillFormat` for text-body fill (solid, gradient, picture).
- `Font.effect_format` — `EffectFormat` for glow / shadow / reflection. `[Added in 2026.05.0]`
- `Font.shadow` — `ShadowFormat` for this run. `[Added in 2026.05.0]`
- `Font.effective_color` — Resolved RGB, walking paragraph/placeholder/layout/master/theme inheritance. `[Added in 2026.05.0]`
- `Font.effective_size` / `Font.effective_bold` / `Font.effective_italic` / `Font.effective_name` — Resolved size / bold / italic / Latin-typeface for this run, walking the same inheritance chain (run `a:rPr` → paragraph `a:defRPr` → body `a:lstStyle` → master `p:txStyles` → presentation `p:defaultTextStyle`). Returns `None` when no ancestor in the chain declares the property. `[Added in 2026.05.0]`
- `Font.language_id` — `MSO_LANGUAGE_ID` enum.
- `Font.use_theme_hyperlink_color` — When the run wraps a hyperlink, toggle using the theme's hyperlink color. `[Added in 2026.05.0]`
- `ColorFormat.rgb` / `.theme_color` / `.brightness` / `.type` / `.alpha` / `.to_rgb()` — Color type resolution, tint/shade, plus `alpha` and `to_rgb()`. `[Added in 2026.05.0]` for `alpha` and `to_rgb`.
- `RGBColor(r, g, b)` / `RGBColor.from_string("RRGGBB")` — Color value type.
- Enums: `MSO_TEXT_STRIKE_TYPE`, `MSO_TEXT_UNDERLINE_TYPE`, `MSO_LANGUAGE_ID`, `MSO_THEME_COLOR_INDEX`.

---

## Bullets, numbering, and fields

Paragraph-level bullets can be a character, an auto-numbering scheme, or
`none`, with font / size / color overrides. Text fields (`a:fld`) auto-
fill at render time with slide number, date, or presentation title.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_AUTO_NUMBER_SCHEME
from pptx.dml.color import RGBColor

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(3))
tf = tb.text_frame

p1 = tf.paragraphs[0]
p1.text = "First bullet"
p1.bullet.character("•")
p1.bullet.font = "Arial"
p1.bullet.size_pct = 1.0
p1.bullet.color.rgb = RGBColor(0x33, 0x33, 0x33)

p2 = tf.add_paragraph()
p2.text = "Numbered item"
p2.bullet.auto_number(PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD, start_at=1)

# slide-number field
p3 = tf.add_paragraph()
p3.add_field("slidenum")

prs.save("out.pptx")
```

- `_Paragraph.bullet` — `_BulletFormat`. `[Added in 2026.05.0]`
- `_BulletFormat.character(char)` — Set bullet to a literal character.
- `_BulletFormat.auto_number(scheme, start_at=None)` — Set auto-number bullet with a `PP_AUTO_NUMBER_SCHEME`. For a whole-text-frame numbered list, loop `for p in text_frame.paragraphs: p.bullet.auto_number(scheme)` — see the "Numbered lists" recipe in `docs/user/text.rst` (issue #655).
- `_BulletFormat.none()` — Remove the bullet.
- `_BulletFormat.clear()` — Clear every bullet property (inherit from layout / master).
- `_BulletFormat.font` — Bullet font family.
- `_BulletFormat.color` — `ColorFormat`.
- `_BulletFormat.size_pct` / `_BulletFormat.size_points` / `_BulletFormat.clear_size()` — Bullet size as a fraction of text or absolute length.
- `_BulletFormat.type` / `.char` / `.number_scheme` / `.start_at` — Read-only inspection of the current bullet.
- `_Paragraph.add_field(field_type, text="")` — Append a text field (`slidenum`, `datetime`, `datetime1`–`datetime13`, `rdatetime`, etc.). `[Added in 2026.05.0]`
- `_Field.field_type` / `_Field.text` / `_Field.font` — Field metadata.
- Enum: `PP_AUTO_NUMBER_SCHEME`.

---

## Hyperlinks

Runs, text fields, and shapes can carry a hyperlink pointing at an external
URL or an internal slide (via a `ppaction://hlinksldjump` action).

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
other = prs.slides.add_slide(prs.slide_layouts[5])
tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.5))
run = tb.text_frame.paragraphs[0].add_run()
run.text = "Visit our site"
run.hyperlink.address = "https://example.com/"

# run-level slide-jump: click the run to jump to another slide
jump_run = tb.text_frame.paragraphs[0].add_run()
jump_run.text = " (see appendix)"
jump_run.hyperlink.target_slide = other

# target another slide with a shape-level click action
from pptx.enum.shapes import MSO_SHAPE
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(1), Inches(3), Inches(2), Inches(1))
shape.click_action.target_slide = prs.slides[0]

prs.save("out.pptx")
```

- `_Run.hyperlink` — `_Hyperlink` proxy.
- `_Hyperlink.address` — Read/write external URL (setting to `None` removes the hyperlink). Font-color writes for hyperlinked runs are `[Added in 2026.05.0]`. For a slide-jump hyperlink, `address` returns `None` — use `target_slide`.
- `_Hyperlink.target_slide` — Read/write `Slide` for a run-level `hlinksldjump` (run the user clicks jumps to the assigned slide). Assigning `None` (or `del run.hyperlink.target_slide`) removes any hyperlink on the run. `[Added in 2026.05.0]`
- `BaseShape.click_action` — `ActionSetting` for click-triggered behaviours (hyperlinks, target-slide, run-program, play-sound).
- `ActionSetting.action` — `PP_ACTION` enum.
- `ActionSetting.hyperlink` — Outer hyperlink access.
- `ActionSetting.target_slide` — Read/write target slide for `hlinksldjump` actions. `[Added in 2026.05.0]`

---

## Fill and line format (DrawingML)

`FillFormat` models the DrawingML fill element (`solid`, `gradient`,
`pattern`, `picture`, `background`). `LineFormat` models the outline with
width, dash style, arrow heads, and fill. The fork adds picture fill
(`blip_fill`), line-ending arrow control, and corrected dash styles.

```python
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR, MSO_LINE_DASH_STYLE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
rect = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1),
)

# solid theme fill with alpha
rect.fill.solid()
rect.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2
rect.fill.fore_color.alpha = 0.75

# outline
rect.line.color.rgb = RGBColor(0x20, 0x20, 0x20)
rect.line.width = Pt(1.5)
rect.line.dash_style = MSO_LINE_DASH_STYLE.DASH

# picture fill for another shape
rect2 = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5), Inches(1), Inches(3), Inches(1),
)
rect2.fill.blip_fill("texture.png")

prs.save("out.pptx")
```

- `FillFormat.solid()` / `.gradient()` / `.patterned()` / `.background()` — Switch fill type.
- `FillFormat.blip_fill(image_file)` — Picture fill. `[Added in 2026.05.0]`
- `FillFormat.fore_color` / `.back_color` / `.pattern` / `.type` — Per-type accessors.
- `_GradientFillFormat.gradient_stops` — Gradient stop collection.
- `_GradientFillFormat.gradient_angle` — Linear-gradient angle in degrees.
- `LineFormat.color` / `.width` / `.fill` — Stroke color, width, and fill overlay.
- `LineFormat.dash_style` — `MSO_LINE_DASH_STYLE`. Fixed `ROUND_DOT` and `UP_DOWN_ARROW` mappings are `[Added in 2026.05.0]`.
- `LineFormat.begin_arrow_head_style` / `.end_arrow_head_style` / `.begin_arrow_head_width` / `.end_arrow_head_width` / `.begin_arrow_head_length` / `.end_arrow_head_length` — Arrow-head authoring. `[Added in 2026.05.0]`
- Enums: `MSO_FILL_TYPE`, `MSO_PATTERN_TYPE`, `MSO_THEME_COLOR`, `MSO_LINE_DASH_STYLE`, `MSO_LINE_FILL_TYPE`, `MSO_ARROWHEAD_STYLE`, `MSO_ARROWHEAD_WIDTH`, `MSO_ARROWHEAD_LENGTH`.

---

## Shadow and effect format

Shadows, glows, and reflections live under `a:effectLst`. The fork adds
shape-level, text-run, and data-label shadow read/write, with an `inherit`
flag for layout-driven defaults.

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
rect = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1),
)
rect.text_frame.text = "With shadow"

sh = rect.shadow
sh.inherit = False
sh.blur_radius = Pt(6)
sh.distance = Pt(3)
sh.angle = 45.0
sh.color.rgb = RGBColor(0x00, 0x00, 0x00)
sh.alpha = 0.4

prs.save("out.pptx")
```

- `BaseShape.shadow` — `ShadowFormat` proxy. `[Added in 2026.05.0]`
- `Font.shadow` / `_BulletFormat.shadow` — Text and bullet shadows. `[Added in 2026.05.0]`
- `ShadowFormat.inherit` — Inherit shadow from layout / master. `[Added in 2026.05.0]`
- `ShadowFormat.blur_radius` / `.distance` / `.angle` / `.color` / `.alpha` / `.clear()` — Outer-shadow controls. `[Added in 2026.05.0]`
- `EffectFormat.shadow` / `.glow` / `.reflection` — `a:effectLst` child structures. `[Added in 2026.05.0]`

---

## Tables

Tables are authored via `SlideShapes.add_table()`. The fork adds cell
borders, row/column mutation, iteration helpers, style application, and
column-width readback.

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
graphic_frame = slide.shapes.add_table(
    rows=2, cols=3, left=Inches(1), top=Inches(1),
    width=Inches(6), height=Inches(2),
)
tbl = graphic_frame.table

tbl.style_id = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"  # Medium Style 2 - Accent 1
tbl.cell(0, 0).text = "Region"
tbl.cell(0, 1).text = "Q1"
tbl.cell(0, 2).text = "Q2"

# extend via add() (fork-era)
tbl.rows.add()
tbl.columns.add(width=Inches(1))

# or the one-liner Table shortcut (fork-era)
tbl.add_row()
tbl.add_column(width=Inches(1))

# merge a range of cells
tbl.cell(1, 1).merge(tbl.cell(1, 2))

prs.save("out.pptx")
```

- `SlideShapes.add_table(rows, cols, left, top, width, height)` — Append a table. Accepts float dimensions. `[Added in 2026.05.0]` for the float overload.
- `Table.rows` / `Table.columns` / `Table.cell(row_idx, col_idx)` / `Table.iter_cells()` — Access helpers.
- `Table.add_row(height=None)` / `Table.add_column(width=None)` — One-liner shortcuts that delegate to `_RowCollection.add` / `_ColumnCollection.add`. `[Added in 2026.05.0]`
- `Table.style_id` — Apply a built-in table style by GUID. `[Added in 2026.05.0]`
- `Table.first_row` / `last_row` / `first_col` / `last_col` / `horz_banding` / `vert_banding` — Style-flag booleans.
- `Table.notify_height_changed()` / `Table.notify_width_changed()` — Hooks to recompute graphic-frame extents after cell mutation.
- `_RowCollection.add(height=None)` — Append a row; height inherits from the last row or defaults to ~370,840 EMU. `[Added in 2026.05.0]`
- `_RowCollection.remove(row)` — Delete a row. `[Added in 2026.05.0]`
- `_ColumnCollection.add(width=None)` — Append a column. `[Added in 2026.05.0]`
- `_ColumnCollection.remove(column)` — Delete a column. `[Added in 2026.05.0]`
- `_Row.height` / `_Row.cells` / `_Row.delete()` — Row-level access. `delete()` is `[Added in 2026.05.0]`.
- `_Column.width` / `_Column.delete()` — Column-level access. `[Added in 2026.05.0]`
- `_Cell.text` / `.text_frame` / `.vertical_anchor` / `.fill` — Core cell props.
- `_Cell.merge(other)` / `_Cell.split()` / `_Cell.is_merge_origin` / `_Cell.is_spanned` / `_Cell.span_height` / `_Cell.span_width` — Cell merging.
- `_Cell.row_idx` / `_Cell.col_idx` — Positional readback. `[Added in 2026.05.0]`
- `_Cell.margin_top` / `.margin_bottom` / `.margin_left` / `.margin_right` — Cell padding.
- `_Cell.border_top` / `.border_bottom` / `.border_left` / `.border_right` / `.border_diagonal_down` / `.border_diagonal_up` — Cell-border `_CellBorder` proxies. `[Added in 2026.05.0]`

- `TextFrame.paragraphs` / `TextFrame.add_paragraph()` / `TextFrame.text` / `TextFrame.clear()`.
- `TextFrame.word_wrap` / `.vertical_anchor` / `.auto_size` / `.margin_left` / `.margin_right` / `.margin_top` / `.margin_bottom` — Body-property knobs.
- `TextFrame.rotation` — Read/write float degrees (rotates text inside the frame, distinct from shape rotation). `[Added in 1.0.2.dev0]`
- `TextFrame.font_scale` — Autofit font-scale percent (reads/writes `a:normAutofit/@fontScale`). `[Added in 1.0.2.dev0]`
- `TextFrame.line_space_reduction` — Autofit line-space reduction percent. `[Added in 1.0.2.dev0]`
- `TextFrame.fit_text(font_family=..., max_size=..., bold=..., italic=..., font_file=...)` — Compute and store autofit hints.
- `TextFrame.replace_text(find, replace)` — Cross-run text replacement preserving origin-run formatting. Returns replacement count. `[Added in 1.0.2.dev0]`
- `_Paragraph.add_run()` / `_Paragraph.add_line_break()` / `_Paragraph.add_field(field_type, text="")` — Content.
- `_Paragraph.add_math_equation(omml_xml)` — Insert a pre-authored OMML equation (see [Math equations](#math-equations)). `[Added in 1.0.2.dev0]`
- `_Paragraph.replace_text(find, replace)` — Paragraph-scoped cross-run replace. `[Added in 1.0.2.dev0]`
- `_Paragraph.delete()` — Remove the paragraph (fresh empty `a:p` inserted when it was the last one). `[Added in 1.0.2.dev0]`
- `_Paragraph.alignment` / `.level` / `.font` / `.line_spacing` / `.space_before` / `.space_after` / `.text` — Paragraph formatting.
- `_Paragraph.bullet` — `_BulletFormat` with `.character(char)` / `.auto_number(scheme, start_at=None)` / `.none()` / `.clear()` / `.font` / `.size_points` / `.size_pct` / `.color`. `[Added in 1.0.2.dev0]`
- `_Run.text` / `_Run.font` / `_Run.hyperlink` — Run properties.
- `_Run.delete()` — Remove the run. `[Added in 1.0.2.dev0]`
- `Font.name` / `.size` / `.bold` / `.italic` / `.underline` / `.language_id` / `.color` / `.fill`.
- `Font.strikethrough` — `bool` / `MSO_TEXT_STRIKE_TYPE` (NONE / SINGLE_LINE / DOUBLE_LINE). `[Added in 1.0.2.dev0]`
- `Font.name_ea` / `Font.name_cs` — East-Asian and complex-script font slots (`a:ea`, `a:cs`). `[Added in 1.0.2.dev0]`
- `Font.effective_color` — Read-only resolver that walks placeholder / master / theme inheritance and returns the rendered `RGBColor`. `[Added in 1.0.2.dev0]`
- `Font.use_theme_hyperlink_color` — Tri-state `bool` toggling `a:uFill`/`uFillTx` on a hyperlink run. `[Added in 1.0.2.dev0]`
- `Font.shadow` — `ShadowFormat` for the run's text-shadow. Full read/write `.inherit` / `.blur_radius` / `.distance` / `.direction` / `.color` API, backed by `a:rPr/a:effectLst/a:outerShdw` (see [Fills, colors, and effects](#fills-colors-and-effects)). Issue #546. `[Added in 1.0.2.dev0]`
- `Font.effect_format` — `EffectFormat` exposing the full `a:effectLst` family (`.shadow` / `.glow` / `.reflection` / `.soft_edge`) on the run's `a:rPr`. `[Added in 1.0.2.dev0]`
- `Font.highlight_color` / `Font.clear_highlight_color()` — Read/write `ColorFormat` for the text-highlight (text-background) swatch PowerPoint exposes on the Home ribbon. Accepts both `.rgb = RGBColor(...)` and `.theme_color = MSO_THEME_COLOR.ACCENT_1`. Backed by `a:rPr/a:highlight`. Issue #675. `[Added in 1.0.2.dev0]`
- `_Hyperlink.address` — Run-level hyperlink URL (setter creates/clears the `a:hlinkClick`). Returns `None` when the run carries a slide-jump instead of a URL.
- `_Hyperlink.target_slide` — Read/write `Slide` for a run-level slide-jump hyperlink (PowerPoint's "Place in This Document"). Setter swaps the run's `a:hlinkClick` for a `hlinksldjump` action and slide relationship; assigning `None` (or `del run.hyperlink.target_slide`) removes any hyperlink on the run. `[Added in 2026.05.0]`

---

## Charts

Charts live on a slide as a graphic frame; each one owns its own
`chart1.xml` part with data, series, axes, plot area, legend, and title.
The fork ships all thirteen `THREE_D_*` chart types, combo charts,
waterfall / treemap / funnel / sunburst / box-whisker / histogram /
pareto / map via the chartEx surface, data-label customisation, and data-
refresh helpers.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 2D column chart
data = CategoryChartData()
data.categories = ["Q1", "Q2", "Q3", "Q4"]
data.add_series("North", (10, 12, 14, 11))
data.add_series("South", (8, 13, 9, 15))

gf = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1), Inches(6), Inches(3),
    data,
)
chart = gf.chart

# overlay a line on the same axes to make a combo chart
line_data = CategoryChartData()
line_data.categories = ["Q1", "Q2", "Q3", "Q4"]
line_data.add_series("Target", (11, 12, 12, 13))
chart.add_plot(XL_CHART_TYPE.LINE, line_data)

# series lines on a stacked column plot (connects segment tops across series)
stacked_data = CategoryChartData()
stacked_data.categories = ["Q1", "Q2", "Q3"]
stacked_data.add_series("2023", (1, 2, 3))
stacked_data.add_series("2024", (2, 3, 4))
stacked_slide = prs.slides.add_slide(prs.slide_layouts[6])
stacked_chart = stacked_slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_STACKED,
    Inches(1), Inches(1), Inches(6), Inches(3),
    stacked_data,
).chart
stacked_plot = stacked_chart.plots[0]
stacked_plot.has_series_lines = True
stacked_plot.series_lines.format.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)

# chart style (plain 1..48 or extended 49..255 via c14:style)
chart.chart_style = 6

# chart title, positioned manually (factor-space, 0..1)
chart.has_title = True
chart.chart_title.text_frame.text = "Results"
chart.chart_title.position = (0.2, 0.02)

# replace data later, preserving number formats and author formulas
chart.replace_data(data)
# chart.replace_data_preserve_formulas(data)

# refresh cached values from the embedded workbook
chart.update_cached_values()

# 3D column chart (fork addition — previously raised NotImplementedError)
data3d = CategoryChartData()
data3d.categories = ["A", "B", "C"]
data3d.add_series("Series 1", (1, 2, 3))
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
slide2.shapes.add_chart(
    XL_CHART_TYPE.THREE_D_COLUMN_CLUSTERED,
    Inches(1), Inches(1), Inches(6), Inches(3),
    data3d,
)

# XY scatter
xy_data = XyChartData()
series = xy_data.add_series("XY")
for i in range(5):
    series.add_data_point(i, i * i)

prs.save("out.pptx")
```

- `SlideShapes.add_chart(chart_type, x, y, cx, cy, chart_data)` — Append a chart. Supports every `XL_CHART_TYPE` variant except a small set of chartex-only types; 3D variants and additional chartex passthrough are `[Added in 1.0.2.dev0]`.
- `Chart.chart_type` / `Chart.has_title` / `Chart.has_legend` — Basic properties.
- `Chart.chart_title` — `ChartTitle` proxy.
- `Chart.chart_style` — Integer style index. Plain 1..48 via `c:style`; extended 49..255 via `c14:style` wrapped in `mc:AlternateContent` `[Added in 1.0.2.dev0]`.
- `Chart.display_blanks_as` — Read/write member of `XL_DISPLAY_BLANKS_AS` (`GAPS` / `ZERO` / `INTERPOLATED`) controlling PowerPoint's "Hidden and Empty Cells" / "Show empty cells as" setting (`c:dispBlanksAs`). Suppress zero-valued bars in a stacked bar chart (issue #859) by assigning `XL_DISPLAY_BLANKS_AS.GAPS`; assigning the XSD default `ZERO` removes the element. `[Added in 1.0.2.dev0]`
- `Chart.series_in_rows` — Read-only `bool | None` reporting whether the chart's source data is laid out with series in rows (Switch Row/Column applied) or in columns (default). Inferred from the shape of the first `c:ser`'s `c:cat` / `c:tx` cell references (issue #828). Returns `None` when the orientation cannot be determined — e.g. XY/scatter / bubble charts (no `c:cat`), inline-literal series, or empty plots. `[Added in 1.0.2.dev0]`
- `Chart.plots` / `Chart.series` / `Chart.category_axis` / `Chart.value_axis` — Chart anatomy.
- `Chart.plot_area` — `PlotArea` proxy for `c:plotArea`; `.format` exposes `ChartFormat` (`.fill`, `.line`, `.shadow`) so the plot-area rectangle can be filled, outlined, or shadowed without dropping into XML. `[Added in 1.0.2.dev0]`
- `Chart.has_secondary_value_axis` / `Chart.secondary_value_axis` — Secondary Y axis. `[Added in 1.0.2.dev0]`
- `Chart.add_plot(chart_type, chart_data)` — Add a second plot to an existing chart (combo charts). `[Added in 1.0.2.dev0]`
- `Chart.apply_template(template)` — Apply a `.crtx` chart template (path, bytes, or file-like) onto an existing chart. `[Added in 1.0.2.dev0]`
- `Chart.clone_to(shapes, x, y, cx, cy)` — Deep-copy the chart onto another slide (same or different presentation) with a fresh distinct embedded workbook. `[Added in 1.0.2.dev0]`
- `Chart.replace_data(chart_data)` — Rewrite data in place, preserving author-set number formats `[Added in 1.0.2.dev0]`, validating chart-data type match (raises `ValueError` for mismatches) `[Added in 1.0.2.dev0]`, and cycling accent colors on added series `[Added in 1.0.2.dev0]`.
- `Chart.replace_data_preserve_formulas(chart_data)` — Data-only refresh that leaves embedded-workbook cells carrying a formula (`<f>`) untouched. `[Added in 1.0.2.dev0]`
- `Chart.update_cached_values()` — Re-read the embedded workbook and rewrite `c:numCache` / `c:strCache`. `[Added in 1.0.2.dev0]`
- `Chart.workbook` — `ChartWorkbook` proxy; exposes `.xlsx_part` and a full workbook read/write path.
- `Chart.has_user_shapes` / `Chart.user_shapes` — Read-only access to the chart's annotation-shapes (`c:userShapes`) drawing: `has_user_shapes` is a non-destructive boolean probe, `user_shapes` returns a `ChartDrawingPart` (or `None`) whose `iter_anchor_elements()` / `anchor_count` enumerate the raw `cdr:relSizeAnchor` / `cdr:absSizeAnchor` anchors. Authoring is deferred; see `docs/dev/analysis/chart-user-shapes.rst`. `[Added in 1.0.2.dev0]`
- `Chart.has_data_table` / `Chart.data_table` — Show or hide the chart *data table* (`c:plotArea/c:dTable`) rendered beneath the plot area. Assigning `has_data_table = True` writes a default `c:dTable` with all four `c:show*` flags on (horizontal and vertical borders, outline, legend keys). The returned `_DataTable` proxy exposes read/write `show_horz_border`, `show_vert_border`, `show_outline`, `show_keys`, and a `format` `ChartFormat` for styling the data-table's fill / line / effect. `chart.data_table` returns `None` when no data table is present. `[Added in 2026.05.0]`
- `Chart.font` — `Font` proxy for chart-space default text properties.
- `ChartTitle.text_frame` / `ChartTitle.format` / `ChartTitle.has_text_frame` — Title body.
- `ChartTitle.position` — `(x, y)` tuple in factor-space `[0, 1]`, or `None` for auto layout. `[Added in 1.0.2.dev0]`
- `Axis.format` / `Axis.tick_labels` / `Axis.major_gridlines` / `Axis.minor_gridlines` / `Axis.axis_title`.
- `Axis.position` — Read/write `XL_AXIS_POSITION`. `[Added in 1.0.2.dev0]`
- `Axis.visible` — Read/write `bool` toggling axis visibility via the `c:delete` element. Writes `val="1"` / `val="0"` explicitly so PowerPoint honors the setting (see issue #852). `[Fixed in 1.0.2.dev0]`
- `TickLabels.rotation` — Read/write clockwise rotation (degrees) of axis tick labels, mapping to `c:txPr/a:bodyPr/@rot`. Accepts `int` or `float`; default `0.0`. `[Added in 1.0.2.dev0]`
- `CategoryAxis.tick_label_skip` / `CategoryAxis.tick_mark_skip` — Read/write `int` (>=1) thinning out how often a category-axis label or major tick is drawn (`c:tickLblSkip/@val` and `c:tickMarkSkip/@val`). `1` (default) draws every category; `2` draws every other; and so on. Assigning `1` removes the backing element; values `<1` raise `ValueError`. `[Added in 1.0.2.dev0]`
- `ValueAxis.crosses` / `.crosses_at` / `.major_unit` / `.minor_unit`.
- `DateAxis.major_unit` / `.minor_unit` — Time-axis spacing. `[Added in 1.0.2.dev0]`
- `Series.values` / `Series.categories` / `Series.name` / `Series.format` / `Series.marker` / `Series.points`.
- `Series.has_error_bars` / `Series.error_bars` / `Series.set_error_bars(type, amount, include, direction)` — Error bars (fixed value / percentage / std deviation / std error). `[Added in 1.0.2.dev0]`
- `Series.trendlines` / `Series.add_trendline(trendline_type, order, period, forward, backward, intercept, display_equation, display_r_squared)` / `Trendline.delete()` — Fitted-curve overlays on a series: linear / logarithmic / polynomial (order 2..6) / power / exponential / moving-average. Each trendline exposes `.trendline_type`, `.order`, `.period`, `.forward`, `.backward`, `.intercept`, `.display_equation`, `.display_r_squared`, `.name`, and `.format` (a `ChartFormat` for fill / line / shadow). `[Added in 1.0.2.dev0]`
- `BarPlot.gap_width` / `BarPlot.overlap` — Bar / column spacing and overlap (integer percentage of bar width).
- `BarPlot.has_series_lines` / `BarPlot.series_lines` — Read/write boolean and `SeriesLines` accessor for series lines (`c:serLines`) on a stacked bar or stacked column plot; `plot.series_lines.format` returns a `ChartFormat` so the connecting line's color, width, and dash style are configured through the familiar `.format.line` API. Setting `has_series_lines = False` removes the element. `[Added in 1.0.2.dev0]`
- `Point.format` / `Point.marker` / `Point.data_label` / `Point.invert_if_negative` — Per-point formatting.
- `DataLabel.text_frame` / `DataLabel.font` / `DataLabel.position` / `DataLabel.show_*`.
- `DataLabel.format` — `ChartFormat` wrapping this single `c:dLbl` with `.fill` / `.line` / `.shadow`. `[Added in 1.0.2.dev0]`
- `DataLabels.text_frame` / `DataLabels.font` / `DataLabels.show_*` / `DataLabels.number_format`.
- `DataLabels.format` — `ChartFormat` wrapping `c:dLbls` (collection-scope fill / line / shadow). `[Added in 1.0.2.dev0]`
- `ChartFormat.fill` / `ChartFormat.line` / `ChartFormat.shadow` — Format block for a chart element. `shadow` is `[Added in 1.0.2.dev0]`.
- `pptx.enum.chart.XL_CHART_TYPE` — Full chart-type enum including 2D, 3D, XY, bubble, radar, doughnut, area, and a chartex passthrough sentinel (`UNSUPPORTED_CHARTEX`) for unknown `cx:` types `[Added in 1.0.2.dev0]`.
- `pptx.enum.chart.XL_ERROR_BAR_TYPE` / `XL_ERROR_BAR_INCLUDE` / `XL_ERROR_BAR_DIRECTION` — Error-bar configuration. `[Added in 1.0.2.dev0]`
- `pptx.enum.chart.XL_DISPLAY_BLANKS_AS` — How blank cells render (`GAPS` / `ZERO` / `INTERPOLATED`), used by `Chart.display_blanks_as`. `[Added in 1.0.2.dev0]`

- `pptx.enum.chart.XL_TRENDLINE_TYPE` — Trendline regression type (LINEAR, LOGARITHMIC, POLYNOMIAL, POWER, EXPONENTIAL, MOVING_AVG). `[Added in 1.0.2.dev0]`
- `pptx.chart.data.CategoryChartData` / `XyChartData` / `BubbleChartData` — Chart-data builders.

---

## Fills, colors, and effects

`FillFormat` handles solid / gradient / pattern / picture / no-fill /
background fills. The fork adds picture fills via `FillFormat.blip_fill()`
and full `ShadowFormat` configuration (`blur_radius`, `distance`,
`direction`, `color` in addition to `.inherit`), plus a universal
`ColorFormat.to_rgb()` resolver with tint / shade and a
`ColorFormat.alpha` accessor for per-color transparency.

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
rect = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(1), Inches(1), Inches(4), Inches(2),
)

# solid fill with 50 % transparency
rect.fill.solid()
rect.fill.fore_color.rgb = RGBColor(0x2E, 0x74, 0xB5)
rect.fill.fore_color.alpha = 0.5

# resolve a color to its actual rendered RGB (universal, handles tint/shade)
print(rect.fill.fore_color.to_rgb())

# picture (blip) fill — any FillFormat tied to a part
# rect.fill.blip_fill("tests/test_files/python-powered.png")

# gradient fill (overwrites the solid fill above)
rect.fill.gradient()
rect.fill.gradient_angle = 45.0

# shadow (full configuration, not just .inherit)
sh = rect.shadow
sh.inherit = False
sh.blur_radius = Pt(4)
sh.distance = Pt(2)
sh.direction = 45.0
sh.color.rgb = RGBColor(0x00, 0x00, 0x00)

# line fill
rect.line.fill.solid()
rect.line.fill.fore_color.rgb = RGBColor(0, 0, 0)
rect.line.width = Emu(12700)

prs.save("out.pptx")
```

- `FillFormat.solid()` / `FillFormat.gradient()` / `FillFormat.patterned()` / `FillFormat.background()` — Fill-type setters.
- `FillFormat.blip_fill(image_file)` — Embed a picture-fill referencing `image_file`. `[Added in 1.0.2.dev0]`
- `FillFormat.fore_color` / `FillFormat.back_color` — `ColorFormat` proxies.
- `FillFormat.gradient_stops` / `FillFormat.gradient_angle` — Gradient controls.
- `FillFormat.pattern` — `MSO_PATTERN_TYPE`.
- `FillFormat.type` — `MSO_FILL_TYPE`.
- `ColorFormat.rgb` / `ColorFormat.theme_color` / `ColorFormat.brightness` — Base color and luminance.
- `ColorFormat.alpha` — Per-color transparency in `[0.0, 1.0]`. `[Added in 1.0.2.dev0]`
- `ColorFormat.to_rgb(theme_colors=None)` — Resolve to the rendered `RGBColor` after tint / shade / theme lookup. `[Added in 1.0.2.dev0]`
- `ColorFormat.type` — `MSO_COLOR_TYPE`.
- `EffectFormat` — `shadow` / `glow` / `reflection` / `soft_edge` accessors.
- `ShadowFormat.inherit` (setter also clears sibling `a:effectRef/@idx`) `[Added in 1.0.2.dev0]`, `.blur_radius` / `.distance` / `.direction` / `.color` — Outer-shadow configuration. All four knobs are `[Added in 1.0.2.dev0]`.
- `GlowFormat.size` / `.color` — Outer glow. `[Added in 1.0.2.dev0]`
- `ReflectionFormat.blur_radius` / `.distance` — Reflection effect. `[Added in 1.0.2.dev0]`
- `SoftEdgeFormat.size` — Soft-edge radius. `[Added in 1.0.2.dev0]`

---

## Animations and transitions

Every slide exposes `Slide.transition` — a `Transition` proxy that reads /
writes the transition type (23 variants including `MORPH`, wrapped in
`mc:AlternateContent` with a `p:fade` fallback for pre-2013 viewers),
duration, advance-on-click, advance-after-time, PowerPoint speed preset
(SLOW / MEDIUM / FAST), `morph_option` granularity
(`byObject` / `byWord` / `byChar`), and per-variant directional accessors
like `wipe_direction`. Shape animations can be authored via
`Shape.set_animation(type, trigger, delay)` for the five common presets
(APPEAR / FADE_IN / FLY_IN / PULSE / FADE_OUT), read back via
`Shape.animation`, and enumerated via `Slide.iter_shape_animations()` and
`Slide.animation_sequence`.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.transition import (
    PP_TRANSITION_TYPE,
    PP_TRANSITION_SPEED,
    PP_TRANSITION_SIDE_DIRECTION,
)
from pptx.enum.animation import MSO_ANIMATION_TYPE, MSO_ANIMATION_TRIGGER

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# slide transition
t = slide.transition
t.type = PP_TRANSITION_TYPE.WIPE
t.wipe_direction = PP_TRANSITION_SIDE_DIRECTION.LEFT
t.speed = PP_TRANSITION_SPEED.FAST
t.duration = 750           # milliseconds (via p14:dur)
t.advance_on_click = True
t.advance_after_time = 5000   # also advance automatically after 5 s

# MORPH transition (wrapped in mc:AlternateContent with p:fade fallback)
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
slide2.transition.type = PP_TRANSITION_TYPE.MORPH
slide2.transition.morph_option = "byObject"

# author an entrance animation on a shape
rect = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(1), Inches(1), Inches(3), Inches(1),
)
rect.set_animation(
    MSO_ANIMATION_TYPE.FADE_IN,
    trigger=MSO_ANIMATION_TRIGGER.ON_CLICK,
    delay=500,
)

# read existing animation
for anim in slide.iter_shape_animations():
    print(anim.shape_id, anim.effect_type, anim.delay_ms, anim.duration_ms)

for effect in slide.animation_sequence:
    print(effect.shape_id, effect.preset_class, effect.preset_id, effect.delay)

# bulk auto-advance the whole deck
prs.set_auto_advance(seconds=8, advance_on_click=False)

prs.save("out.pptx")
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `Slide.transition` — `Transition` proxy.
- `Transition.type` — `PP_TRANSITION_TYPE` (NONE, CUT, FADE, WIPE, PUSH, COVER, SPLIT, RANDOM_BARS, SHAPE, UNCOVER, WHEEL, MORPH, and ~12 more).
- `Transition.duration` — Milliseconds via `p14:dur`, or `None`.
- `Transition.speed` — `PP_TRANSITION_SPEED` (SLOW / MEDIUM / FAST) mapping `@spd`.
- `Transition.advance_on_click` — `bool`.
- `Transition.advance_after_time` — Milliseconds via `@advTm`, or `None`.
- `Transition.morph_option` — `"byObject"` / `"byWord"` / `"byChar"` (only meaningful when `type == MORPH`).
- `Transition.wipe_direction` — `PP_TRANSITION_SIDE_DIRECTION` on `p:wipe`/`p:push`.
- `BaseShape.animation` — Read-only `AnimationEffect` proxy.
- `BaseShape.set_animation(effect_type, trigger="onClick", delay=0)` — Author a preset effect.
- `AnimationEffect.type` — `MSO_ANIMATION_TYPE` (APPEAR, FADE_IN, FLY_IN, PULSE, FADE_OUT).
- `AnimationEffect.trigger` — `MSO_ANIMATION_TRIGGER` (ON_CLICK, AFTER_PREVIOUS).
- `AnimationEffect.delay` — Milliseconds.
- `AnimationEffect.shape_id` — Owning shape ID.
- `Slide.iter_shape_animations()` — Read-only iterator of `ShapeAnimation` proxies.
- `ShapeAnimation.shape_id` / `.effect_type` / `.delay_ms` / `.duration_ms` / `.element`.
- `Slide.animation_sequence` — Tuple of `AnimationEffectView` for the main sequence.
- `AnimationEffectView.shape_id` / `.preset_class` / `.preset_id` / `.preset_subtype` / `.delay`.
- `Slide.has_animations` / `Slide.timing_xml` — Round-trip introspection.
- `Presentation.set_auto_advance(seconds, advance_on_click=False)` — Bulk-set `advance_after_time` + `advance_on_click` across every slide.
- `pptx.enum.animation.MSO_ANIMATION_TYPE` / `MSO_ANIMATION_TRIGGER`.
- `pptx.enum.transition.PP_TRANSITION_TYPE` / `PP_TRANSITION_SPEED` / `PP_TRANSITION_SIDE_DIRECTION`.

---

## Comments

Legacy PowerPoint comments (`p:cmLst`) are readable and writable. Each
`Slide` exposes `Slide.has_comments` and `Slide.comments`; the collection
supports `add_comment()` with automatic `CommentAuthor` bookkeeping.

```python
import datetime as dt

from pptx import Presentation
from pptx.comments import CommentAuthors

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# register the author once at the package level, then anchor a comment to it
authors_part = prs.part.package.get_or_add_comment_authors_part()
authors = CommentAuthors(authors_part)
author = authors.get_or_add(name="Ben", initials="BH")

slide.comments.add_comment(
    author=author,
    text="Reviewer note",
    position=(1000, 1000),                # EMU-scaled (x, y)
    datetime_value=dt.datetime(2026, 5, 2, 10, 0, 0),
)

for comment in slide.comments:
    who = comment.author
    print(who.name if who else "?", comment.text, comment.datetime)

prs.save("out.pptx")
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `Slide.comments` — `Comments` collection (iterable, indexable, `len()`).
- `Slide.has_comments` — `True` when a `CommentsPart` is attached.
- `Comments.add_comment(author, text, position=(0, 0), datetime_value=None)` — Append a comment anchored to a pre-registered `CommentAuthor`.
- `Comment.text` (read/write) / `Comment.author` / `Comment.author_id` / `Comment.idx` / `Comment.datetime` (read/write) / `Comment.position` (read/write `(x, y)` EMU).
- `CommentAuthor.id` / `.name` / `.initials`.
- `CommentAuthors.add_author(name, initials)` / `CommentAuthors.get_by_id(id)` / `CommentAuthors.get_or_add(name, initials)` — Access via `slide.part.comment_authors`.

---

## Math equations

OMML equations are read via `Shape.has_math_equation` and
`Shape.math_equation_xml`, and written into a text-frame paragraph via
`_Paragraph.add_math_equation(omml_xml)` which wraps the fragment in the
`mc:AlternateContent/mc:Choice[Requires="a14"]/a14:m` scaffolding PowerPoint
emits for an inline equation (with an `mc:Fallback/a:r` run carrying the
OMML reduced to its visible text). The caller is responsible for producing
the OMML; conversion between OMML / MathML / LaTeX is out of scope.

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2))
p = tb.text_frame.paragraphs[0]
r = p.add_run()
r.text = "Pythagoras: "

# literal OMML fragment (in practice, often produced by MML2OMML.XSL)
omml_xml = (
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:r><m:t>a^2 + b^2 = c^2</m:t></m:r>'
    '</m:oMath>'
)
p.add_math_equation(omml_xml)

# read it back off any shape
for shape in slide.shapes:
    if shape.has_math_equation:
        print(shape.math_equation_xml[:80], "...")

prs.save("out.pptx")
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `_Paragraph.add_math_equation(omml_xml)` — Append an inline equation wrapped in the `a14:m` / `mc:Fallback` scaffolding PowerPoint uses.
- `BaseShape.has_math_equation` — `True` when the shape contains an `a14:m` equation (descends into `mc:AlternateContent/mc:Choice`).
- `BaseShape.math_equation_xml` — Raw OMML fragment (`m:oMath` element text) or `None`.

---

## SmartArt

SmartArt is preserved on round-trip — the four diagram parts
(`diagramData`, `diagramLayout`, `diagramColors`, `diagramQuickStyle`) are
lifted across save / reload automatically. A `GraphicFrame` that contains a
diagram reports `has_smart_art == True` and exposes the diagram via
`.smart_art`; the four XML parts are readable as raw bytes. Structured
node-tree authoring is intentionally out of scope for the MVP.

```python
from pptx import Presentation

# open a deck containing SmartArt
# prs = Presentation("with-smartart.pptx")
prs = Presentation()  # (no SmartArt in a fresh deck)

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_smart_art:
            diagram = shape.smart_art
            print(len(diagram.data_xml), "bytes of diagramData")
            # diagram.data_xml / layout_xml / colors_xml / quick_style_xml
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `GraphicFrame.has_smart_art` — `True` for SmartArt diagrams.
- `GraphicFrame.smart_art` — `SmartArt` proxy.
- `SmartArt.data_xml` / `.layout_xml` / `.colors_xml` / `.quick_style_xml` — Raw bytes of the four diagram parts.
- `MSO_SHAPE_TYPE.IGX_GRAPHIC` — `shape_type` discriminator for SmartArt.

---

## OLE embedding

`SlideShapes.add_ole_object()` embeds arbitrary OLE / package files.
Upstream already supported common Office progIds (Excel, Word); the fork
broadens this to accept any `prog_id` plus an explicit `extension` so
arbitrary ZIP / PDF / HTML / DOC payloads can be embedded as clickable
icons.

```python
import io
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import PROG_ID

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# minimal payload bytes
pdf_bytes = io.BytesIO(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
icon_png = io.BytesIO(
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\rIDATx\x9cc\xf8\xcf"
    b"\xc0\x00\x00\x00\x03\x00\x01\xde\xfb\xc4f\x00\x00\x00\x00IEND\xaeB`\x82"
)

# convenience: well-known progIds from PROG_ID enum
# (these include XLSX, DOCX, PDF, ...)

# generic path: any prog_id string + extension
slide.shapes.add_ole_object(
    pdf_bytes,
    prog_id="Package",
    left=Inches(1), top=Inches(1),
    width=Inches(1), height=Inches(1),
    icon_file=icon_png,
    icon_width=Inches(1), icon_height=Inches(1),
    extension="pdf",
)

prs.save("out.pptx")
```

- `SlideShapes.add_ole_object(object_file, prog_id, left, top, width=None, height=None, icon_file=None, icon_width=None, icon_height=None, extension=None)` — Embed an OLE / package object with a custom icon. Accepts well-known `PROG_ID` enum members or arbitrary strings; when `prog_id` is not a known member, `extension` becomes required and the package content-type is emitted. `[Added in 1.0.2.dev0]` for the arbitrary `prog_id` / `extension` path.
- `pptx.enum.shapes.PROG_ID` — Convenience enum of commonly-embedded progIds (XLSX, DOCX, PDF, HTML, ...). Extended set is `[Added in 1.0.2.dev0]`.
- `GraphicFrame.has_chart` / `GraphicFrame.has_table` / `GraphicFrame.has_smart_art` — Type probes.
- OLE objects produced by `add_ole_object()` appear in `shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT`.

---

## Hyperlinks and click actions

Shapes and runs support click-actions and hyperlinks. `ActionSetting`
covers both a plain hyperlink and the PowerPoint-specific "click actions"
(jump to slide, run program, play sound, etc.). The fork adds
`ActionSetting.screen_tip` (hover tooltip), `set_sound()` /
`remove_sound()` / `sound`, jump-to-named-slide targets for runs,
`BaseShape.hover_action` (the mouse-over counterpart of `click_action`,
backed by `a:hlinkMouseOver`), and `Font.use_theme_hyperlink_color` to
override the hyperlink color.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.action import PP_ACTION
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
slide2 = prs.slides.add_slide(prs.slide_layouts[6])

rect = slide1.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(1), Inches(1), Inches(2), Inches(1),
)

# hyperlink
rect.click_action.hyperlink.address = "https://example.com/"
rect.click_action.screen_tip = "Open example.com"

# jump-to-slide action
rect2 = slide1.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(4), Inches(1), Inches(2), Inches(1),
)
rect2.click_action.target_slide = slide2

# mouse-over action: hyperlink that fires as the pointer passes over the shape
rect3 = slide1.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(1), Inches(3), Inches(2), Inches(1),
)
rect3.hover_action.hyperlink.address = "https://example.net/"
rect3.hover_action.screen_tip = "Hover to open"

# run-level hyperlink via a run's font / hyperlink
tb = slide1.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(0.5))
p = tb.text_frame.paragraphs[0]
run = p.add_run()
run.text = "jump"
run.hyperlink.address = "https://example.org/"
run.font.use_theme_hyperlink_color = False  # keep the run's own color

prs.save("out.pptx")
```

- `BaseShape.click_action` — `ActionSetting`.
- `BaseShape.hover_action` — `ActionSetting` bound to `a:hlinkMouseOver`; parallel surface (`.action`, `.hyperlink`, `.target_slide`, `.screen_tip`, `.set_sound()` / `.remove_sound()` / `.sound`) for mouse-hover behaviors. `[Added in 1.0.2.dev0]`
- `ActionSetting.hyperlink` — `Hyperlink` proxy with `.address` (read/write URL).
- `ActionSetting.action` — `PP_ACTION` (HYPERLINK, FIRST_SLIDE, NEXT_SLIDE, PREVIOUS_SLIDE, LAST_SLIDE, NAMED_SLIDE, END_SHOW, RUN_PROGRAM, ...).
- `ActionSetting.target_slide` — Read/write `Slide` (auto-sets `action` to `NAMED_SLIDE`).
- `ActionSetting.screen_tip` — Read/write hover-tooltip text. `[Added in 1.0.2.dev0]`
- `ActionSetting.set_sound(sound_file, filename=None)` / `ActionSetting.remove_sound()` / `ActionSetting.sound` — Click-action sound attachment. `[Added in 1.0.2.dev0]`
- `Sound.blob` / `Sound.name` / `Sound.rId` — Embedded click-sound payload. `[Added in 1.0.2.dev0]`
- `_Run.hyperlink` — Run-scoped hyperlink with `.address`.
- `Font.use_theme_hyperlink_color` — Tri-state `bool` to toggle the theme hyperlink color override. `[Added in 1.0.2.dev0]`
- `ThemePart.theme.hlink_color` / `ThemePart.theme.folHlink_color` — Theme hyperlink / followed-hyperlink colors (read-only). `[Added in 1.0.2.dev0]`

---

## Headers, footers, slide numbers, and date fields

Masters, layouts, and notes masters expose `header_footer` so the default
footer / slide-number / date visibility can be set centrally; per-slide
footer text is authored by inserting a paragraph-level `_Field` whose
`field_type` encodes an auto-refresh value (slide number, date/time).

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()

# central defaults on the master
hf = prs.slide_masters[0].header_footer
hf.footer_visible = True
hf.slide_number_visible = True
hf.date_visible = True

# per-slide footer text
slide = prs.slides.add_slide(prs.slide_layouts[0])

# insert an auto-refresh slide-number field into a textbox on the slide
tb = slide.shapes.add_textbox(Inches(9), Inches(7), Inches(1), Inches(0.3))
p = tb.text_frame.paragraphs[0]
fld = p.add_field("slidenum", text="#")

prs.save("out.pptx")
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `SlideMaster.header_footer` / `SlideLayout.header_footer` / `NotesMaster.header_footer` — `_HeaderFooter` proxy.
- `_HeaderFooter.slide_number_visible` / `.footer_visible` / `.header_visible` / `.date_visible` — Tri-state toggles mapping `p:hf`.
- `_Paragraph.add_field(field_type, text="")` — Insert an auto-refresh `a:fld` (e.g. `slidenum`, `datetime`, `datetime1`...`datetime13`).
- `_Field.field_type` / `_Field.text` / `_Field.font` — Field accessors.

---

## Placeholders

The generic `SlidePlaceholder` — returned for content, body, and object
placeholders — now supports the full `insert_chart()` /
`insert_picture()` / `insert_table()` rich-content insertion API. The
specialized `ChartPlaceholder` / `PicturePlaceholder` / `TablePlaceholder`
subclasses still exist for backwards compatibility. `PicturePlaceholder`
now takes a `crop=False` keyword to fit the image inside the placeholder
bounds without cropping. Placeholders are looked up by `idx` on
`slide.placeholders`.

```python
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

chart_data = CategoryChartData()
chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
chart_data.add_series("North", (10, 12, 14, 11))
chart_data.add_series("South", (8, 13, 9, 15))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1), Inches(6), Inches(4),
    chart_data,
).chart

chart.has_title = True
chart.chart_title.text_frame.text = "Quarterly results"
chart.has_legend = True

# add a secondary plot for a combo chart
chart.add_plot(XL_CHART_TYPE.LINE, chart_data)

prs.save("out.pptx")
```

- `SlideShapes.add_chart(chart_type, x, y, cx, cy, chart_data)` — Append a new chart. Every 3-D chart-type writer is `[Added in 2026.05.0]`.
- `SlideShapes.clone_chart(chart, x, y, cx=None, cy=None)` — Clone a chart from any slide or presentation into this slide, including its embedded workbook. `[Added in 2026.05.0]`
- `GraphicFrame.has_chart` / `GraphicFrame.chart` — Predicate and accessor.
- `Chart.chart_type` — `XL_CHART_TYPE` enum value.
- `Chart.chart_title` / `Chart.has_title` — Title access. Manual title positioning (`ChartTitle.position`) is `[Added in 2026.05.0]`.
- `Chart.category_axis` / `Chart.value_axis` / `Chart.date_axis` — Axis accessors.
- `Chart.plots` — Sequence of `_BasePlot` members.
- `Chart.add_plot(chart_type, chart_data)` — Add a secondary plot (combo chart). `[Added in 2026.05.0]`
- `Chart.series` — Flat `SeriesCollection` across every plot.
- `Chart.has_legend` / `Chart.legend` — Legend toggle and proxy.
- `Chart.chart_style` — Built-in chart style id (1–48).
- `Chart.display_blanks_as` — `XL_DISPLAY_BLANKS_AS`. `[Added in 2026.05.0]`
- `Chart.font` — Default font for every chart-text element. `[Added in 2026.05.0]`
- `Chart.plot_area` — `PlotArea` proxy with `.format` (fill / line) and `.layout`. `[Added in 2026.05.0]`
- `Chart.replace_data(chart_data)` — Rewrite all data in place, regenerating the backing workbook. `formatCode` preservation is `[Added in 2026.05.0]`.
- `Chart.replace_data_preserve_formulas(chart_data)` — Rewrite the workbook cell values without overwriting Excel-entered formulas. `[Added in 2026.05.0]`
- `Chart.apply_template(crtx_file)` — Apply a PowerPoint `.crtx` chart template. `[Added in 2026.05.0]`
- `Chart.clone_to(shapes, x, y, cx=None, cy=None)` — Deep-copy this chart into another slide. `[Added in 2026.05.0]`
- `Chart.user_shapes` — Read access to annotation shapes on the chart drawing part. `[Added in 2026.05.0]`
- Secondary-value-axis charts. `[Added in 2026.05.0]`
- Enum: `XL_CHART_TYPE` (bar / column / line / pie / area / radar / doughnut / XY scatter / bubble / stock / surface / cone / pyramid / 3-D variants / chartEx: waterfall / treemap / funnel / sunburst / box-whisker / histogram / pareto / map).

---

## Chart data

`CategoryChartData`, `XyChartData`, and `BubbleChartData` build the data
model the chart parts need. Categories can be grouped into multi-level
hierarchies, number formats can be applied per series or per category, and
XY / bubble point data can be supplied as tuples.

```python
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData

# hierarchical categories
cd = CategoryChartData()
ym = cd.categories.add_category("2026")
ym.add_sub_category("Q1")
ym.add_sub_category("Q2")
cd.add_series("Sales", (10, 12))

# scatter
xd = XyChartData()
xs = xd.add_series("Points")
xs.add_data_point(1.0, 2.5)
xs.add_data_point(2.0, 3.5)

# bubble
bd = BubbleChartData()
bs = bd.add_series("Accounts")
bs.add_data_point(1.0, 2.0, size=3.0)
```

- `CategoryChartData.categories` — `Categories` collection.
- `CategoryChartData.add_series(name, values=(), number_format=None)` — Append a category series.
- `CategoryChartData.add_category(label)` / `Categories.add_category(label)` — Single-level or group category.
- `Category.add_sub_category(label)` — Multi-level (grouped) category.
- `Categories.number_format` / `Categories.depth` / `Categories.leaf_count` — Category metadata.
- `Category.label` / `Category.depth` / `Category.idx` / `Category.numeric_str_val(date_1904=False)` — Category access.
- `XyChartData.add_series(name, number_format=None)` — Append an XY series.
- `XySeriesData.add_data_point(x, y)` — XY point.
- `BubbleChartData.add_series(name, number_format=None)` — Bubble series.
- `BubbleSeriesData.add_data_point(x, y, size)` — Bubble point with size.
- `CategoryChartData.add_series(..., values=(...))` tolerates `None` for missing labels / values. `[Added in 2026.05.0]` for the tolerance.
- `ChartData` — Legacy alias for `CategoryChartData`.

---

## Chart axes and legend

Both category and value axes expose major / minor tick settings, tick
labels, gridlines, axis titles, and (fork-era) `visible`, custom number
formats, date-axis major/minor unit, and axis position.

```python
from pptx.util import Pt
from pptx.enum.chart import XL_TICK_MARK

chart = ...  # a Chart
cat_axis = chart.category_axis
cat_axis.visible = True
cat_axis.has_major_gridlines = False
cat_axis.minor_tick_mark = XL_TICK_MARK.NONE

val_axis = chart.value_axis
val_axis.minimum_scale = 0
val_axis.maximum_scale = 100
val_axis.major_unit = 20

tl = val_axis.tick_labels
tl.font.size = Pt(9)
tl.number_format = "$#,##0"
tl.rotation = -45
```

- `Chart.category_axis` / `Chart.value_axis` / `Chart.date_axis` — Axis accessors.
- `_BaseAxis.visible` — Show/hide the axis. `[Added in 2026.05.0]`
- `_BaseAxis.has_title` / `_BaseAxis.axis_title` — `AxisTitle` with `.text_frame`, `.format`, `.has_text_frame`.
- `_BaseAxis.has_major_gridlines` / `.has_minor_gridlines` / `.major_gridlines` / `.minor_gridlines` — Gridline toggles and `MajorGridlines` proxy.
- `_BaseAxis.major_tick_mark` / `.minor_tick_mark` / `.tick_labels` — Tick-mark positions and `TickLabels`.
- `_BaseAxis.category_type` / `.crosses` / `.crosses_at` — Axis positioning.
- `_BaseAxis.scale_minimum` / `.scale_maximum` / `.minimum_scale` / `.maximum_scale` — Scale bounds.
- `_BaseAxis.reverse_order` — `True` draws the axis in reverse. `[Added in 2026.05.0]`
- `_BaseAxis.format` — `ChartFormat` (fill / line).
- `_BaseAxis.position` — `XL_AXIS_POSITION`. `[Added in 2026.05.0]`
- `ValueAxis.major_unit` / `.minor_unit` — Tick spacing.
- `DateAxis.major_unit` / `.minor_unit` — Date-axis tick spacing in the axis's time base. `[Added in 2026.05.0]`
- `TickLabels.font` / `.number_format` / `.number_format_is_linked` / `.offset` / `.rotation` — Tick-label formatting. `rotation` is `[Added in 2026.05.0]`.
- `Chart.has_legend` / `Chart.legend` — Legend access.
- `Legend.position` / `.include_in_layout` / `.font` / `.horz_offset`.
- `MajorGridlines.format` — `ChartFormat`.

---

## Chart series, points, and data labels

Every series exposes per-point customisation, line / marker format, data
labels (series-level and per-point), error bars, and trendlines.

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_TRENDLINE_TYPE

chart = ...  # a Chart
series = chart.series[0]

# series-level formatting
series.format.line.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

# series-level data labels
labels = series.data_labels
labels.show_value = True
labels.number_format = "0.0"
labels.font.size = Pt(10)
labels.format.fill.solid()
labels.format.fill.fore_color.rgb = RGBColor(0xFF, 0xF5, 0x00)

# customize a single point
pt = series.points[2]
pt.data_label.text_frame.text = "peak"
pt.format.fill.solid()
pt.format.fill.fore_color.rgb = RGBColor(0xCC, 0x33, 0x33)

# error bars + trendline
series.has_error_bars = True
series.error_bars.include_negative = False
series.add_trendline(XL_TRENDLINE_TYPE.LINEAR, display_r_squared=True)
```

- `Chart.series` — Flat `SeriesCollection` across plots.
- `_BaseSeries.name` / `.values` / `.format` / `.index` — Core accessors.
- `_BaseSeries.points` — Sequence of `Point`.
- `_BaseSeries.data_labels` — Series-level `DataLabels`.
- `_BaseSeries.invert_if_negative` — Toggle mirror-fill for negative values. `[Added in 2026.05.0]`
- `DataLabels.format` — `ChartFormat` for series data labels (fill / line / shadow). `[Added in 2026.05.0]`
- `_BaseCategorySeries.categories` — Categories for this series.
- `BarSeries.fill` / `.line` / `.invert_if_negative` — Bar-specific formatting.
- `LineSeries.smooth` — Smoothed-line flag.
- `_MarkerMixin.marker` — Marker (style / size / format) for line / scatter / radar series.
- `XySeries.iter_values()` — Iterate over (x, y) pairs.
- `BubbleSeries.bubble_sizes` — Iterate over bubble magnitudes.
- `_BaseSeries.add_trendline(trendline_type, order=None, forward=None, backward=None, display_r_squared=False, display_equation=False)` — Append a trendline. `[Added in 2026.05.0]`
- `_BaseSeries.has_error_bars` / `_BaseSeries.error_bars` — Error-bar toggle and `ErrorBars` proxy. `[Added in 2026.05.0]`
- `Trendline.type` / `.order` / `.forward` / `.backward` / `.display_equation` / `.display_r_squared`. `[Added in 2026.05.0]`
- `ErrorBars.direction` / `.end_style` / `.include_positive` / `.include_negative` / `.type` / `.value`. `[Added in 2026.05.0]`
- `Point.data_label` — Per-point `DataLabel` (customisable `text_frame`, `position`, `font`, `format.fill`, `format.line`, `format.shadow`). Per-point `format` is `[Added in 2026.05.0]`.
- `DataLabel.format` — `ChartFormat` wrapping the individual `c:dLbl`. `[Added in 2026.05.0]`
- `DataLabels.show_value` / `.show_category_name` / `.show_series_name` / `.show_percentage` / `.show_legend_key` — Content toggles.
- `DataLabels.number_format` / `.number_format_is_linked` / `.position` / `.font` / `.text_frame` — Formatting. `text_frame` is `[Added in 2026.05.0]`.
- `Marker.style` / `.size` / `.format` — `XL_MARKER_STYLE` marker authoring.

---

## Chart templates and cloning

A `.crtx` PowerPoint chart template can be applied to any chart, copying
its colours / effects / axis layout. A chart can also be cloned across
slides, taking its embedded workbook with it.

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation("deck.pptx")
chart = prs.slides[0].shapes[0].chart

# apply a .crtx
chart.apply_template("house-style.crtx")

# deep-copy the chart onto another slide
second = prs.slides[1]
clone = chart.clone_to(second.shapes, Inches(1), Inches(1),
                       cx=Inches(6), cy=Inches(4))

prs.save("out.pptx")
```

- `Chart.apply_template(crtx_file)` — Apply a `.crtx` chart template. `[Added in 2026.05.0]`
- `Chart.clone_to(shapes, x, y, cx=None, cy=None)` — Cross-slide chart copy with its embedded workbook. `[Added in 2026.05.0]`
- `SlideShapes.clone_chart(chart, x, y, cx=None, cy=None)` — Inverse form, invoked on the destination `SlideShapes`. `[Added in 2026.05.0]`
- `Chart.replace_data(chart_data)` — Rewrite all data preserving format codes. `[Added in 2026.05.0]` for formatCode preservation.
- `Chart.replace_data_preserve_formulas(chart_data)` — Rewrite values only; formulas in the embedded workbook survive. `[Added in 2026.05.0]`

---

## SmartArt and equations

SmartArt is read-only: enumerate diagrams on a slide and read their node
text, layout, and underlying data partname. OMML math equations can be
written onto any paragraph and read from any text frame.

```python
from pptx import Presentation

prs = Presentation("with-smartart.pptx")
slide = prs.slides[0]

for shape in slide.shapes:
    if getattr(shape, "has_smart_art", False):
        sa = shape.smart_art
        print(sa.layout_name)
        for node in sa.nodes:
            print("  " * node.level, node.text)

    if shape.has_math_equation:
        print(shape.math_equation_xml)
```

- `BaseShape.has_smart_art` / `BaseShape.smart_art` — SmartArt detection and proxy. `[Added in 2026.05.0]`
- `SmartArt.layout_name` / `.data_partname` / `.nodes` / `.text` — Diagram metadata. `[Added in 2026.05.0]`
- `SmartArtNode.text` / `.level` / `.children` — Hierarchical nodes. `[Added in 2026.05.0]`
- `BaseShape.has_math_equation` / `BaseShape.math_equation_xml` — OMML inspection. `[Added in 2026.05.0]`
- `_Paragraph.add_math_equation(omml_xml)` — Author an OMML expression inline. `[Added in 2026.05.0]`

---

## Action settings and click actions

`click_action` wraps the shape's on-click behaviour: hyperlink, jump to
slide, run program, launch OLE object, play sound. The fork adds first-
class `target_slide`, `screen_tip`, and embedded-sound support.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide1 = prs.slides.add_slide(prs.slide_layouts[5])
slide2 = prs.slides.add_slide(prs.slide_layouts[5])

shape = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(1), Inches(1), Inches(2), Inches(1))
action = shape.click_action

# jump to another slide
action.target_slide = slide2
action.screen_tip = "Go to slide 2"

# play a sound
action.set_sound("chime.wav")
print(action.sound.name)

prs.save("out.pptx")
```

- `BaseShape.click_action` — `ActionSetting` proxy.
- `ActionSetting.action` — `PP_ACTION` enum (read-only).
- `ActionSetting.hyperlink` — `_Hyperlink` proxy.
- `ActionSetting.target_slide` — Read/write slide for `hlinksldjump` actions. `[Added in 2026.05.0]`
- `ActionSetting.screen_tip` — Hover tooltip text. `[Added in 2026.05.0]`
- `ActionSetting.sound` / `ActionSetting.set_sound(path_or_stream)` / `ActionSetting.remove_sound()` — Embedded click-sound authoring. `[Added in 2026.05.0]`
- `Sound.name` / `Sound.blob` — Read-only embedded-sound view. `[Added in 2026.05.0]`
- `pptx.media.Audio` — Value object representing the embedded sound file. `[Added in 2026.05.0]`
- Enum: `PP_ACTION`.

---

## Animations and timing

Slide timing (`p:timing`) encodes entrance, emphasis, exit, and path
animations bound to shapes or text ranges. The fork ships read access
plus a basic authoring API for entrance / exit effects, delays, and
start-conditions.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.animation import MSO_ANIMATION_EFFECT, MSO_ANIMATION_TRIGGER

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(1), Inches(1), Inches(2), Inches(1))
shape.set_animation(
    effect=MSO_ANIMATION_EFFECT.FADE,
    trigger=MSO_ANIMATION_TRIGGER.ON_CLICK,
    delay=0.5,
    duration=1.0,
)

# read back
for eff in slide.animation_sequence:
    print(eff.trigger, eff.effect, eff.target_shape_id)

prs.save("out.pptx")
```

- `Slide.has_animations` / `Slide.timing_xml` — Raw timing XML access. `[Added in 2026.05.0]`
- `Slide.animation_sequence` — Tuple of `AnimationEffectView` in document order. `[Added in 2026.05.0]`
- `Slide.iter_shape_animations()` — Iterate over `ShapeAnimation` entries. `[Added in 2026.05.0]`
- `BaseShape.animation` — Read-only `AnimationEffect` or `None`. `[Added in 2026.05.0]`
- `BaseShape.set_animation(effect, trigger=..., delay=0, duration=None)` — Author an effect on this shape. `[Added in 2026.05.0]`
- `AnimationEffectView.effect` / `.trigger` / `.delay` / `.duration` / `.target_shape_id`. `[Added in 2026.05.0]`
- Enums: `MSO_ANIMATION_EFFECT`, `MSO_ANIMATION_TRIGGER`.

---

## Transitions

Slide transitions control how one slide gives way to the next (fade, push,
wipe, morph, etc.). The fork exposes the full transition surface: `type`,
`speed`, `duration`, `advance_on_click`, `advance_after_time`, and morph-
specific options.

```python
from pptx import Presentation
from pptx.enum.transition import PP_TRANSITION_TYPE, PP_TRANSITION_SPEED

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

t = slide.transition
t.type = PP_TRANSITION_TYPE.MORPH
t.morph_option = "byObject"
t.speed = PP_TRANSITION_SPEED.MEDIUM
t.duration = 800              # milliseconds
t.advance_on_click = True
t.advance_after_time = 10000  # milliseconds

prs.save("out.pptx")
```

- `Slide.transition` — `Transition` proxy. `[Added in 2026.05.0]`
- `Transition.type` — Read/write `PP_TRANSITION_TYPE`. `[Added in 2026.05.0]`
- `Transition.morph_option` — `byObject` / `byWord` / `byChar`. `[Added in 2026.05.0]`
- `Transition.speed` — `PP_TRANSITION_SPEED`. `[Added in 2026.05.0]`
- `Transition.duration` — Duration in ms. `[Added in 2026.05.0]`
- `Transition.advance_on_click` / `Transition.advance_after_time` — Advance behaviour. `[Added in 2026.05.0]`
- `Transition.wipe_direction` / other direction-specific properties (wipe / push / reveal / cover / uncover) — Direction controls. `[Added in 2026.05.0]`
- `Presentation.set_auto_advance(seconds, apply_to="all")` — Shortcut that sets `advance_after_time` on every slide. `[Added in 2026.05.0]`
- Enums: `PP_TRANSITION_TYPE`, `PP_TRANSITION_SPEED`, `PP_TRANSITION_SIDE_DIRECTION`.

---

## Comments (legacy and threaded)

Legacy PowerPoint comments (`p:cmAuthorLst` + per-slide `p:cmLst`) and
Microsoft-365 threaded comments (`threadedComments{N}.xml` + `authors.xml`)
are both supported. `[Added in 2026.05.0]`.

```python
from pptx import Presentation

prs = Presentation("with-comments.pptx")
slide = prs.slides[0]

print("has_comments:", slide.has_comments)
for c in slide.comments:
    print(c.author_name, c.author_initials, c.text)
    for reply in getattr(c, "replies", ()):
        print("  ↳", reply.author_name, reply.text)

# author a new legacy comment
slide.comments.add_comment(author="Ben", initials="BH",
                           text="Tighten this section.")

prs.save("out.pptx")
```

- `Slide.has_comments` / `Slide.comments` — Predicate and `Comments` collection. `[Added in 2026.05.0]`
- `Comments.__iter__` / `__len__` / `Comments.add_comment(author, initials, text, x=None, y=None)` — Enumerate and author comments. `[Added in 2026.05.0]`
- `Comment.author_name` / `Comment.author_initials` / `Comment.text` / `Comment.time` / `Comment.pos`. `[Added in 2026.05.0]`
- `ThreadedComment.author_name` / `.text` / `.created` / `.replies` — 365-era threaded-comment view. `[Added in 2026.05.0]`
- `CommentAuthor.name` / `CommentAuthor.initials` — Per-package author record.

---

## Notes slides and notes master

Every slide can have a notes slide; every presentation has a single notes
master. Notes slides carry text and shapes.

```python
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])

notes = slide.notes_slide
notes.notes_text_frame.text = "Remember to mention the Q1 uptick."

# notes master (fork-era accessor)
nm = prs.notes_master
print(nm.header_footer.slide_number_visible)

prs.save("out.pptx")
```

- `Slide.has_notes_slide` / `Slide.notes_slide` — Notes-slide access.
- `NotesSlide.notes_text_frame` — First text-placeholder body shape's `TextFrame`.
- `NotesSlide.notes_placeholder` — Notes placeholder (`PP_PLACEHOLDER.BODY`), or `None`.
- `NotesSlide.placeholders` / `NotesSlide.shapes` — Full access.
- `NotesSlide.clone_master_placeholders(notes_master)` — Rebuild placeholders from the notes master.
- `Presentation.notes_master` — The notes master. `[Added in 2026.05.0]`
- `NotesMaster.placeholders` / `NotesMaster.shapes` / `NotesMaster.header_footer` — Notes-master access.

---

## Header, footer, slide number, and date

Slide masters, layouts, and the notes master all expose a `_HeaderFooter`
proxy that controls slide-number, header, footer, and date placeholders
via the `p:hf` element.

```python
from pptx import Presentation

prs = Presentation()
master = prs.slide_master
hf = master.header_footer

hf.slide_number_visible = True
hf.footer_visible = True
hf.date_visible = True

# header applies to the notes master only
prs.notes_master.header_footer.header_visible = True

prs.save("out.pptx")
```

- `SlideMaster.header_footer` / `SlideLayout.header_footer` / `NotesMaster.header_footer` — `_HeaderFooter` proxy.
- `_HeaderFooter.slide_number_visible` — Read/write.
- `_HeaderFooter.footer_visible` / `_HeaderFooter.date_visible` — Read/write.
- `_HeaderFooter.header_visible` — Notes-master only.
- `_Paragraph.add_field("slidenum")` / `add_field("datetime")` — Insert a live field into a header / footer / title. `[Added in 2026.05.0]`

---

## Theme and colors

`SlideMaster.theme_colors` resolves theme-color role names to `RGBColor`.
Colors can also be set on any element via `ColorFormat.theme_color` and
retrieved as RGB via `ColorFormat.to_rgb()` — both fork-era additions.

```python
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor

prs = Presentation("branded.pptx")
master = prs.slide_master

print(master.theme_colors)           # {'accent1': RGBColor(...), ...}

# author a shape with a theme-color fill
slide = prs.slides[0]
shape = slide.shapes[0]
shape.fill.solid()
shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1

# resolve to RGB later
print(shape.fill.fore_color.to_rgb())
```

- `SlideMaster.theme_colors` — `{role_name: RGBColor}` mapping. `[Added in 2026.05.0]`
- `ColorFormat.type` — `MSO_COLOR_TYPE` enum.
- `ColorFormat.rgb` — `RGBColor` for direct colors.
- `ColorFormat.theme_color` — `MSO_THEME_COLOR` role name.
- `ColorFormat.brightness` — Tint / shade of a theme color.
- `ColorFormat.alpha` — Per-color transparency (0..1). `[Added in 2026.05.0]`
- `ColorFormat.to_rgb()` — Universal reducer that resolves theme / scheme colors against the master's theme. `[Added in 2026.05.0]`
- Enums: `MSO_COLOR_TYPE`, `MSO_THEME_COLOR` (+ `MSO_THEME_COLOR_INDEX`).

---

## Font embedding

Fonts used in a presentation can be embedded (TTF or OTF) into the
package. `[Added in 2026.05.0]`.

```python
from pptx import Presentation

prs = Presentation()
prs.embed_font(
    "Inter-Regular.ttf", typeface="Inter",
    regular=True, bold=False, italic=False,
)
print(prs.embedded_fonts)   # ("Inter", ...)

prs.save("out.pptx")
```

- `Presentation.embed_font(font_file, typeface, regular=True, bold=False, italic=False, bold_italic=False)` — Embed a TrueType / OpenType font file into the package. `[Added in 2026.05.0]`
- `Presentation.embedded_fonts` — Tuple of typeface names currently embedded. `[Added in 2026.05.0]`
- `FontFile` — Internal part representation.
- `pptx.text.fonts.Font` — Low-level TTF-table reader used by `fit_text`.

---

## View props and presentation metadata

`Presentation.view_props` exposes editor-view state stored in
`viewProps.xml` (normal-view outline / thumbnail split, notes-view zoom,
last-viewed-slide id, etc.). `[Added in 2026.05.0]`.

```python
from pptx import Presentation

prs = Presentation("deck.pptx")
vp = prs.view_props
print(vp.last_view, vp.slide_view_props)
```

- `Presentation.view_props` — `ViewProps` proxy. `[Added in 2026.05.0]`
- `ViewProps.last_view` / `.show_comments` / `.slide_view_props` / `.notes_text_view_props` / `.outline_view_props` / `.sorter_view_props` — View-state fields. `[Added in 2026.05.0]`
- `Presentation.first_slide_num` — Starting slide number. `[Added in 2026.05.0]`

---

## Document properties

Core (Dublin-Core), custom (typed, user-defined), and extended (application)
properties are all exposed.

```python
from pptx import Presentation

prs = Presentation()

cp = prs.core_properties
cp.author = "Ben"
cp.title = "Pitch deck"

prs.custom_properties["ReviewerCount"] = 3
prs.custom_properties["IsDraft"] = True

ep = prs.extended_properties
ep.set("Company", "Example Inc")
print(ep.get("Application"))

prs.save("out.pptx")
```

- `Presentation.core_properties` — `CoreProperties` proxy (author / title / subject / keywords / category / comments / content_status / identifier / language / version / created / last_modified_by / last_printed / modified / revision).
- `Presentation.custom_properties` — `CustomProperties` dict-like typed mapping. `[Added in 2026.05.0]`
- `CustomProperties.__getitem__` / `__setitem__` / `__delitem__` / `__contains__` / `__iter__` / `__len__` / `.add(name, value)` / `.get(name, default=None)` / `.names()` / `.items()` — Full mapping interface (accepts `str`, `int`, `float`, `bool`, `datetime`). `[Added in 2026.05.0]`
- `Presentation.extended_properties` — `ExtendedProperties` proxy for `docProps/app.xml`. `[Added in 2026.05.0]`
- `ExtendedProperties.get(name)` / `ExtendedProperties.set(name, value)` — Generic reads/writes. Typed accessors (`company`, `manager`, `application`, `app_version`, `total_time`, `pages`, `words`, `template`, `presentation_format`, `slides`, `hidden_slides`, `notes`, `mm_clips`, `title_of_parts`, `heading_pairs`, etc.) are generated from a declarative spec. `[Added in 2026.05.0]`

---

## Encryption and protection

`.pptx` packages protected with ECMA-376 Agile Encryption can be read and
written with a password. `[Added in 2026.05.0]`.

```python
from pptx import Presentation

# open an encrypted deck
prs = Presentation("secret.pptx", password="hunter2")

# save re-encrypted
prs.save("secret-out.pptx", password="hunter2")

# or save without encryption
prs.save("plain.pptx")
```

- `pptx.Presentation(pptx, password=None)` — Open a password-protected package. Raises `EncryptedPackageError` on a wrong password. `[Added in 2026.05.0]`
- `Presentation.save(file, password=None)` — Write the package, optionally re-encrypting with `password`. `[Added in 2026.05.0]`
- `pptx.opc._crypto` — Internal Agile-Encryption plumbing. `[Added in 2026.05.0]`
- `pptx.exc.EncryptedPackageError` — Raised when decryption fails. `[Added in 2026.05.0]`

---

## Packaging and I/O options

`Presentation.save()` supports:
- A regular `.pptx` / `.pptm` save (default).
- Slideshow `.ppsx` output via `save_ppsx()`. `[Added in 2026.05.0]`
- Flat-OPC (`<pkg:package>`) single-XML serialisation via
  `save_flat_xml()`. `[Added in 2026.05.0]`
- Agile-encrypted output via `password=` on `save()`.
  `[Added in 2026.05.0]`
- Reproducible builds via fixed zip timestamps.
  `[Added in 2026.05.0]`

Opening supports:
- `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx` packages.
  `.potx` / `.potm` / `.ppsx` template discrimination is
  `[Added in 2026.05.0]`.
- Flat-OPC single-XML input auto-detected.
  `[Added in 2026.05.0]`
- Agile-encrypted input via `password=`.
  `[Added in 2026.05.0]`
- Paper-size / aspect-ratio presets via `pptx_format=` (16:9 default,
  `A4`, `LETTER`, `WIDESCREEN`, `ON_SCREEN_4_3`, etc.).
  `[Added in 2026.05.0]`

```python
from pptx import Presentation

# apply a .potx template to existing content (strip_slides + merge)
template = Presentation("brand.potx").strip_slides()
content = Presentation("content.pptx")
template.merge(content)
template.save("branded.pptx")

# Flat-OPC
template.save_flat_xml("branded.xml")

# slideshow
template.save_ppsx("branded.ppsx")
```

- `Presentation.save(file, password=None)` — Save options as above.
- `Presentation.save_ppsx(file)` — Slideshow content-type. `[Added in 2026.05.0]`
- `Presentation.save_flat_xml(file)` — Flat-OPC. `[Added in 2026.05.0]`
- `Presentation(..., pptx_format=...)` — Paper-size / aspect-ratio preset on create. `[Added in 2026.05.0]`
- `Presentation.strip_slides()` — Remove every slide; keep masters / layouts / theme / embedded fonts. `[Added in 2026.05.0]`
- `pptx.opc.flat_opc` — `write_flat_opc`, `is_flat_opc` helpers.
- Reproducible-build tooling: fixed epoch timestamps on every zip entry. `[Added in 2026.05.0]`
- XXE and zip-bomb hardening across the opening path. `[Added in 2026.05.0]`

---

## Cross-presentation operations

Whole presentations can be merged, slides can be imported one-by-one, and
charts can be cloned across slides or packages.

```python
from pptx import Presentation

a = Presentation("deck-a.pptx")
b = Presentation("deck-b.pptx")

# merge all of b's slides into a, with full media / chart / OLE fidelity
a.merge(b)

# import a single slide from a different deck
a.slides.add_slide_from_external(b.slides[2], a.slide_layouts[5])

# clone a chart between slides
chart = a.slides[0].shapes[0].chart
chart.clone_to(a.slides[1].shapes, 100000, 100000)

a.save("combined.pptx")
```

- `Presentation.merge(other_presentation)` — Append every slide in `other_presentation`, cloning pictures / media / charts (each with its own embedded workbook) / OLE objects / external hyperlinks, without mutating the source deck. Returns the list of new slides. `[Added in 2026.05.0]`
- `Slides.duplicate(slide, index=None)` — Duplicate a slide in this deck. `[Added in 2026.05.0]`
- `Slides.add_slide_from_external(source_slide, slide_layout)` — Import a single slide from another presentation. `[Added in 2026.05.0]`
- `SlideShapes.clone_chart(chart, x, y, cx=None, cy=None)` / `Chart.clone_to(shapes, x, y, cx=None, cy=None)` — Cross-slide chart clone. `[Added in 2026.05.0]`
- `Picture.replace_image(image_file)` — In-place image swap (supports pictures authored on one deck and reused in merges). `[Added in 2026.05.0]`
- `Movie.replace_media(new_path_or_file, mime_type=None)` — In-place media swap. `[Added in 2026.05.0]`

---

## Units and helpers

`pptx.util` provides length and color constructors. Every length value is
a typed `int` subclass carrying an EMU quantum (914,400 per inch) and
supports arithmetic that returns the correct length type.

```python
from pptx.util import Inches, Emu, Pt, Cm, Mm
from pptx.dml.color import RGBColor

print(Inches(1) == Emu(914400))
print(Pt(12).emu, Cm(2.5).emu)
print(RGBColor(0x2E, 0x74, 0xB5))
```

- `pptx.util.Length` — Base class; `.inches`, `.pt`, `.emu`, `.cm`, `.mm`, `.centipoints`.
- `pptx.util.Inches` / `.Emu` / `.Pt` / `.Cm` / `.Mm` / `.Centipoints` — Typed constructors.
- `pptx.dml.color.RGBColor(r, g, b)` / `RGBColor.from_string("RRGGBB")` — Color value with hex output.
- `pptx.enum.*` — Every enumeration (`MSO_SHAPE`, `MSO_CONNECTOR_TYPE`, `PP_PLACEHOLDER`, `XL_CHART_TYPE`, `XL_MARKER_STYLE`, `XL_LEGEND_POSITION`, `XL_TICK_MARK`, `PP_TRANSITION_TYPE`, `MSO_ANIMATION_EFFECT`, etc.).

---

## API concepts

`python-pptx` is organised in three layers:

- **Document API** (`src/pptx/presentation.py`, `src/pptx/slide.py`,
  `src/pptx/shapes/`, `src/pptx/text/`, `src/pptx/chart/`,
  `src/pptx/table.py`) — proxy objects wrapping OOXML elements. This is
  where the overwhelming majority of user code lives.
- **Parts layer** (`src/pptx/parts/*.py`) — `XmlPart` subclasses that own
  the XML trees for each of the package's constituent parts (presentation,
  slide, slide-master, slide-layout, notes, chart, image, media, embedded
  package, comments, threaded comments, tags, core / extended / custom
  props) and manage the relationships between them.
- **oxml layer** (`src/pptx/oxml/*.py`) — `CT_*` classes extending
  `lxml.etree.ElementBase` and mapping directly onto schema element names.

`lxml` handles the XML parsing, serialisation, and XPath work beneath the
library. `pptx.util` carries `Length` subclasses; `pptx.dml.color.RGBColor`
is the colour value type; `ElementProxy` / `ParentedElementProxy` /
`PartElementProxy` / `Subshape` are the proxy base classes.

```python
from pptx import Presentation
from pptx.util import Inches, Emu, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
# any Length is just a typed int — freely interchangeable
cx = Inches(6)
print(cx, cx.emu, Emu(5486400))
print(Pt(24), RGBColor(0x2E, 0x74, 0xB5))
prs.save("out.pptx")
```

- `pptx.util.Length` / `Inches` / `Cm` / `Mm` / `Pt` / `Emu` / `Centipoints` — Length constructors and arithmetic.
- `pptx.dml.color.RGBColor` — `(r, g, b)` triple with `from_string()`, hex output.
- `pptx.shared.ElementProxy` / `ParentedElementProxy` / `PartElementProxy` / `Subshape` — Proxy base classes.
- `pptx.opc.constants.CONTENT_TYPE` / `RELATIONSHIP_TYPE` — Content-type and rel-type constants used by the parts layer.
- `pptx.oxml.ns.qn(tag)` — Clark-notation tag expansion; only needed when dropping into the oxml layer.
- `pptx.oxml.qn` — Re-exported for backwards compatibility. `[Added in 2026.05.0]` restoration.

---

*This file is generated and maintained by hand — see `HISTORY.rst` for the
full change log, `docs/user/*.rst` for narrative tutorials, and
`docs/api/*.rst` for per-class API reference pages.*
