# Features

`loadfix/python-pptx` is a fork of
[python-pptx](https://github.com/scanny/python-pptx) that extends the library
with full-fidelity presentation merge, slide duplication / reorder / delete,
hidden slides, section management, shape lifecycle (delete / duplicate /
replace / flip / z-order), effective (group-transform composited) geometry,
XPath and by-name shape lookup, picture / movie replacement, linked pictures,
combo and 3D charts, secondary-value-axis charts, chart templates, cached-value
refresh, chart cloning across slides / packages, `replace_data` variants that
preserve formulas and number formats, full `ShadowFormat`, alpha and `to_rgb`
across every color type, blip (picture) fills, line arrow ends, slide
transitions (23 variants including MORPH), shape animations, timing
introspection, comments read/write, OMML equation read/write, SmartArt
round-trip, OLE embedding for arbitrary progIds (ZIP/PDF/HTML/...), auto-play,
headers/footers/date/slide-numbers, Content placeholder chart/picture/table
insertion, font embedding (TTF + OTF), custom and extended document
properties, slide-level tags, XXE hardening + zip-bomb guard, paper-size
presets (letter/A4/16x9/4x3/tuple), `.ppsx` read + write, Flat-OPC XML save,
reproducible zip timestamps, password protection, and a long tail of smaller
OOXML capabilities that were previously out of reach.

This document is the full, single-page catalogue of what the library can do
today. Each section covers one feature area, opens with a short overview,
shows a copy-pasteable snippet against a fresh `Presentation()`, and then
lists the public methods, properties, and classes that make up that surface.
Items marked `[Added in 1.0.2.dev0]` are additions from this fork — every
other item is inherited from the upstream base.

**Table of contents**

- [Opening and saving presentations](#opening-and-saving-presentations)
- [Slides](#slides)
- [Slide masters and layouts](#slide-masters-and-layouts)
- [Sections](#sections)
- [Shapes](#shapes)
- [Autoshapes and connectors](#autoshapes-and-connectors)
- [Pictures](#pictures)
- [Media (audio and video)](#media-audio-and-video)
- [Tables](#tables)
- [Text frames, paragraphs, and runs](#text-frames-paragraphs-and-runs)
- [Charts](#charts)
- [Fills, colors, and effects](#fills-colors-and-effects)
- [Animations and transitions](#animations-and-transitions)
- [Comments](#comments)
- [Math equations](#math-equations)
- [SmartArt](#smartart)
- [OLE embedding](#ole-embedding)
- [Hyperlinks and click actions](#hyperlinks-and-click-actions)
- [Headers, footers, slide numbers, and date fields](#headers-footers-slide-numbers-and-date-fields)
- [Placeholders](#placeholders)
- [Font embedding](#font-embedding)
- [Document properties](#document-properties)
- [Slide-level tags](#slide-level-tags)
- [Security](#security)
- [Rendering](#rendering)
- [API concepts](#api-concepts)

---

## Opening and saving presentations

The top-level `pptx.Presentation()` factory opens a `.pptx`, `.pptm`, `.potx`,
or `.ppsx` package; called with no argument it creates a fresh presentation
from a built-in template. A `pptx_format` keyword selects one of the four
built-in paper-size presets (`"4x3"` / `"standard"`, `"16x9"` / `"widescreen"`,
`"letter"`, `"a4"`) or takes a `(cx, cy)` EMU tuple for a fully custom slide
size. `Presentation.save()` writes a normal `.pptx`; `save_ppsx()` writes a
slideshow package; `save_flat_xml()` emits ECMA-376 Part 4 Flat-OPC
single-file XML. Every save accepts a `zip_date_time` keyword for
reproducible (byte-identical) output and a `password` keyword for ECMA-376
Agile-Encryption-protected output.

```python
from pptx import Presentation
from pptx.util import Emu

# fresh default (4:3) presentation
prs = Presentation()
prs.save("out.pptx")

# widescreen
prs = Presentation(pptx_format="16x9")
prs.save("wide.pptx")

# US Letter paper
prs = Presentation(pptx_format="letter")

# A4 landscape
prs = Presentation(pptx_format="a4")

# fully custom slide size (10" x 5.625" in EMU)
prs = Presentation(pptx_format=(Emu(9144000), Emu(5143500)))

# reproducible save (byte-identical for the same content)
prs.save("rep.pptx", zip_date_time=(2026, 1, 1, 0, 0, 0))

# save as PowerPoint Show (.ppsx) — opens directly into slideshow playback
prs.save_ppsx("show.ppsx")

# save as Flat OPC (single-file .xml)
prs.save_flat_xml("deck.xml")

# open an encrypted package
# prs = Presentation("secret.pptx", password="hunter2")
# prs.save("out.pptx", password="hunter2")

# append one deck to another
# other = Presentation("chapter2.pptx")
# appended_slides = prs.merge(other)

# use an existing .pptx as a blank template (keep masters, drop slides)
# blank = Presentation("branded-template.pptx").strip_slides()
```

- `pptx.Presentation(pptx=None, pptx_format=None, password=None)` — Factory returning a `pptx.presentation.Presentation`. `pptx_format` accepts paper-size preset strings or `(cx, cy)` EMU tuples `[Added in 1.0.2.dev0]`; `password` decrypts on open `[Added in 1.0.2.dev0]`.
- Paper-size presets (case-insensitive): `"4x3"` / `"standard"`, `"16x9"` / `"widescreen"` `[Added in 1.0.2.dev0]`, `"letter"` `[Added in 1.0.2.dev0]`, `"a4"` `[Added in 1.0.2.dev0]`.
- `Presentation.save(file, zip_date_time=None, password=None)` — Write a regular `.pptx`. `zip_date_time` and `password` are `[Added in 1.0.2.dev0]`.
- `Presentation.save_ppsx(file, zip_date_time=None, password=None)` — Write as a PowerPoint Show (`.ppsx`). `[Added in 1.0.2.dev0]`
- `Presentation.save_flat_xml(file)` — Write as ECMA-376 Part 4 Flat-OPC single-file XML. `[Added in 1.0.2.dev0]`
- `Presentation.merge(other_presentation)` — Append every slide of another presentation with full-fidelity deep-copy (images, charts with distinct embedded workbooks, OLE, external hyperlinks). Returns the list of newly-appended `Slide` objects. `[Added in 1.0.2.dev0]`
- `Presentation.strip_slides()` — Remove every slide (and any sections) from this presentation while preserving slide masters, slide layouts, theme, embedded fonts, and other template-level resources. Returns `self` so it chains after the `Presentation()` factory call (`blank = Presentation("template.pptx").strip_slides()`). `[Added in 1.0.2.dev0]`
- `Presentation.slides` / `Presentation.slide_masters` / `Presentation.slide_layouts` / `Presentation.notes_master` — Core collections.
- `Presentation.slide_width` / `Presentation.slide_height` — Slide dimensions (read/write `Length`).
- `Presentation.sections` — `Sections` collection (presentation sections). `[Added in 1.0.2.dev0]`
- `Presentation.core_properties` / `Presentation.custom_properties` / `Presentation.extended_properties` — Document properties (see [Document properties](#document-properties)). `custom_properties` and `extended_properties` are `[Added in 1.0.2.dev0]`.
- `Presentation.embed_font(font_file, typeface, style="regular")` / `Presentation.embedded_fonts` — Embed a TTF/OTF font. `[Added in 1.0.2.dev0]`
- `Presentation.set_auto_advance(seconds, advance_on_click=False)` — Apply the same auto-advance delay to every slide. `[Added in 1.0.2.dev0]`
- `pptx.exc.EncryptedPackageError` — Raised on missing/wrong password or missing `msoffcrypto-tool`. `[Added in 1.0.2.dev0]`
- `pptx.exc.PackageTooLargeError` — Raised when a package exceeds the configurable uncompressed-size limit. `[Added in 1.0.2.dev0]`
- `pptx.exc.UnsupportedImageTypeError` — Raised for unsupported image formats like SVG. `[Added in 1.0.2.dev0]`

---

## Slides

A `Presentation.slides` collection supports add, duplicate, delete, reorder
(move), and external-slide-copy. Each `Slide` exposes `shapes`,
`placeholders`, `slide_layout`, `slide_id`, `notes_slide`, `has_notes_slide`,
`background`, `follow_master_background`, `name`, and the fork-era
`is_hidden`, `comments`, `has_comments`, `tags`, `has_tags`, `transition`,
`animation_sequence`, `iter_shape_animations`, `has_animations`,
`timing_xml`, and `find_shapes_by_xpath`.

```python
from pptx import Presentation

prs = Presentation()
layout = prs.slide_layouts[0]

# add a slide, give it a title
slide = prs.slides.add_slide(layout)
slide.shapes.title.text = "Hello"

# duplicate a slide (fork feature)
dupe = prs.slides.duplicate(slide)
dupe.shapes.title.text = "Hello (copy)"

# reorder
prs.slides.move_slide(dupe, 0)

# hide it
dupe.is_hidden = True

# delete
prs.slides.delete(dupe)

# look up by slide id
original = prs.slides.get(slide.slide_id)

# find a shape by XPath
title_shapes = slide.find_shapes_by_xpath(
    ".//p:sp[p:nvSpPr/p:cNvPr/@name='Title 1']"
)

prs.save("out.pptx")
```

- `Presentation.slides` — `Slides` collection (sequence).
- `Slides.add_slide(slide_layout)` — Append a new slide.
- `Slides.duplicate(slide, index=None)` — Deep-clone an existing slide within the presentation. `[Added in 1.0.2.dev0]`
- `Slides.delete(slide)` — Remove a slide. `[Added in 1.0.2.dev0]`
- `Slides.move_slide(slide, new_idx)` — Reorder. `[Added in 1.0.2.dev0]`
- `Slides.add_slide_from_external(source_slide, slide_layout)` — Full-fidelity copy from another presentation (images / charts / OLE rewritten into this package). `[Added in 1.0.2.dev0]`
- `Slides.get(slide_id, default=None)` — Look up by slide ID.
- `Slides.index(slide)` — Positional lookup.
- `Slide.slide_id` / `Slide.slide_layout` / `Slide.shapes` / `Slide.placeholders` / `Slide.name` / `Slide.element` / `Slide.part`.
- `Slide.is_hidden` (read/write `bool`) — `p:sld/@show="0"` for hidden slides. `[Added in 1.0.2.dev0]`
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

`Presentation.slide_masters` and `Presentation.slide_layouts` are sequences
of `SlideMaster` / `SlideLayout` objects. The fork widens both with
shape-creation APIs (they subclass `_BaseGroupShapes` so `add_shape`,
`add_picture`, `add_textbox`, `add_connector`, `add_group_shape`, and
`build_freeform` are available on masters and layouts), positional `name`
fallback for PowerPoint-authored masters with an empty `@name`, a
`SlideMaster.name` setter, and `SlideLayouts.remove`, `index`, `get_by_name`.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
master = prs.slide_masters[0]

# masters now carry a positional name when @name is blank
print(master.name)  # "Master 1" or similar

# rename it
master.name = "Corporate Master"

# bake a company logo onto every slide that inherits from this master
master.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(3), Inches(0.3))

# add a shape to a layout
layout = prs.slide_layouts[1]
layout.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(0), Inches(6.9), Inches(10), Inches(0.1),
)

# enumerate layouts by name
blank = prs.slide_layouts.get_by_name("Blank")

prs.save("out.pptx")
```

- `Presentation.slide_masters` — `SlideMasters` sequence.
- `Presentation.slide_layouts` — `SlideLayouts` sequence on the primary master.
- `SlideMaster.shapes` / `SlideMaster.placeholders` — Shape tree (write-enabled `[Added in 1.0.2.dev0]`).
- `SlideMaster.name` (read/write) — `[Added in 1.0.2.dev0]`. Returns `"Master N"` positional fallback when `p:cSld/@name` is empty.
- `SlideMaster.slide_layouts` — Child layouts.
- `SlideMaster.background` / `SlideMaster.theme_colors` — Master-level theming.
- `SlideMaster.header_footer` — Default header/footer overrides.
- `SlideLayouts.get_by_name(name, default=None)` — Name lookup.
- `SlideLayouts.index(slide_layout)` — Positional lookup.
- `SlideLayouts.remove(slide_layout)` — Delete an unused layout.
- `SlideLayout.shapes` / `SlideLayout.placeholders` — Shape tree (write-enabled `[Added in 1.0.2.dev0]`).
- `SlideLayout.name` / `SlideLayout.used_by_slides` / `SlideLayout.slide_master` / `SlideLayout.header_footer`.
- `SlideLayout.iter_cloneable_placeholders()` — Placeholders new slides inherit.
- `MasterShapes` / `LayoutShapes` — Shape collections; inherit `add_shape`, `add_picture`, `add_textbox`, `add_connector`, `add_group_shape`, `build_freeform`. `[Added in 1.0.2.dev0]`

---

## Sections

`Presentation.sections` is an ordered collection of PowerPoint-2010
sections (`p14:sectionLst`). Each `Section` carries a read/write `name`, a
stable GUID `id`, an `index` property, slide membership with `add_slide` /
`remove_slide` / `move_slide`, and reorder helpers `move_before` /
`move_after`.

```python
from pptx import Presentation

prs = Presentation()
layout = prs.slide_layouts[0]
s1 = prs.slides.add_slide(layout)
s2 = prs.slides.add_slide(layout)
s3 = prs.slides.add_slide(layout)

intro = prs.sections.add_section(name="Intro")
details = prs.sections.add_section(name="Details")

intro.add_slide(s1)
details.add_slide(s2)
details.add_slide(s3)

# iterate
for section in prs.sections:
    print(section.index, section.name, len(section.slides))

# reorder
details.move_before(intro)

# find the section that owns a slide
owner = prs.sections.find_containing(s2)

# look up by name
found = prs.sections.get_by_name("Intro")

prs.save("out.pptx")
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `Presentation.sections` — `Sections` collection.
- `Sections.add_section(name=..., section_id=..., insert_before=...)` — Add a new section.
- `Sections.remove(section)` — Detach a section (slides are not deleted).
- `Sections.get_by_id(id)` / `Sections.get_by_name(name)` — Lookup.
- `Sections.index(section)` — Positional lookup.
- `Sections.find_containing(slide)` — Return the `Section` that owns a given `Slide`, or `None`.
- `Section.name` (read/write) / `Section.id` (GUID) / `Section.index`.
- `Section.slides` — Tuple of slides owned by this section.
- `Section.add_slide(slide)` — Assign a slide (raises `ValueError` if already owned elsewhere).
- `Section.move_slide(slide)` — Reassign a slide from one section to another in one step.
- `Section.remove_slide(slide)` — Remove slide from this section (does not delete from the deck).
- `Section.move_before(other)` / `Section.move_after(other)` — Reorder without touching slide order.
- `Section.element` — Underlying `p14:section` element.

---

## Shapes

Every shape derives from `BaseShape`. The fork extends `BaseShape` with
`delete`, `duplicate`, `replace_with`, `flip_horizontal` / `flip_vertical`
(as properties) plus `flip_horizontally()` / `flip_vertically()` methods,
`is_hidden`, `alt_text` / `title` accessibility metadata,
z-order helpers, `effective_left` / `effective_top` /
`effective_width` / `effective_height` (applying the enclosing group
transforms), `has_math_equation` / `math_equation_xml`, and the
shape-animation surface (`animation` / `set_animation`). `SlideShapes`
gains `get_by_name`, `find_all_by_name`, `add_picture_link`, `clone_chart`,
and integrates with the `Slide.find_shapes_by_xpath` lookup.

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

# add, duplicate, delete, replace
rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              Inches(1), Inches(1), Inches(2), Inches(1))
clone = rect.duplicate()
clone.left = Inches(4)

# flip, hide, z-order
clone.flip_horizontally()
clone.is_hidden = True
clone.bring_to_front()
print(clone.zorder_index)

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

---

## Autoshapes and connectors

Autoshapes are added via `SlideShapes.add_shape(MSO_SHAPE.*)`. The fork
adds `MSO_SHAPE.LINE` (so `.auto_shape_type` on a straight-line `prst="line"`
resolves cleanly), read/write `Connector.adjustments` for elbow and curved
connectors, and line-end arrow configuration via `LineFormat.begin_arrow` /
`.end_arrow`.

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

# swap the bytes, keep the shape
pic.replace_image(BytesIO(png))

# linked (external URL) picture
link = slide.shapes.add_picture_link(
    "https://example.com/logo.png",
    Inches(3), Inches(1), Inches(1), Inches(1),
)

prs.save("out.pptx")
```

- `SlideShapes.add_picture(image_file, left, top, width=None, height=None)` — Embed a picture. Accepts path, file-like, or `BytesIO`. MPO images `[Added in 1.0.2.dev0]`; EMF content-type fixed `[Added in 1.0.2.dev0]`.
- `SlideShapes.add_picture_link(url, left, top, width=None, height=None)` — Insert a linked (external URL) picture; no bytes stored. `[Added in 1.0.2.dev0]`
- `Picture.replace_image(image_file)` — Swap the embedded image while keeping position, size, rotation, crop, outline, and masking. `[Added in 1.0.2.dev0]`
- `Picture.image` — `Image` object with `.blob`, `.content_type`, `.ext`, `.filename`, `.size`, `.dpi`. Returns `None` when the `p:pic` has no `p:blipFill` (malformed file). `[Added in 1.0.2.dev0]`
- `Picture.crop_left` / `.crop_top` / `.crop_right` / `.crop_bottom` — Crop fractions (read/write).
- `Picture.auto_shape_type` — Masking shape for the picture (read/write).
- `Picture.line` — `LineFormat` for the picture outline.
- `Picture.shadow` — `ShadowFormat` (see [Fills, colors, and effects](#fills-colors-and-effects)).
- `Picture.delete()` — Remove the picture and drop the image relationship. `[Added in 1.0.2.dev0]`
- `pptx.exc.UnsupportedImageTypeError` — Raised for SVG and other unsupported formats. `[Added in 1.0.2.dev0]`

---

## Media (audio and video)

`SlideShapes.add_movie()` adds an audio or video clip. The fork adds an
`autoplay=True` kwarg on `add_movie` for the common "start with the slide"
case, `Movie.replace_media()` for swapping the underlying bytes while keeping
the shape, `Movie.blob` / `Movie.ext` / `Movie.content_type` for *reading*
the embedded media bytes and MIME type, `Movie.start_time` /
`Movie.start_condition` for fine-grained timing control, `Movie.delete()`
that correctly removes the timing-tree entry and the three media-related
slide-part rels, MIME-type registrations for common audio formats (so
`add_movie` works on a slide whose layout already holds an MP3), and a
movie-shape name that strips the file extension to match PowerPoint's
convention.

```python
import io
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# add an audio clip (fake bytes for example)
mp3 = io.BytesIO(b"ID3\x03\x00\x00\x00\x00\x00\x00")
movie = slide.shapes.add_movie(
    mp3, Inches(1), Inches(1), Inches(0.5), Inches(0.5),
    mime_type="audio/mpeg",
    autoplay=True,  # start with the slide; default is click-to-play
)

# override the default-shorthand for finer-grained control
movie.start_condition = "afterPrevious"
movie.start_time = 2.5

# read embedded media bytes back out
ext = movie.ext                 # e.g. 'mp3'
ctype = movie.content_type      # e.g. 'audio/mpeg'
media_bytes = movie.blob        # raw bytes of the embedded clip

# swap the media in place
movie.replace_media(io.BytesIO(b"ID3\x03\x00\x00\x00\x00\x00\x00"),
                    mime_type="audio/mpeg")

# delete cleanly (drops all 3 slide rels + the p:video timing entry)
# movie.delete()

prs.save("out.pptx")
```

- `SlideShapes.add_movie(movie_file, left, top, width, height, poster_frame_image=None, mime_type=None, autoplay=False)` — Add a movie / audio. Supports MP3, WAV, AIFF, MIDI, MP4, MPEG, OGG, WMA, AVI, MOV, WMV, SWF, and generic `audio/*` / `video/*`. `autoplay=True` sets `start_condition="withPrevious"` on the new shape (issue #427). `[Added in 1.0.2.dev0]`.
- `Movie.replace_media(new_path_or_file, mime_type=None)` — Swap the bytes while preserving position, poster, hyperlink, and timing entries. `[Added in 1.0.2.dev0]`
- `Movie.blob` / `Movie.ext` / `Movie.content_type` — Read embedded media bytes, file extension, and MIME type. Return `None` when the shape has no associated media part. `[Added in 1.0.2.dev0]`
- `Movie.start_condition` — `"onClick"` / `"withPrevious"` / `"afterPrevious"`. `[Added in 1.0.2.dev0]`
- `Movie.start_time` — Float seconds, or `None` for `"indefinite"`. `[Added in 1.0.2.dev0]`
- `Movie.delete()` — Remove the movie shape, the three media rels, and the `p:timing/.../p:video` entry. `[Added in 1.0.2.dev0]`
- `Movie.media_type` / `Movie.media_format` / `Movie.poster_frame` — Read-only introspection.
- `Movie.shape_type` — `MSO_SHAPE_TYPE.MEDIA`.
- `ActionSetting.set_sound(sound_file, filename=None)` / `ActionSetting.remove_sound()` / `ActionSetting.sound` — Click-action sound attachment. `[Added in 1.0.2.dev0]`

---

## Tables

Tables support cell merge / split, borders, margins, text direction,
banding flags (first-row / last-row / first-col / last-col / horz-banding /
vert-banding), row / column add / delete, per-cell `row_idx` / `col_idx`,
and the fork addition of `Table.style_id` for applying built-in PowerPoint
table styles by their GUID.

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

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
- `_Cell.merge(other_cell)` / `_Cell.split()` / `_Cell.is_merge_origin` / `_Cell.is_spanned` / `_Cell.span_height` / `_Cell.span_width` — Merge handling.
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
`a:clrScheme`), `Font.strikethrough`, `Font.use_theme_hyperlink_color`,
`Font.name_ea` / `name_cs` (East-Asian and complex-script slots), and a
complete bullet-format API via `_Paragraph.bullet`.

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_TEXT_STRIKE_TYPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2))
tf = tb.text_frame

# rotation of the entire text body
tf.rotation = 90.0

# build a paragraph
p = tf.paragraphs[0]
r = p.add_run()
r.text = "Hello {NAME}"
r.font.name = "Calibri"
r.font.size = Pt(18)
r.font.bold = True
r.font.strikethrough = MSO_TEXT_STRIKE_TYPE.SINGLE_LINE
r.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

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
- `_Hyperlink.address` — Run-level hyperlink URL (setter creates/clears the `a:hlinkClick`).

---

## Charts

Charts are added via `SlideShapes.add_chart(chart_type, x, y, cx, cy,
chart_data)`. The fork extends chart support with 3D types, combo charts
(a second `add_plot` onto an existing chart), secondary value axes,
`Chart.apply_template()` to apply a `.crtx` package, `Chart.clone_to()` for
cross-slide / cross-deck duplication, `Chart.replace_data()` preserving
author-set `formatCode`, `Chart.replace_data_preserve_formulas()`,
`Chart.update_cached_values()`, `ChartTitle.position`, `Chart.plot_area.format`
(plot-area fill / line / shadow), `Axis.major_unit` /
`.minor_unit` for date axes, `DataLabel.format` and `DataLabels.format`
`ChartFormat` wrappers (border, fill, shadow on data labels), error bars,
theme accent-color cycling on cloned series, the `UNSUPPORTED_CHARTEX`
sentinel that preserves unknown `cx:` charts on round-trip, and
extended-chart-style IDs (1–48 plain plus 49–255 via `c14:style`).

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData, XyChartData
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
- `Chart.font` — `Font` proxy for chart-space default text properties.
- `ChartTitle.text_frame` / `ChartTitle.format` / `ChartTitle.has_text_frame` — Title body.
- `ChartTitle.position` — `(x, y)` tuple in factor-space `[0, 1]`, or `None` for auto layout. `[Added in 1.0.2.dev0]`
- `Axis.format` / `Axis.tick_labels` / `Axis.major_gridlines` / `Axis.minor_gridlines` / `Axis.axis_title`.
- `Axis.position` — Read/write `XL_AXIS_POSITION`. `[Added in 1.0.2.dev0]`
- `Axis.visible` — Read/write `bool` toggling axis visibility via the `c:delete` element. Writes `val="1"` / `val="0"` explicitly so PowerPoint honors the setting (see issue #852). `[Fixed in 1.0.2.dev0]`
- `TickLabels.rotation` — Read/write clockwise rotation (degrees) of axis tick labels, mapping to `c:txPr/a:bodyPr/@rot`. Accepts `int` or `float`; default `0.0`. `[Added in 1.0.2.dev0]`
- `ValueAxis.crosses` / `.crosses_at` / `.major_unit` / `.minor_unit`.
- `DateAxis.major_unit` / `.minor_unit` — Time-axis spacing. `[Added in 1.0.2.dev0]`
- `Series.values` / `Series.categories` / `Series.name` / `Series.format` / `Series.marker` / `Series.points`.
- `Series.has_error_bars` / `Series.error_bars` / `Series.set_error_bars(type, amount, include, direction)` — Error bars (fixed value / percentage / std deviation / std error). `[Added in 1.0.2.dev0]`
- `Point.format` / `Point.marker` / `Point.data_label` / `Point.invert_if_negative` — Per-point formatting.
- `DataLabel.text_frame` / `DataLabel.font` / `DataLabel.position` / `DataLabel.show_*`.
- `DataLabel.format` — `ChartFormat` wrapping this single `c:dLbl` with `.fill` / `.line` / `.shadow`. `[Added in 1.0.2.dev0]`
- `DataLabels.text_frame` / `DataLabels.font` / `DataLabels.show_*` / `DataLabels.number_format`.
- `DataLabels.format` — `ChartFormat` wrapping `c:dLbls` (collection-scope fill / line / shadow). `[Added in 1.0.2.dev0]`
- `ChartFormat.fill` / `ChartFormat.line` / `ChartFormat.shadow` — Format block for a chart element. `shadow` is `[Added in 1.0.2.dev0]`.
- `pptx.enum.chart.XL_CHART_TYPE` — Full chart-type enum including 2D, 3D, XY, bubble, radar, doughnut, area, and a chartex passthrough sentinel (`UNSUPPORTED_CHARTEX`) for unknown `cx:` types `[Added in 1.0.2.dev0]`.
- `pptx.enum.chart.XL_ERROR_BAR_TYPE` / `XL_ERROR_BAR_INCLUDE` / `XL_ERROR_BAR_DIRECTION` — Error-bar configuration. `[Added in 1.0.2.dev0]`
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
`remove_sound()` / `sound`, jump-to-named-slide targets for runs, and
`Font.use_theme_hyperlink_color` to override the hyperlink color.

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

# "Title and Content" layout has a Content placeholder at idx 1
slide = prs.slides.add_slide(prs.slide_layouts[1])
ph = slide.placeholders[1]

# insert a chart into a generic content placeholder (fork lift)
data = CategoryChartData()
data.categories = ["A", "B", "C"]
data.add_series("S1", (1, 2, 3))
ph.insert_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, data)

# "Picture with Caption" has a PicturePlaceholder at idx 1
# slide2 = prs.slides.add_slide(prs.slide_layouts[8])
# ph2 = slide2.placeholders[1]
# ph2.insert_picture("logo.png", crop=False)   # fit, preserving aspect ratio

prs.save("out.pptx")
```

- `Slide.placeholders` — `SlidePlaceholders` keyed by `idx`.
- `SlidePlaceholder.insert_chart(chart_type, chart_data)` — Replace the placeholder with a chart. `[Added in 1.0.2.dev0]` (previously `ChartPlaceholder`-only).
- `SlidePlaceholder.insert_picture(image_file, crop=True)` — Replace with a picture. `crop=False` fits without cropping. `[Added in 1.0.2.dev0]` for both the lift and the `crop` keyword.
- `SlidePlaceholder.insert_table(rows, cols)` — Replace with a table. `[Added in 1.0.2.dev0]`
- `PicturePlaceholder` / `ChartPlaceholder` / `TablePlaceholder` — Specialized subclasses (inherit the same insertion API).
- `_PlaceholderFormat.idx` / `_PlaceholderFormat.type` — Placeholder descriptor.
- `LayoutPlaceholder` / `MasterPlaceholder` — Layout / master variants with dimension inheritance.
- `Slide.placeholders` / `SlideLayout.placeholders` / `SlideMaster.placeholders` — Placeholder collections.

---

## Font embedding

`Presentation.embed_font()` writes an arbitrary TrueType (`.ttf`) or
OpenType (`.otf`) font file into the package as a `ppt/fonts/fontN.fntdata`
part and registers it under a `p:embeddedFontLst/p:embeddedFont` entry for
the given typeface name. PowerPoint treats TTF and OTF payloads
identically; both are carried as `application/x-fontdata` content.

```python
from io import BytesIO

from pptx import Presentation

prs = Presentation()

# embed_font accepts a path or a file-like
ttf_bytes = b"\x00\x01\x00\x00"  # placeholder TTF magic; use real font bytes in practice
prs.embed_font(BytesIO(ttf_bytes), typeface="Pacifico", style="regular")
prs.embed_font(BytesIO(ttf_bytes), typeface="Pacifico", style="bold")

print(prs.embedded_fonts)   # ("Pacifico",)

prs.save("out.pptx")
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `Presentation.embed_font(font_file, typeface, style="regular")` — Embed a `.ttf` / `.otf` file; `style` is one of `"regular"`, `"bold"`, `"italic"`, `"boldItalic"`. Repeated calls for the same `typeface`/`style` replace; different styles accumulate on the same entry.
- `Presentation.embedded_fonts` — Tuple of embedded typeface names in document order.

---

## Document properties

Three categories of document properties are exposed:

- **Core** (Dublin-Core) via `Presentation.core_properties`.
- **Custom** (typed user-defined) via `Presentation.custom_properties` —
  the store `{ DOCPROPERTY }` field codes resolve against.
- **Extended** (application) via `Presentation.extended_properties` —
  `docProps/app.xml` fields such as `company`, `manager`, `slide_count`.

```python
import datetime as dt

from pptx import Presentation

prs = Presentation()

# core (Dublin Core)
cp = prs.core_properties
cp.author = "Ben"
cp.title = "Quarterly Report"
cp.subject = "Q1 results"

# custom (dict-like, typed)
prs.custom_properties["ReviewerCount"] = 3
prs.custom_properties["IsDraft"] = True
prs.custom_properties["ReleaseDate"] = dt.datetime(2026, 1, 1)
prs.custom_properties.update({"Department": "Finance"})

# extended (app.xml)
ep = prs.extended_properties
ep.company = "Example Inc"
ep.manager = "Alex"
print(ep.application, ep.app_version, ep.slide_count)

prs.save("out.pptx")
```

- `Presentation.core_properties` — `CoreProperties` (`author`, `title`, `subject`, `keywords`, `category`, `comments`, `content_status`, `identifier`, `language`, `version`, `created`, `last_modified_by`, `last_printed`, `modified`, `revision`).
- `Presentation.custom_properties` — `CustomProperties` dict-like store for typed user-defined properties. Supports `str` / `int` (32-bit signed) / `float` / `bool` / `datetime.datetime`. Full mapping interface: `__getitem__`, `__setitem__`, `__delitem__`, `__contains__`, `__iter__`, `__len__`, `keys()`, `items()`, `values()`, `update()`, `clear()`, `pop()`, `setdefault()`. `[Added in 1.0.2.dev0]`
- `Presentation.extended_properties` — `ExtendedPropertiesPart` for `docProps/app.xml`. Exposes typed accessors (`company`, `manager`, `hyperlink_base`, `application`, `app_version`, `presentation_format`, `template`, `slide_count`, `notes_count`, `title_of_parts`, `heading_pairs`, ...). `slide_count` is automatically refreshed at save-time from the live `sldIdLst`. `[Added in 1.0.2.dev0]`

---

## Slide-level tags

Each `Slide` carries an optional tag store (`ppt/tags/tagN.xml`) — an
OPC part that maps string keys to string values. The fork exposes this
as a dict-like `Slide.tags` collection; `Slide.has_tags` is `True` when a
tag part is attached.

```python
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# write
slide.tags["CUSTOMER_ID"] = "CUST-42"
slide.tags["CATEGORY"] = "finance"

# read
for key, value in slide.tags.items():
    print(key, "=", value)

print("CUSTOMER_ID" in slide.tags, slide.has_tags)

prs.save("out.pptx")
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `Slide.has_tags` — `True` when a `TagsPart` is attached.
- `Slide.tags` — `SlideTags` dict-like collection. Full mapping interface: `__getitem__`, `__setitem__`, `__delitem__`, `__contains__`, `__iter__`, `__len__`, `get()`, `keys()`, `values()`, `items()`.

---

## Security

The fork hardens the XML parser and the zip reader against malicious
input. The lxml `XMLParser` used to parse every `.pptx` part (and the one
used to read chart-embedded workbooks) disables entity resolution, DTD
loading, and network access — defeating both "billion laughs"
entity-expansion attacks and XML external entity (XXE) attacks. The zip
reader inspects the central directory's uncompressed-size fields before
loading members into memory and raises `PackageTooLargeError` when the
declared total exceeds a configurable limit (default 2 GiB).

```python
import os

# tune the zip-bomb guard (default 2 GiB)
os.environ["PPTX_MAX_UNCOMPRESSED_SIZE"] = str(512 * 1024 * 1024)  # 512 MiB

# or programmatically
import pptx.opc.serialized as serialized
serialized.MAX_UNCOMPRESSED_PACKAGE_SIZE = 512 * 1024 * 1024

from pptx import Presentation
from pptx.exc import PackageTooLargeError

try:
    Presentation("suspect.pptx")
except PackageTooLargeError as exc:
    print("rejected:", exc)
```

All members in this section are `[Added in 1.0.2.dev0]`.

- `pptx.exc.PackageTooLargeError` — Raised when the declared uncompressed size of a package exceeds the limit.
- `pptx.exc.EncryptedPackageError` — Raised when opening an encrypted package without a valid password or without `msoffcrypto-tool`.
- `pptx.exc.UnsupportedImageTypeError` — Raised on SVG / unrecognized image input.
- `PPTX_MAX_UNCOMPRESSED_SIZE` environment variable — Override the zip-bomb guard.
- `pptx.opc.serialized.MAX_UNCOMPRESSED_PACKAGE_SIZE` module attribute — Same, at runtime.
- XML parser hardening — Applied transparently to every `XMLParser` used by the library.

---

## Rendering

python-pptx does **not** render slides. It emits (and round-trips) the
OOXML a renderer consumes; pixel / video / PDF output is out of scope
and requires a layout engine the library does not ship. The user guide's
"Rendering to video, PDF, or image formats" section
(`docs/user/use-cases.rst`) covers the recommended integration paths:

- `libreoffice --headless --convert-to pdf` (and `--convert-to png` for
  first-slide-only; `pdftoppm` or ImageMagick `convert` for per-slide image
  export).
- PowerPoint COM automation (Windows only).
- Aspose.Slides (commercial).
- `python-pptx-interface` (third-party library that drives PowerPoint /
  LibreOffice).

The same guide covers `.pptx` → `.mp4` (issue #1049) and `.pptx` → `.pdf`
(issue #584); neither is implemented in python-pptx itself.

---

## API concepts

`python-pptx` is organised in three layers:

- **Presentation API** (`src/pptx/api.py`, `src/pptx/presentation.py`,
  `src/pptx/slide.py`, `src/pptx/shapes/*.py`, `src/pptx/chart/*.py`,
  `src/pptx/text/*.py`, `src/pptx/dml/*.py`, …) — Proxy objects wrapping
  oxml elements. This is where virtually all user code lives.
- **Parts layer** (`src/pptx/parts/*.py`) — `XmlPart` subclasses that own
  the XML trees for each part (presentation, slide, slide-master,
  slide-layout, notes-master, notes-slide, theme, chart, image, media,
  font, comments, tags, core / extended / custom properties, …) and
  manage the relationships between them.
- **oxml layer** (`src/pptx/oxml/*.py`) — `CT_*` classes extending
  `lxml.etree.ElementBase` and mapping directly onto schema element names.

`lxml` handles XML parsing, serialisation, and XPath work beneath the
library. `pptx.util` carries `Length` subclasses (`Inches`, `Cm`, `Mm`,
`Pt`, `Emu`, `Centipoints`) and a bundle of helpers; `pptx.dml.color`
carries `RGBColor`.

```python
from pptx import Presentation
from pptx.util import Inches, Cm, Pt, Emu
from pptx.dml.color import RGBColor

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# lengths — typed ints, freely interchangeable
w = Inches(2)
print(int(w), Cm(5.08))          # same length, different constructor
print(Pt(12), RGBColor(0x2E, 0x74, 0xB5))

slide.shapes.add_textbox(Inches(1), Cm(2.54), Emu(914400), Pt(72))
prs.save("out.pptx")
```

- `pptx.util.Length` / `Inches` / `Cm` / `Mm` / `Pt` / `Emu` / `Centipoints` — Length constructors and arithmetic. All inherit from `int` so they compose freely.
- `pptx.dml.color.RGBColor` — `(r, g, b)` triple with `from_string()` and hex output.
- `pptx.shared.ElementProxy` / `ParentedElementProxy` / `PartElementProxy` / `Subshape` — Proxy base classes.
- `pptx.opc.constants.CONTENT_TYPE` / `RELATIONSHIP_TYPE` — Content-type and rel-type constants used by the parts layer.
- `pptx.oxml.ns.qn(tag)` — Clark-notation tag expansion; useful when dropping into the oxml layer.
- `pptx.enum.*` — Every public enumeration (`MSO_SHAPE`, `MSO_SHAPE_TYPE`, `MSO_CONNECTOR`, `MSO_AUTO_SIZE`, `MSO_COLOR_TYPE`, `MSO_FILL_TYPE`, `MSO_LINE_END_TYPE`/`WIDTH`/`LENGTH`, `MSO_PATTERN_TYPE`, `MSO_THEME_COLOR`, `MSO_TEXT_STRIKE_TYPE`, `MSO_TEXT_UNDERLINE_TYPE`, `MSO_VERTICAL_ANCHOR`, `MSO_LANGUAGE_ID`, `MSO_ANIMATION_TYPE`, `MSO_ANIMATION_TRIGGER`, `PP_ACTION`, `PP_ALIGN` / `PP_PARAGRAPH_ALIGNMENT`, `PP_AUTO_NUMBER_SCHEME`, `PP_MEDIA_TYPE`, `PP_PLACEHOLDER`, `PP_TRANSITION_TYPE`, `PP_TRANSITION_SPEED`, `PP_TRANSITION_SIDE_DIRECTION`, `XL_CHART_TYPE`, `XL_AXIS_POSITION`, `XL_ERROR_BAR_*`, `XL_LABEL_POSITION`, `XL_LEGEND_POSITION`, `XL_MARKER_STYLE`, `XL_TICK_MARK`, `XL_CATEGORY_TYPE`, `XL_DATA_LABEL_POSITION`).

---

*This file is generated and maintained by hand — see `HISTORY.rst` for the
full change log, `docs/user/*.rst` for narrative tutorials, and
`docs/api/*.rst` for per-class API reference pages.*
