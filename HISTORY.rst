.. :changelog:

Release History
---------------

Unreleased
++++++++++

- verify: #694 resolved by F7 foundation. The presentation-sections subsystem
  covers the issue's user stories (iterate ``prs.sections`` and each
  ``section.slides``, search by name via ``Sections.get_by_name``, sort
  section objects, round-trip author-supplied section GUIDs). Adds an
  end-to-end regression suite ``DescribeIssue694RegressionSections`` under
  ``tests/test_presentation.py`` that exercises those scenarios against a
  round-tripped real package.
- feat: #1004 expose ``Transition.speed`` using the ``PP_TRANSITION_SPEED``
  enum (``SLOW`` / ``MEDIUM`` / ``FAST`` → ``p:transition/@spd``) and add
  the per-variant ``Transition.wipe_direction`` accessor
  (``PP_TRANSITION_SIDE_DIRECTION``) for reading / writing the ``@dir``
  attribute on ``p:wipe``. Setting ``wipe_direction`` also switches the
  transition variant to wipe. Introduces ``CT_SideDirectionTransition``
  (typed round-trip of the schema's ``CT_SideDirectionTransition``),
  registered for ``p:wipe`` and ``p:push``. Remaining per-variant flags
  (``Cover.direction``, ``Split.orientation``, ``Fade.through_black``,
  ``Wheel.spokes``, etc.) can follow the same pattern incrementally.
- feat: #942 activate MORPH slide transition. ``Slide.transition.type =
  PP_TRANSITION_TYPE.MORPH`` now writes the ``p14:morph`` variant wrapped
  in an ``mc:AlternateContent`` container with a ``p:fade`` fallback (the
  form PowerPoint itself emits) so Office 2013+ viewers apply MORPH and
  older viewers render a graceful fade. Adds ``Transition.morph_option``
  for selecting the matching granularity (one of ``"byObject"`` /
  ``"byWord"`` / ``"byChar"``, default ``"byObject"``). Switching the
  transition away from MORPH (to any plain variant or to ``NONE``)
  automatically unwraps the ``mc:AlternateContent``. The ``Transition``
  proxy now also reads/writes ``duration`` / ``advance_on_click`` /
  ``advance_after_time`` transparently on a wrapped or direct
  ``p:transition`` element.
- verify: #400 animation-control umbrella resolved by Foundation F8
  (``feat/foundation-f8-animations-transitions``) plus the four leaf
  branches it aggregates — ``feat/issue-102-shape-animation``,
  ``feat/issue-1106-entrance-exit-animations``,
  ``feat/issue-264-shape-animation-control``, and
  ``feat/issue-861-animation-delay``. Adds a regression test
  (``tests/test_issue_400_animation_umbrella.py``) that exercises the
  specific flow the #400 reporter asked for: authoring an entrance
  animation on a shape, round-tripping it through save + reopen,
  introspecting the animated targets via ``Slide.timing_xml``, and
  mixing the animation tree with a ``Slide.transition`` without
  clobbering either subtree. The structured authoring API
  (``Slide.animations.add_entrance_effect(...)``) is delivered by the
  four leaves; F8 provides the typed element classes and the
  ``Slide.has_animations`` / ``Slide.timing_xml`` introspection surface
  the umbrella pins.
- feat: #877 cross-slide chart copy. Adds ``Chart.clone_to(shapes, x, y,
  cx, cy)`` and ``SlideShapes.clone_chart(source_chart, x, y, cx, cy)``
  for duplicating a chart onto another slide in the same presentation or
  into a different presentation entirely. The duplicate receives its own
  ``ChartPart`` (deep copy of the source ``c:chartSpace`` XML via
  F1 ``PartRelationshipCloner``) and its own ``EmbeddedXlsxPart`` (cloned
  via F5 ``clone_embedded_xlsx``) so each chart keeps a working "Edit
  Data" workbook. All non-xlsx relationships the source chart part
  carries (chart images, theme-override, ...) are re-established on the
  duplicate; cross-package clones materialise each referenced part in
  the destination package so the target file can be saved and opened
  standalone. Adds ``ChartPart.clone_from(source_chart_part, package)``
  as the supporting primitive.
- feat(chart): #239 ``Chart.replace_data_preserve_formulas(chart_data)``.
  A targeted-refresh counterpart to ``Chart.replace_data`` that walks the
  data cells of the embedded workbook and rewrites only the cells that
  don't carry an ``<f>`` formula element, so author-entered formulas
  survive a data refresh. Reuses the F5 ``WorkbookUpdater`` /
  ``_SingleCellCacheRefresher`` machinery so cached chart values stay in
  sync. Limited to data-only refresh — cannot add/remove series or
  categories — and skipped formula cells retain their prior
  ``c:numCache`` entry until ``Chart.update_cached_values`` is called.
- Foundation: cross-part embedded-workbook handler (F5)
- Foundation: presentation sections (F7). Adds read/write access to
  PowerPoint-2010 *sections* (``p14:sectionLst`` under
  ``p:extLst/p:ext[uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"]``) via a
  new ``Presentation.sections`` collection, with ``Sections.add_section``,
  ``Sections.remove``, ``Sections.get_by_id``, ``Sections.get_by_name``,
  and per-section ``Section.name``, ``Section.id`` (GUID), ``Section.slides``,
  ``Section.add_slide``, and ``Section.remove_slide``.
- Foundation: SmartArt scaffolding (F9 — MVP). Surfaces SmartArt graphic
  frames in ``slide.shapes`` via new ``GraphicFrame.has_smart_art``,
  ``GraphicFrame.smart_art``, and ``shape_type ==
  MSO_SHAPE_TYPE.IGX_GRAPHIC`` branches, plus a ``SmartArt`` proxy
  exposing the four interlinked parts (``data_xml`` / ``layout_xml`` /
  ``colors_xml`` / ``quick_style_xml``) as read-only raw bytes. Adds
  ``GRAPHIC_DATA_URI_SMART_ART`` spec constant, the ``dgm:`` namespace,
  and a ``CT_DgmRelIds`` oxml class for the four-rId child of
  ``a:graphicData``. Round-trip preservation of the diagramData /
  diagramLayout / diagramColors / diagramQuickStyle parts is automatic
  via the ``PartFactory`` fallthrough. Structured node-tree authoring
  and layout selection are out of scope for this foundation and remain
  open under #83; see ``docs/dev/analysis/f9-smartart.rst`` for the
  per-subsystem roadmap.
- feat: #257 Presentation sections — leaf API. Adds ordering and location
  conveniences on top of the F7 foundation: ``Section.index`` /
  ``Sections.index(section)`` expose a section's zero-based position;
  ``Section.move_before(other)`` / ``Section.move_after(other)`` reorder
  sections without touching the slide list; ``Sections.find_containing(slide)``
  returns the |Section| that owns a given slide (or |None|); and
  ``Section.add_slide()`` now raises a helpful ``ValueError`` naming the
  existing owner when a slide is already assigned to another section — use
  the new ``Section.move_slide(slide)`` to reassign in one step.
- #376 auto-play by seconds — add ``Presentation.set_auto_advance(seconds,
  advance_on_click=False)`` bulk helper that applies the same auto-advance
  delay to every slide in the deck. ``seconds`` accepts int or float (e.g.
  ``0.5``) and is converted to the milliseconds stored in
  ``p:transition/@advTm``; passing ``None`` clears the timer, restoring
  click-only advance. ``advance_on_click`` defaults to ``False`` for
  kiosk-style (timer-only) playback; set it to ``True`` to allow either
  a click or the timer to advance. Builds on the F8 per-slide
  ``Slide.transition.advance_after_time`` / ``advance_on_click`` API.
- fix: #954 ``SlideShapes.add_movie()`` produced a duplicate ``p:timing``
  element when the slide already carried a pre-existing ``p:timing``
  wrapped inside ``mc:AlternateContent``/``mc:Choice`` (the form
  PowerPoint emits when timing content references 2010+ extensions
  such as a morph trigger). The wrapped timing was invisible to the
  previous ``./p:timing`` xpath and a second, orphan ``p:timing`` was
  appended as a direct child of ``p:sld``. The fix finds an existing
  ``p:timing`` whether plain or wrapped (via F3's ``mc:AlternateContent``
  traversal and F8's typed ``CT_SlideTiming``) and merges the new
  ``p:video`` into its ``p:childTnLst``.
- Foundation: animations/transitions XML layer (F8 — MVP). Adds element
  classes for ``p:timing`` / ``p:tnLst`` / ``p:par`` / ``p:seq`` / ``p:cTn``
  and the ``p:transition`` subtree (including the ``p14:morph`` Office 2010
  extension) so slides round-trip these elements without dropping them.
  Exposes ``Slide.transition`` returning a ``Transition`` proxy with
  ``.type`` (``PP_TRANSITION_TYPE`` enum — fade/wipe/push/cover/morph/…),
  ``.duration`` (milliseconds via ``p14:dur``), ``.advance_on_click``, and
  ``.advance_after_time``. Adds ``Slide.has_animations`` and
  ``Slide.timing_xml`` for round-trip debugging. Structured entrance /
  exit / emphasis / motion-path / MORPH authoring APIs layer onto this
  foundation incrementally; see ``docs/dev/analysis/f8-animations-transitions.rst``.
- docs: #1049 PPT ---> MP4: Automation — add a "Rendering to video, PDF,
  or image formats" section to the user guide clarifying that python-pptx
  does not render slides and pointing at ``libreoffice --headless
  --convert-to``, PowerPoint COM automation, Aspose.Slides, and
  python-pptx-interface as integration points for downstream rendering.
- docs: #501 Animated GIFs only showing first frame — close as wontfix
  (renderer-capability limitation, not a library bug). Add an
  "Animated GIFs" section to the user guide explaining that the ``<p:pic>``
  XML python-pptx emits for a ``.gif`` is byte-for-byte identical to what
  PowerPoint writes itself, that GIF cycling is a renderer behaviour keyed
  off the image MIME type (no ``p:timing`` entry required), and summarising
  which renderers animate GIFs in which modes; point at
  ``add_movie`` + an ``ffmpeg`` ``.gif`` → ``.mp4`` conversion as the
  workaround when a target renderer (notably LibreOffice) does not support
  embedded-GIF animation. Adds a cross-referencing note to the ``add_picture``
  quickstart example.
- #151 xmlchemy - ZeroOrMoreChoice
- security: #1055 harden XML parser and zip reader against malicious input.
  The lxml ``XMLParser`` used to parse every ``.pptx`` part (and the parser
  used to read chart-embedded workbooks) now explicitly disables entity
  resolution, DTD loading, and network access, defeating both "billion
  laughs" entity-expansion attacks and XML external entity (XXE) attacks.
  The zip reader now inspects the central-directory uncompressed-size
  fields before loading members into memory and raises
  ``pptx.exc.PackageTooLargeError`` when the declared total exceeds a
  configurable limit (default 2 GiB; override via the
  ``PPTX_MAX_UNCOMPRESSED_SIZE`` environment variable or the
  ``pptx.opc.serialized.MAX_UNCOMPRESSED_PACKAGE_SIZE`` module attribute).
  A new ``docs/dev/security.rst`` documents the trust model and the
  defenses applied.
- #1066 Create 16x9 presentations by default? ``Presentation(pptx_format=...)``
  now accepts ``"16x9"`` / ``"widescreen"`` or ``"4x3"`` / ``"standard"`` to
  select a built-in template aspect ratio; a new ``default-16x9.pptx`` ships
  alongside the existing 4x3 default.
- feat: #355 Font embedding. ``Presentation.embed_font(font_file, typeface,
  style="regular")`` adds a TrueType / OpenType font file to the presentation
  as a ``ppt/fonts/font{n}.fntdata`` part and registers it under an
  ``p:embeddedFontLst/p:embeddedFont`` entry for `typeface`. Supports the four
  PowerPoint style slots (``regular``, ``bold``, ``italic``, ``boldItalic``).
  ``Presentation.embedded_fonts`` returns the tuple of embedded typeface
  names.
- #349 feature: Axis.position
- feat: #132 ``Slides.duplicate(slide, index=None)`` to clone a slide within
  its own presentation. The duplicate inherits from the same slide layout as
  the source, carries a deep copy of the shape tree, and shares (by
  relationship reuse) the source's image, chart, OLE-object, media, and
  hyperlink parts — no content is re-embedded. A freshly-allocated slide-id
  is assigned; the source's slide-id is unchanged. Notes slides on the
  source are not copied onto the duplicate (notes slides carry a
  back-reference to their owning slide and can't be shared); accessing
  ``notes_slide`` on the duplicate creates a fresh empty one on demand.
- Foundation: cross-part rel cloning helper
- Foundation: DrawingML effectLst descriptor family
- Foundation: mc:AlternateContent traversal
- Foundation F6: consolidation docs for slide-id management (shipped via
  #68/#67/#132/#1036). Extracted the shared reposition logic used by
  ``Slides.move_slide`` and ``Slides.duplicate(index=...)`` into a single
  ``Slides._reposition_sldId`` helper, added a regression test covering a
  move → delete → duplicate → external-insert sequence to pin the
  "deleted-id not recycled while a higher-numbered slide survives"
  invariant, and documented the sldId allocator, the F6 API surface, and
  its interaction with F1 at ``docs/dev/analysis/f6-slide-id-manager.rst``.
- fix: #490 ``Chart.replace_data`` raised ``KeyError: 'rId3'`` when the chart's
  ``c:externalData`` element referenced a relationship that wasn't present in
  the chart-part's rels (charts pasted from pre-2007 ``.xls`` workbooks, or
  charts whose embedded-workbook relationship was stripped by another client).
  ``ChartWorkbook.xlsx_part`` now returns |None| for an unresolved rId so
  ``replace_data`` transparently synthesizes a fresh embedded workbook.
- docs: F4 foundation + #583 scaffold — design analysis for the chartex
  (``cx:``) namespace at ``docs/dev/analysis/chartex-foundation.rst``,
  documenting the content type, relationship type, part class, OXML element
  hierarchy, ``XL_CHART_TYPE`` layoutId mapping, and spec references
  required before Office 2016+ extended chart types (funnel, treemap,
  sunburst, waterfall, histogram, box-and-whisker, map) can be created or
  read in detail. Companion to #386 (chartex passthrough) which ships the
  ``UNSUPPORTED_CHARTEX`` sentinel.
- feat: #130 Shape.shadow and ShadowFormat object. ``ShadowFormat`` now exposes
  the full outer-shadow property set (``blur_radius``, ``distance``,
  ``direction``, ``color``) in addition to ``inherit``, and a new
  ``ChartFormat.shadow`` property surfaces the same API on chart elements like
  ``Axis.format``, ``Series.format``, and ``MajorGridlines.format``.
- feat: #71 cell borders. ``_Cell`` gains ``border_left``, ``border_right``,
  ``border_top``, ``border_bottom``, ``border_diagonal_down`` and
  ``border_diagonal_up`` |LineFormat| properties, each writing to the
  corresponding ``a:lnL`` / ``a:lnR`` / ``a:lnT`` / ``a:lnB`` / ``a:lnTlToBr``
  / ``a:lnBlToTr`` child of ``a:tcPr`` so explicit per-cell edge and diagonal
  borders can be set in color, width, and dash style. Borders inherited from
  the applied table style are not reported by these properties.
- fix: #925 shape of group type has incorrect size. ``BaseShape`` gains
  read-only ``effective_left``, ``effective_top``, ``effective_width`` and
  ``effective_height`` properties that return the shape's slide-relative
  geometry after the enclosing ``p:grpSp`` ancestors' ``a:chOff``/``a:chExt``
  → ``a:off``/``a:ext`` transforms have been composited. The pre-existing
  ``left``/``top``/``width``/``height`` properties continue to expose the raw
  XML values (in the enclosing group's child coordinate system for a nested
  shape) for backwards compatibility.
- fix: #347 Doughnut chart not displaying data labels. The default
  ``c:dLbls`` block emitted for a newly-added doughnut or exploded-doughnut
  chart set ``c:showVal val="0"``, so PowerPoint had the element wired up
  but every show-flag turned off and drew no labels. The default now emits
  ``c:showVal val="1"`` so the numeric values render the first time the
  chart is opened, matching PowerPoint's own "Add Data Labels" behavior.
- fix: #396 ``Chart.replace_data()`` now validates that the supplied
  |ChartData| subclass matches the chart's type family and raises
  ``ValueError`` pointing at the correct subclass instead of silently
  writing the wrong XML shape (which caused PowerPoint to offer to
  "repair" the file on open). XY/scatter charts require
  ``XyChartData``, bubble charts require ``BubbleChartData``, and all
  other chart types require ``CategoryChartData``.
- fix: #749 ``Shape.auto_shape_type`` raised ``KeyError: 'line'`` on an
  auto-shape whose ``a:prstGeom`` element had ``prst="line"``. A new
  ``MSO_SHAPE.LINE`` enum member now maps the ``"line"`` preset, and
  ``Shape.auto_shape_type`` returns |None| (instead of raising) when the
  preset value is not a known member of ``MSO_AUTO_SHAPE_TYPE``.
- fix: #332 MSO_LINE_DASH_STYLE.ROUND_DOT correct mapping
- #787 Support .MPO image files
- #849 expose `_Cell.row_idx` and `col_idx`
- fix: #1042 Image.content_type and Image.ext incorrect for EMF files
- #702 reproducible builds via fixed zip timestamps
- build: #1103 update pyparsing usage for 3.x
- docs: #296 document that ``Table.height`` and ``_Row.height`` are authored
  (minimum) values. PowerPoint's layout engine grows rows at render time to
  fit their text content; python-pptx cannot reproduce that calculation, so
  the height reported by the library may be less than the rendered height
  until the file is opened and saved by PowerPoint (known limitation /
  wontfix). The user guide gains a new "Table height and row height" section
  describing the limitation and suggested workarounds.
- fix: #611 (resolved with #332) ``MSO_LINE_DASH_STYLE.ROUND_DOT`` had been
  mapped to the ``sysDot`` preset, which the #611 reporter pointed to as
  "correct" while others pointed to the opposite (``dot``). Per
  ``ST_PresetLineDashVal`` in ``dml-main.xsd`` and PowerPoint's own output
  for ``msoLineRoundDot``, ``dot`` is the correct preset; the mapping was
  corrected as part of the #332 fix above and a regression test now pins the
  mapping so a naive revert per #611's proposal cannot land silently.
- fix: #674 correct UP_DOWN_ARROW adjustment list
- fix: #773 guard fit_text() when no fitting layout is found
- #715 python-pptx fit text within text placeholder. ``TextFrame`` gains
  read/write ``font_scale`` and ``line_space_reduction`` properties that
  expose the ``fontScale`` and ``lnSpcReduction`` attributes of
  ``a:normAutofit``. Assigning either attribute ensures an ``a:normAutofit``
  child is present on ``a:bodyPr``, replacing any existing ``a:noAutofit``
  or ``a:spAutoFit`` choice sibling. This lets callers pre-compute the
  autofit hints that PowerPoint reads on first render, so placeholder text
  appears correctly scaled without requiring the user to edit the box. The
  existing ``TextFrame.fit_text()`` convenience method already works on
  placeholder text frames because placeholder width/height are inherited
  through the layout/master chain.
- #776 Point.invert_if_negative
- perf: #644 Poor performance when creating a big presentation. Part-name
  allocation now caches per-template allocations instead of scanning the full
  part graph on every addition, changing ``Package.next_partname()`` from
  amortized O(N) to O(1) and overall presentation build-up from O(N**2) to O(N).
- feat: #832 add ``_RowCollection.add()`` to append a row to a table
- feat: #837 (and #791) delete a specific row from a table via
  ``_Row.delete()`` or ``_RowCollection.remove(row)``
- #769 lookup placeholder by idx
- feat: #49 shape: set z-order of shape in slide. ``BaseShape`` gains
  ``bring_to_front()``, ``send_to_back()``, ``bring_forward()`` and
  ``send_backward()`` mutators plus a read-only ``zorder_index`` property that
  reports the shape's zero-based position within its parent shape-tree
  (``p:spTree`` or ``p:grpSp``).
- #946 Connector.adjustments for elbow/curved connectors
- #420 universal ``to_rgb()`` across color classes
- #308 Resolve RGB values of theme color. ``ColorFormat.to_rgb()`` now
  applies any ``<a:lumMod>`` / ``<a:lumOff>`` luminance modifiers
  (tint/shade, set via ``ColorFormat.brightness``) to the resolved base
  color, so the returned ``RGBColor`` matches the RGB PowerPoint actually
  renders — not just the unmodified scheme / preset / sRGB base. The
  transform applies across every color type (``srgbClr``, ``schemeClr``,
  ``prstClr``, ``hslClr``, ``scrgbClr``, ``sysClr``)
  via the standard ECMA-376 HSL formula ``new_L = old_L * lumMod + lumOff``.
- #716 DataLabel border and fill via ChartFormat
- fix: #936 text_frame.fit_text crashes if no wrapped representation fits
        in the width of the shape
- #100 paragraph bullet API
- #114 paragraph bullet font, color, and size
- #141 secondary value axis support
- feat: #544 Add confidence-interval / error-bar support for chart series.
  New ``Series.error_bars`` and ``Series.has_error_bars`` read-properties
  expose a new ``ErrorBars`` wrapper, and ``Series.set_error_bars(...)``
  attaches a fresh set of error bars configured from the ``XL_ERROR_BAR_TYPE``,
  ``XL_ERROR_BAR_INCLUDE``, and ``XL_ERROR_BAR_DIRECTION`` enums. Supports
  the fixed-value, percentage, standard-deviation, and standard-error modes;
  custom per-point magnitudes are recognized but not populated in this MVP.
- feat: #381 How to get refresh the showing of picture or chart when I
  updated the embedded excel data? — add
  :meth:`Chart.update_cached_values` which re-reads the embedded Excel
  workbook and rewrites the ``c:numCache`` / ``c:strCache`` entries in
  chart XML so the chart displays the workbook's current values without
  requiring PowerPoint's own *Refresh Data* action.
- fix: #529 Charts not using theme/Accent colors on new series. When
  ``Chart.replace_data()`` adds series beyond the original count by cloning
  the last existing ``c:ser``, any explicit ``a:solidFill/a:srgbClr`` color
  carried over from the source is now replaced with a cycling
  ``a:schemeClr val="accent1..6"`` reference so newly-added series follow
  PowerPoint's theme-accent rotation.
- fix: #1072 add ``DataLabels.text_frame`` so collection-level data-label
  text-body properties (word-wrap, auto-size, vertical anchor, margins) are
  written to ``c:dLbls/c:txPr`` per spec rather than being unreachable via a
  misplaced ``c:tx/c:rich`` path. ``data_labels.text_frame.word_wrap = False``
  now round-trips through PowerPoint as intended.
- #502 support audio MIME types / docs clarification
- fix: #323 Cannot add a video slide when another slide-layout in the master
  template contains an ``audio/mpeg`` (mp3) part. Resolved by #502 and #734 —
  audio MIME-types now load as ``MediaPart`` (gaining the ``.sha1`` attribute
  ``_MediaParts._find_by_sha1`` relies on), and that lookup now defensively
  skips any non-``MediaPart`` entry so a future unmapped media MIME-type
  cannot regress the original ``AttributeError``. Regression tests pin both
  layers.
- fix: #926 ``SlideShapes.add_movie()`` raised ``AttributeError: 'Part' object
  has no attribute 'sha1'`` when the slide part already contained an audio
  clip. Now resolved via the audio content-type registrations added for #502
  (pre-existing audio parts load as ``MediaPart`` with a ``sha1`` attribute)
  together with a defensive skip in ``_MediaParts._find_by_sha1()`` for any
  media rel whose target is a generic ``Part`` — regression tests added.
- feat: #752 accept arbitrary `prog_id` + `extension` in
  `SlideShapes.add_ole_object()` to embed zip/pdf/html/custom files
- #446 shape.shadow.inherit attribute not working
- Add #176 Fitting images to placeholder — `PicturePlaceholder.insert_picture()`
  now accepts a `crop=False` keyword argument that scales the image to fit
  inside the placeholder bounds preserving aspect ratio, with no cropping.
- feat: #199 add ``SlidePlaceholder.insert_chart()`` so a chart can be
  inserted into any slide placeholder, not just a specialized
  ``ChartPlaceholder``.
- feat: #333 Support content placeholders — ``SlidePlaceholder`` (the class
  returned for generic "content", "body", and "object" placeholders) now
  supports the full ``insert_chart()``, ``insert_picture()``, and
  ``insert_table()`` rich-content insertion API. The three methods have been
  lifted onto the common ``_BaseSlidePlaceholder`` so specialized
  ``ChartPlaceholder``, ``PicturePlaceholder``, and ``TablePlaceholder``
  classes continue to expose the same interface while any generic slide
  placeholder can now be populated with a chart, picture, or table without
  needing the specialized subclass.
- fix: #652 Errors when loading a SVG image into a picture placeholder.
  ``SlideShapes.add_picture()`` and ``PicturePlaceholder.insert_picture()``
  now detect SVG content in the incoming image and raise the new
  ``pptx.exc.UnsupportedImageTypeError`` with a message pointing the caller
  at the pre-rasterize workaround. Previously, an opaque
  ``PIL.UnidentifiedImageError`` bubbled up from Pillow. See the new
  "Inserting SVG images" section of the user guide.
- #337 support East-Asian and complex-script font slots
- #375 line end arrows via LineFormat
- feat: #1070 read ``.potx`` (and ``.ppsx``) presentation packages
- fix: #288 add_table accepts float dimensions
- fix: #608 Name of placeholder in layout gets reset to placeholder basename when adding slide
- #68 feature: reorder a slide
- #67 feature: delete a slide
- #41 Shape.delete() method
- #533 Shape.duplicate() for simple shapes
- #515 Expose preset-shape path geometry
- #1036 move slide across presentations (basic)
- fix: #131 Can't preview presentation in gmail — ``docProps/app.xml`` now has
  its ``<Slides>`` count recomputed from the live ``sldIdLst`` at save-time,
  restoring Gmail attachment preview and other downstream tools that key on
  that count.
- #574 Font.strikethrough property
- #940 Set/change font color when working with Hyperlinks is impossible
  (add ``Font.use_theme_hyperlink_color`` flag and
  ``ThemePart.theme.hlink_color`` / ``folHlink_color`` accessors)
- #201 slide numbers / date / footer via p:hf and a:fld
- feat: #338 combo charts — add ``Chart.add_plot(chart_type, chart_data)`` so
  a second plot (e.g. a line overlaid on a column chart) can be appended to
  an existing chart sharing the same axes. MVP supports bar/column and line
  chart types for the added plot; series values are emitted inline
  (``c:numLit`` / ``c:strLit``) rather than appended to the embedded
  workbook, so PowerPoint's *Edit Data* dialog exposes only the original
  plot's data range. See ``docs/dev/analysis/combo-chart.rst`` for the
  embedded-workbook-handler (foundation F5) follow-up that would sync the
  workbook.
- #516 Chart Colors Sometimes Using Extended Alternates
- #386 surface unsupported chartex chart types without dropping
- feat: #734 add ``<a:snd>`` (sound) support to click-action
- #487 read/write legacy PowerPoint comments
- Add #126 read-only access to OMML equations via
  ``Shape.has_math_equation`` and ``Shape.math_equation_xml``. Writing
  equations and LaTeX/MathML conversion remain deferred.
- docs: #892 the "parse equations on a slide" use case is resolved by
  #126's read-only API; add a regression test pinning the
  find-equation-in-slide workflow and a design analysis at
  ``docs/dev/analysis/omml-parsing.rst`` that sketches the remaining
  structured-parsing work (typed ``Equation`` proxy, OMML→MathML / LaTeX
  conversion) kept deferred until a concrete use case surfaces.
- feat: #27 apply table styles. Adds ``Table.style_id`` read/write property
  exposing the GUID stored in ``a:tblPr/a:tableStyleId``. The typed,
  by-name API for selecting built-in PowerPoint table styles is deferred to
  a follow-up release; assignment of unknown GUIDs is not validated against
  the ``tableStyles.xml`` part.
- feat: #668 read and write password-protected (ECMA-376 Agile Encryption)
  ``.pptx`` files. ``Presentation(pptx, password=...)`` decrypts on open and
  ``prs.save(pkg_file, password=...)`` encrypts on save. Unifies the
  ``password`` keyword with the ``zip_date_time`` keyword added for #702;
  both are accepted together and are orthogonal — ``zip_date_time`` applies
  to the inner (plaintext) zip members, then the package is wrapped in an
  OLE2 / Agile-Encryption CFBF container when ``password`` is given.
  Requires the optional ``msoffcrypto-tool`` dependency at runtime (no new
  dependency when the feature is not used); a clear
  ``pptx.exc.EncryptedPackageError`` is raised when the dependency is
  absent or the password is wrong.

1.0.2 (2024-08-07)
++++++++++++++++++

- fix: #1003 restore read-only enum members

1.0.1 (2024-08-05)
++++++++++++++++++

- fix: #1000 add py.typed

1.0.0 (2024-08-03)
++++++++++++++++++

- fix: #929 raises on JPEG with image/jpg MIME-type
- fix: #943 remove mention of a Px Length subtype
- fix: #972 next-slide-id fails in rare cases
- fix: #990 do not require strict timestamps for Zip
- Add type annotations

0.6.23 (2023-11-02)
+++++++++++++++++++

- fix: #912 Pillow<=9.5 constraint entails security vulnerability

0.6.22 (2023-08-28)
+++++++++++++++++++

- Add #909 Add imgW, imgH params to `shapes.add_ole_object()`
- fix: #754 _Relationships.items() raises
- fix: #758 quote in autoshape name must be escaped
- fix: #746 update Python 3.x support in docs
- fix: #748 setup's `license` should be short string
- fix: #762 AttributeError: module 'collections' has no attribute 'abc'
       (Windows Python 3.10+)

0.6.21 (2021-09-20)
+++++++++++++++++++

- Fix #741 _DirPkgReader must implement .__contains__()

0.6.20 (2021-09-14)
+++++++++++++++++++

- Fix #206 accommodate NULL target-references in relationships.
- Fix #223 escape image filename that appears as literal in XML.
- Fix #517 option to display chart categories/values in reverse order.
- Major refactoring of ancient package loading code.

0.6.19 (2021-05-17)
+++++++++++++++++++

- Add shapes.add_ole_object(), allowing arbitrary Excel or other binary file to be
  embedded as a shape on a slide. The OLE object is represented as an icon.

0.6.18 (2019-05-02)
+++++++++++++++++++

- .text property getters encode line-break as a vertical-tab (VT, '\v', ASCII 11/x0B).
  This is consistent with PowerPoint's copy/paste behavior and allows like-breaks (soft
  carriage-return) to be distinguished from paragraph boundary. Previously, a line-break
  was encoded as a newline ('\n') and was not distinguishable from a paragraph boundary.

  .text properties include Shape.text, _Cell.text, TextFrame.text, _Paragraph.text and
  _Run.text.

- .text property setters accept vertical-tab character and place a line-break element in
  that location. All other control characters other than horizontal-tab ('\t') and
  newline ('\n') in range \x00-\x1F are accepted and escaped with plain-text like
  "_x001B" for ESC (ASCII 27).

  Previously a control character other than tab or newline in an assigned string would
  trigger an exception related to invalid XML character.

0.6.17 (2018-12-16)
+++++++++++++++++++

- Add SlideLayouts.remove() - Delete unused slide-layout
- Add SlideLayout.used_by_slides - Get slides based on this slide-layout
- Add SlideLayouts.index() - Get index of slide-layout in master
- Add SlideLayouts.get_by_name() - Get slide-layout by its str name

0.6.16 (2018-11-09)
+++++++++++++++++++

- Feature #395 DataLabels.show_* properties, e.g. .show_percentage
- Feature #453 Chart data tolerates None for labels

0.6.15 (2018-09-24)
+++++++++++++++++++

- Fix #436 ValueAxis._cross_xAx fails on c:dateAxis

0.6.14 (2018-09-24)
+++++++++++++++++++

- Add _Cell.merge()
- Add _Cell.split()
- Add _Cell.__eq__()
- Add _Cell.is_merge_origin
- Add _Cell.is_spanned
- Add _Cell.span_height
- Add _Cell.span_width
- Add _Cell.text getter
- Add Table.iter_cells()
- Move pptx.shapes.table module to pptx.table
- Add user documentation 'Working with tables'

0.6.13 (2018-09-10)
+++++++++++++++++++

- Add Chart.font
- Fix #293 Can't hide title of single-series Chart
- Fix shape.width value is not type Emu
- Fix add a:defRPr with c:rich (fixes some font inheritance breakage)

0.6.12 (2018-08-11)
+++++++++++++++++++

- Add Picture.auto_shape_type
- Remove Python 2.6 testing from build
- Update dependencies to avoid vulnerable Pillow version
- Fix #260, #301, #382, #401
- Add _Paragraph.add_line_break()
- Add Connector.line

0.6.11 (2018-07-25)
+++++++++++++++++++

- Add gradient fill.
- Add experimental "turbo-add" option for producing large shape-count slides.

0.6.10 (2018-06-11)
+++++++++++++++++++

- Add `shape.shadow` property to autoshape, connector, picture, and group
  shape, returning a `ShadowFormat` object.
- Add `ShadowFormat` object with read/write (boolean) `.inherit` property.
- Fix #328 add support for 26+ series in a chart

0.6.9 (2018-05-08)
++++++++++++++++++

- Add `Picture.crop_x` setters, allowing picture cropping values to be set,
  in addition to interrogated.
- Add `Slide.background` and `SlideMaster.background`, allowing the
  background fill to be set for an individual slide or for all slides based
  on a slide master.
- Add option `shapes` parameter to `Shapes.add_group_shape`, allowing a group
  shape to be formed from a number of existing shapes.
- Improve efficiency of `Shapes._next_shape_id` property to improve
  performance on high shape-count slides.

0.6.8 (2018-04-18)
++++++++++++++++++

- Add `GroupShape`, providing properties specific to a group shape, including
  its `shapes` property.
- Add `GroupShapes`, providing access to shapes contained in a group shape.
- Add `SlideShapes.add_group_shape()`, allowing a group shape to be added to
  a slide.
- Add `GroupShapes.add_group_shape()`, allowing a group shape to be added to
  a group shape, enabling recursive, multi-level groups.
- Add support for adding jump-to-named-slide behavior to shape and run
  hyperlinks.

0.6.7 (2017-10-30)
++++++++++++++++++

- Add `SlideShapes.build_freeform()`, allowing freeform shapes (such as maps)
  to be specified and added to a slide.
- Add support for patterned fills.
- Add `LineFormat.dash_style` to allow interrogation and setting of dashed
  line styles.

0.6.6 (2017-06-17)
++++++++++++++++++

- Add `SlideShapes.add_movie()`, allowing video media to be added to a slide.

- fix #190 Accommodate non-conforming part names having '00' index segment.
- fix #273 Accommodate non-conforming part names having no index segment.
- fix #277 ASCII/Unicode error on non-ASCII multi-level category names
- fix #279 BaseShape.id warning appearing on placeholder access.

0.6.5 (2017-03-21)
++++++++++++++++++

- #267 compensate for non-conforming PowerPoint behavior on c:overlay element

- compensate for non-conforming (to spec) PowerPoint behavior related to
  c:dLbl/c:tx that results in "can't save" error when explicit data labels
  are added to bubbles on a bubble chart.

0.6.4 (2017-03-17)
++++++++++++++++++

- add Chart.chart_title and ChartTitle object
- #263 Use Number type to test for numeric category

0.6.3 (2017-02-28)
++++++++++++++++++

- add DataLabel.font
- add Axis.axis_title

0.6.2 (2017-01-03)
++++++++++++++++++

- add support for NotesSlide (slide notes, aka. notes page)
- add support for arbitrary series ordering in XML
- add Plot.categories providing access to hierarchical categories in an
  existing chart.
- add support for date axes on category charts, including writing a dateAx
  element for the category axis when ChartData categories are date or
  datetime.

**BACKWARD INCOMPATIBILITIES:**

Some changes were made to the boilerplate XML used to create new charts. This
was done to more closely adhere to the settings PowerPoint uses when creating
a chart using the UI. This may result in some appearance changes in charts
after upgrading. In particular:

* Chart.has_legend now defaults to True for Line charts.
* Plot.vary_by_categories now defaults to False for Line charts.

0.6.1 (2016-10-09)
++++++++++++++++++

- add Connector shape type

0.6.0 (2016-08-18)
++++++++++++++++++

- add XY chart types
- add Bubble chart types
- add Radar chart types
- add Area chart types
- add Doughnut chart types
- add Series.points and Point
- add Point.data_label
- add DataLabel.text_frame
- add DataLabel.position
- add Axis.major_gridlines
- add ChartFormat with .fill and .line
- add Axis.format (fill and line formatting)
- add ValueAxis.crosses and .crosses_at
- add Point.format (fill and line formatting)
- add Slide.slide_id
- add Slides.get() (by slide id)
- add Font.language_id
- support blank (None) data points in created charts
- add Series.marker
- add Point.marker
- add Marker.format, .style, and .size

0.5.8 (2015-11-27)
++++++++++++++++++

- add Shape.click_action (hyperlink on shape)
- fix: #128 Chart cat and ser names not escaped
- fix: #153 shapes.title raises on no title shape
- fix: #170 remove seek(0) from Image.from_file()

0.5.7 (2015-01-17)
++++++++++++++++++

- add PicturePlaceholder with .insert_picture() method
- add TablePlaceholder with .insert_table() method
- add ChartPlaceholder with .insert_chart() method
- add Picture.image property, returning Image object
- add Picture.crop_left, .crop_top, .crop_right, and .crop_bottom
- add Shape.placeholder_format and PlaceholderFormat object

**BACKWARD INCOMPATIBILITIES:**

Shape.shape_type is now unconditionally `MSO_SHAPE_TYPE.PLACEHOLDER` for all
placeholder shapes. Previously, some placeholder shapes reported
`MSO_SHAPE_TYPE.AUTO_SHAPE`, `MSO_SHAPE_TYPE.CHART`,
`MSO_SHAPE_TYPE.PICTURE`, or `MSO_SHAPE_TYPE.TABLE` for that property.

0.5.6 (2014-12-06)
++++++++++++++++++

- fix #138 - UnicodeDecodeError in setup.py on Windows 7 Python 3.4

0.5.5 (2014-11-17)
++++++++++++++++++

- feature #51 - add Python 3 support

0.5.4 (2014-11-15)
++++++++++++++++++

- feature #43 - image native size in shapes.add_picture() is now calculated
  based on DPI attribute in image file, if present, defaulting to 72 dpi.
- feature #113 - Add Paragraph.space_before, Paragraph.space_after, and
  Paragraph.line_spacing

0.5.3 (2014-11-09)
++++++++++++++++++

- add experimental feature TextFrame.fit_text()

0.5.2 (2014-10-26)
++++++++++++++++++

- fix #127 - Shape.text_frame fails on shape having no txBody

0.5.1 (2014-09-22)
++++++++++++++++++

- feature #120 - add Shape.rotation
- feature #97 - add Font.underline
- issue #117 - add BMP image support
- issue #95 - add BaseShape.name setter
- issue #107 - all .text properties should return unicode, not str
- feature #106 - add .text getters to Shape, TextFrame, and Paragraph

- Rename Shape.textframe to Shape.text_frame.
  **Shape.textframe property (by that name) is deprecated.**

0.5.0 (2014-09-13)
++++++++++++++++++

- Add support for creating and manipulating bar, column, line, and pie charts
- Major refactoring of XML layer (oxml)
- Rationalized graphical object shape access
  **Note backward incompatibilities below**

**BACKWARD INCOMPATIBILITIES:**

A table is no longer treated as a shape. Rather it is a graphical object
contained in a GraphicFrame shape, as are Chart and SmartArt objects.

Example::

    table = shapes.add_table(...)

    # becomes

    graphic_frame = shapes.add_table(...)
    table = graphic_frame.table

    # or

    table = shapes.add_table(...).table

As the enclosing shape, the id, name, shape type, position, and size are
attributes of the enclosing GraphicFrame object.

The contents of a GraphicFrame shape can be identified using three available
properties on a shape: has_table, has_chart, and has_smart_art. The enclosed
graphical object is obtained using the properties GraphicFrame.table and
GraphicFrame.chart. SmartArt is not yet supported. Accessing one of these
properties on a GraphicFrame not containing the corresponding object raises
an exception.

0.4.2 (2014-04-29)
++++++++++++++++++

- fix: issue #88 -- raises on supported image file having uppercase extension
- fix: issue #89 -- raises on add_slide() where non-contiguous existing ids

0.4.1 (2014-04-29)
++++++++++++++++++

- Rename Presentation.slidemasters to Presentation.slide_masters.
  Presentation.slidemasters property is deprecated.
- Rename Presentation.slidelayouts to Presentation.slide_layouts.
  Presentation.slidelayouts property is deprecated.
- Rename SlideMaster.slidelayouts to SlideMaster.slide_layouts.
  SlideMaster.slidelayouts property is deprecated.
- Rename SlideLayout.slidemaster to SlideLayout.slide_master.
  SlideLayout.slidemaster property is deprecated.
- Rename Slide.slidelayout to Slide.slide_layout. Slide.slidelayout property
  is deprecated.
- Add SlideMaster.shapes to access shapes on slide master.
- Add SlideMaster.placeholders to access placeholder shapes on slide master.
- Add _MasterPlaceholder class.
- Add _LayoutPlaceholder class with position and size inheritable from master
  placeholder.
- Add _SlidePlaceholder class with position and size inheritable from layout
  placeholder.
- Add Table.left, top, width, and height read/write properties.
- Add rudimentary GroupShape with left, top, width, and height properties.
- Add rudimentary Connector with left, top, width, and height properties.
- Add TextFrame.auto_size property.
- Add Presentation.slide_width and .slide_height read/write properties.
- Add LineFormat class providing access to read and change line color and
  width.
- Add AutoShape.line
- Add Picture.line

- Rationalize enumerations. **Note backward incompatibilities below**

**BACKWARD INCOMPATIBILITIES:**

The following enumerations were moved/renamed during the rationalization of
enumerations:

- ``pptx.enum.MSO_COLOR_TYPE`` --> ``pptx.enum.dml.MSO_COLOR_TYPE``
- ``pptx.enum.MSO_FILL`` --> ``pptx.enum.dml.MSO_FILL``
- ``pptx.enum.MSO_THEME_COLOR`` --> ``pptx.enum.dml.MSO_THEME_COLOR``
- ``pptx.constants.MSO.ANCHOR_*`` --> ``pptx.enum.text.MSO_ANCHOR.*``
- ``pptx.constants.MSO_SHAPE`` --> ``pptx.enum.shapes.MSO_SHAPE``
- ``pptx.constants.PP.ALIGN_*`` --> ``pptx.enum.text.PP_ALIGN.*``
- ``pptx.constants.MSO.{SHAPE_TYPES}`` -->
  ``pptx.enum.shapes.MSO_SHAPE_TYPE.*``

Documentation for all enumerations is available in the Enumerations section
of the User Guide.

0.3.2 (2014-02-07)
++++++++++++++++++

- Hotfix: issue #80 generated presentations fail to load in Keynote and other
  Apple applications

0.3.1 (2014-01-10)
++++++++++++++++++

- Hotfix: failed to load certain presentations containing images with
  uppercase extension

0.3.0 (2013-12-12)
++++++++++++++++++

- Add read/write font color property supporting RGB, theme color, and inherit
  color types
- Add font typeface and italic support
- Add text frame margins and word-wrap
- Add support for external relationships, e.g. linked spreadsheet
- Add hyperlink support for text run in shape and table cell
- Add fill color and brightness for shape and table cell, fill can also be set
  to transparent (no fill)
- Add read/write position and size properties to shape and picture
- Replace PIL dependency with Pillow
- Restructure modules to better suit size of library

0.2.6 (2013-06-22)
++++++++++++++++++

- Add read/write access to core document properties
- Hotfix to accomodate connector shapes in _AutoShapeType
- Hotfix to allow customXml parts to load when present

0.2.5 (2013-06-11)
++++++++++++++++++

- Add paragraph alignment property (left, right, centered, etc.)
- Add vertical alignment within table cell (top, middle, bottom)
- Add table cell margin properties
- Add table boolean properties: first column (row header), first row (column
  headings), last row (for e.g. totals row), last column (for e.g. row
  totals), horizontal banding, and vertical banding.
- Add support for auto shape adjustment values, e.g. change radius of corner
  rounding on rounded rectangle, position of callout arrow, etc.

0.2.4 (2013-05-16)
++++++++++++++++++

- Add support for auto shapes (e.g. polygons, flowchart symbols, etc.)

0.2.3 (2013-05-05)
++++++++++++++++++

- Add support for table shapes
- Add indentation support to textbox shapes, enabling multi-level bullets on
  bullet slides.

0.2.2 (2013-03-25)
++++++++++++++++++

- Add support for opening and saving a presentation from/to a file-like
  object.
- Refactor XML handling to use lxml objectify

0.2.1 (2013-02-25)
++++++++++++++++++

- Add support for Python 2.6
- Add images from a stream (e.g. StringIO) in addition to a path, allowing
  images retrieved from a database or network resource to be inserted without
  saving first.
- Expand text methods to accept unicode and UTF-8 encoded 8-bit strings.
- Fix potential install bug triggered by importing ``__version__`` from
  package ``__init__.py`` file.

0.2.0 (2013-02-10)
++++++++++++++++++

First non-alpha release with basic capabilities:

- open presentation/template or use built-in default template
- add slide
- set placeholder text (e.g. bullet slides)
- add picture
- add text box
