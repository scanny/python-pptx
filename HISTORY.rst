.. :changelog:

Release History
---------------

2026.05.0 — CalVer alignment across the loadfix series
++++++++++++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-02

This release switches to CalVer (YYYY.MM.patch) versioning, matching
loadfix/python-docx and loadfix/python-xlsx.

Unreleased
++++++++++

- docs: #823 add recipe for editing footer/slide-number/date placeholders.
  Issue #823 (https://github.com/scanny/python-pptx/issues/823) asked
  how to edit "the character in the lower-left corner" of slides —
  typically the footer, slide-number, or date placeholder authored on
  the slide master (and optionally overridden on a specific layout).
  The ``docs/user/slides.rst`` "Editing the footer, slide-number, or
  date placeholder text" section now shows how to locate each latent
  placeholder by ``placeholder_format.type`` (portable across masters
  whose ``idx`` values differ), edit its text, and pair the edit with
  a ``header_footer`` visibility toggle. A new
  ``tests/test_issue_823_footer_edit_recipe.py`` regression suite pins
  the code paths the recipe depends on — including the latent-
  placeholders-not-cloned contract and the watermark-shape fallback.
  No public-API change.

- verify: #525 textbox auto-grow to fit text via ``TextFrame.auto_size``.
  Issue #525 (https://github.com/scanny/python-pptx/issues/525) asked for a
  way to have a textbox resize vertically to fit its text. The capability
  has always been present:
  :attr:`~pptx.text.text.TextFrame.auto_size` accepts
  :class:`~pptx.enum.text.MSO_AUTO_SIZE` and writes the corresponding
  ``a:bodyPr`` auto-fit choice — ``a:spAutoFit`` for
  ``SHAPE_TO_FIT_TEXT`` (PowerPoint resizes the shape to fit the text when
  the file is opened), ``a:normAutofit`` for ``TEXT_TO_FIT_SHAPE`` (text
  shrinks to fit the shape), and ``a:noAutofit`` for ``NONE``. Assigning
  ``None`` clears the choice so the setting is inherited from the
  layout/master/theme. The library cannot *compute* the new shape bounds
  itself (that requires per-font metrics at render time), but it emits the
  right XML so PowerPoint performs the resize on open. A new
  ``tests/test_issue_525_shape_autofit_verify.py`` suite pins each
  auto-fit choice and its round-trip through save + reopen, plus the
  single-sibling invariant on reassignment. No code change — the feature
  is the already-existing ``TextFrame.auto_size`` setter.

- verify: #623 text overflow control (wrap, single-line, clip) supported
  via existing :attr:`TextFrame.word_wrap` and :attr:`TextFrame.auto_size`.
  Issue #623 (https://github.com/scanny/python-pptx/issues/623) asked
  for a way to control what happens when a textbox's content is longer
  than fits. All three PowerPoint behaviours are already on the public
  surface: ``word_wrap = True`` wraps at the shape edge
  (``a:bodyPr/@wrap="square"``), ``word_wrap = False`` overflows a
  single line (``a:bodyPr/@wrap="none"``), and
  ``auto_size = MSO_AUTO_SIZE.NONE`` emits ``<a:noAutofit/>`` so
  PowerPoint clips anything outside the shape. A new
  ``tests/test_issue_623_text_overflow_verify.py`` regression suite
  pins each recipe — including the round-trip — and a new
  "Controlling text overflow" section in ``docs/user/text.rst``
  documents the three-way choice.

- verify: #430 ``SlideShapes.add_movie`` accepts MP3 / audio clips.
  Issue #430 (https://github.com/scanny/python-pptx/issues/430) reported
  that passing an MP3 to :meth:`SlideShapes.add_movie` failed — the
  method's name suggests video-only, but in OOXML the same ``p:pic``
  shape carries audio, distinguished only by an ``<a:audioFile>`` child
  in place of ``<a:videoFile>``. The fork-era audio-MIME work already
  shipped support: passing ``mime_type="audio/mpeg"`` (or any other
  ``audio/*`` value such as ``audio/mp3``, ``audio/wav``,
  ``audio/x-wav``) makes ``add_movie`` emit ``<a:audioFile>``, register
  the matching ``audio/*`` content-type override, and produce a clip
  PowerPoint opens as audio. ``tests/test_issue_430_add_movie_mp3_verify.py``
  pins the behavior end-to-end — MP3 by path, MP3 by file-like (the
  reporter's call signature), save + reopen round-trip, and the
  ``audio/*`` prefix sniff across multiple spellings — so the path
  cannot silently regress.

- verify: #514 ``TextFrame.fit_text`` sets ``run.font.size`` on every
  affected run. Issue #514
  (https://github.com/scanny/python-pptx/issues/514) asked for
  :meth:`TextFrame.fit_text` to record the chosen point size on each
  run's ``a:rPr/@sz`` so callers can introspect the result — e.g. to
  align effective sizes across multiple text frames ("use the smaller
  of the two shrunk sizes"). The private ``TextFrame._set_font`` helper
  has always walked every run's ``rPr`` (plus ``endParaRPr``) and
  written ``@sz`` with the fit size, so :attr:`Font.size` returns a
  concrete |Length| rather than ``None`` after the call. The
  ``_extents`` tuple is recomputed from ``shape.width`` / ``height``
  on every invocation, so resizing a shape between ``fit_text()`` calls
  produces a size appropriate to the new extents.
  ``tests/test_issue_514_fit_text_after_resize.py`` pins the
  verify-close contract: ``run.font.size`` is not ``None`` after
  ``fit_text()``, every run in every paragraph shares the same size,
  cross-frame alignment via ``min(size1, size2)`` works, the
  EMU↔centipoint round-trip is exact, a grow-then-refit produces a
  larger size and a shrink-then-refit a smaller one, and the
  ``auto_size = NONE`` / ``word_wrap = True`` side-effects hold on
  every call.

- verify: #556 replace title text with run-level ``__`` placeholders
  while preserving formatting. Issue #556
  (https://github.com/scanny/python-pptx/issues/556) asked how to
  replace a title authored as ``"My title __ placeholder __"`` —
  substituting each double-underscore token with a value — *without*
  losing the bold / italic / size / colour / underline applied to the
  surrounding runs. The capability has shipped since Wave 6 (#836) via
  :meth:`TextFrame.replace_text` and :meth:`_Paragraph.replace_text`,
  which rewrite every occurrence of ``find`` with ``replace`` across
  consecutive ``a:r`` runs while keeping the origin run's ``a:rPr``;
  Wave 8 (#285) pinned the textbox-facing regression. ``slide.shapes.
  title`` returns a :class:`~pptx.shapes.placeholder.SlidePlaceholder`
  whose ``text_frame`` exposes the same API, so
  ``slide.shapes.title.text_frame.replace_text("__", "NAME")`` rewrites
  both ``__`` tokens in one call and returns ``2``.
  ``tests/test_issue_556_title_replace_preserve_verify.py`` pins the
  three-scenario verify-close suite: both tokens rewritten from a
  single formatted run, the same scenario surviving a ``Presentation.
  save`` + reopen round-trip, and a three-run PowerPoint-split layout
  where only the placeholder run is rewritten and the flanking prose
  runs' ``a:rPr`` is left untouched.

- verify: #617 trendlines work on column-chart series. Issue #617
  (https://github.com/scanny/python-pptx/issues/617) reported that
  trendlines could be added to line-chart series but not column-chart
  series. The capability has always been present on this fork:
  :meth:`~pptx.chart.series._BaseSeries.add_trendline` and
  :attr:`~pptx.chart.series._BaseSeries.trendlines` live on the shared
  ``_BaseSeries`` base, which |BarSeries| inherits via
  ``_BaseCategorySeries``. Column charts render as ``c:barChart`` with
  ``c:barDir=col`` — the series class is the same |BarSeries| used by
  bar charts — so every :class:`~pptx.enum.chart.XL_TRENDLINE_TYPE`
  member (LINEAR, LOGARITHMIC, POLYNOMIAL, POWER, EXPONENTIAL,
  MOVING_AVG) attaches, round-trips, and displays on-chart equation /
  R-squared labels exactly as it does on line-chart series.
  ``tests/test_issue_617_column_trendline_verify.py`` pins the
  end-to-end contract — API-surface guard, every trendline type,
  multiple trendlines per series, ``display_equation`` /
  ``display_r_squared``, a ``Presentation.save`` + reopen round-trip,
  the expected ``c:trendline/c:trendlineType`` XML fragment under a
  ``c:barChart``/``c:ser``, and every column-chart variant (clustered,
  stacked, 100% stacked). No code change — the behaviour is the
  existing trendline feature exercised through a column-chart shape.

- docs: #1022 clarify ``ActionSetting.screen_tip`` visibility rules.
  Issue #1022 (https://github.com/scanny/python-pptx/issues/1022)
  reported that assigning ``click_action.screen_tip`` stores the
  tooltip in the XML but PowerPoint does not show it on hover. The
  ``<a:hlinkClick tooltip="..."/>`` element is spec-valid without an
  ``r:id`` / ``action`` target, but PowerPoint only *renders* a
  ScreenTip when the hyperlink element also carries an actionable
  target — a URL, slide jump, embedded sound, or ``ppaction://``
  action verb. This matches PowerPoint's own Insert Hyperlink dialog,
  which will not accept a ScreenTip without a target. The
  ``ActionSetting.screen_tip`` property docstring, the "Mouse-hover
  actions" section of ``docs/user/text.rst``, and a new
  ``tests/test_issue_1022_screen_tip.py`` regression suite now pin
  the library-side contract and document the workaround (pair the
  tooltip with a URL via ``hyperlink.address`` or a slide jump via
  ``target_slide``). No code change to the setter — the observed
  behaviour is genuine PowerPoint semantics, not a library bug.

- verify: #576 picture hyperlink supported via ``BaseShape.click_action``.
  Issue #576 (https://github.com/scanny/python-pptx/issues/576) asked
  whether a caller can attach a hyperlink — either an external URL or a
  jump to another slide — to a :class:`~pptx.shapes.picture.Picture`.
  The capability has always been present: ``Picture`` inherits from
  :class:`~pptx.shapes.base.BaseShape`, which exposes
  :attr:`~pptx.shapes.base.BaseShape.click_action` (and, since Wave 14,
  :attr:`~pptx.shapes.base.BaseShape.hover_action`) —
  :class:`~pptx.action.ActionSetting` proxies bound to the picture's
  ``p:nvPicPr/p:cNvPr`` element. ``pic.click_action.hyperlink.address``
  sets a URL, ``pic.click_action.target_slide`` a slide-jump,
  ``pic.click_action.screen_tip`` a tooltip, and
  ``pic.hover_action.hyperlink.address`` a mouse-over hyperlink; assigning
  |None| clears the action and its relationship cleanly.
  ``tests/test_issue_576_picture_hyperlink_verify.py`` pins the
  ten-scenario verify-close suite — API-surface guard, URL hyperlink
  round-trip, slide-jump round-trip, hover-action round-trip, ScreenTip,
  and clearing both click and hover actions — and
  ``docs/user/understanding-shapes.rst`` now documents the picture
  click/hover recipe inline.

- verify: #613 clarify ``SlideShapes.add_table(height=)`` semantics. Issue
  #613 (https://github.com/scanny/python-pptx/issues/613) reported that
  ``add_table(rows, cols, left, top, width, height)`` looked as if
  ``height`` was the *per-row* height — the reporter expected a
  "total_height" kwarg that would divide across the rows. The argument
  has always been the total (authored) table height; ``CT_Table.new_tbl``
  derives ``rowheight = height // rows`` and the last row absorbs any
  integer-division remainder, so ``sum(row.height) == height`` exactly.
  The :meth:`SlideShapes.add_table` docstring and ``docs/user/table.rst``
  now spell this out unambiguously and cross-reference the "Table height
  and row height" caveat (PowerPoint may grow rows at render time to fit
  wrapped text). ``tests/test_issue_613_add_table_height_verify.py`` pins
  the contract end-to-end: the graphic-frame extent equals ``height``,
  row heights are ``height // rows`` with remainder in the last row, the
  symmetric invariant holds for widths, the ``p:xfrm/a:ext/@cy``
  attribute carries the same value, and the whole shape survives a
  ``Presentation.save`` + reopen unchanged.

- verify: #835 chart-carrying slides round-trip via
  ``add_slide_from_external`` and ``merge``. Issue #835
  (https://github.com/scanny/python-pptx/issues/835) asked to combine
  charts from different slides in different decks. The underlying
  capability — copying a chart-bearing slide from a source presentation
  into a target presentation with the chart's data intact — is served by
  :meth:`Slides.add_slide_from_external` (per-slide) and
  :meth:`Presentation.merge` (bulk), both of which route through the
  Wave 7 #934 cross-package cloner so chart parts get a distinct embedded
  workbook in the target.
  ``tests/test_issue_835_merge_charts_verify.py`` pins the end-to-end
  behaviour: a source deck with three chart slides (distinct series
  values) round-trips through both APIs with
  ``chart.plots[0].series[0].values`` preserved across save + reopen, and
  every merged chart carries its own :class:`EmbeddedXlsxPart` so
  PowerPoint's "Edit Data" dialog operates on the target deck in
  isolation from the source.

- verify: #1099 notes read correctly beyond slide 20. Issue #1099
  (https://github.com/scanny/python-pptx/issues/1099) reported that
  ``slide.notes_slide.notes_text_frame.text`` returned an empty string
  for slides past index 20 in a specific deck and speculated about "a
  known bug due to XML parsing"; the reporter never supplied the deck
  or a repro. Inspecting the code path confirms there is no
  index-dependent logic: :attr:`.Slide.has_notes_slide`,
  :meth:`.Slide.notes_slide`, :attr:`.NotesSlide.notes_placeholder`,
  and :attr:`.NotesSlide.notes_text_frame` all resolve the notes slide
  via the slide-part's ``RT.NOTES_SLIDE`` relationship and then iterate
  the notes-slide placeholders — no step consults the slide's ordinal
  in the presentation's slide-id list, so there is no mechanism that
  could drop notes past "slide 20".
  ``tests/test_issue_1099_notes_after_20_verify.py`` pins the scenario
  end-to-end: it authors 25- and 50-slide decks (including a mix of
  short, long, and Unicode-heavy notes), saves, reopens via
  ``Presentation(buf)``, and asserts every slide — with explicit
  assertions on the slides past index 20 — round-trips its notes text
  exactly. A parametrized case re-runs the pin at 21, 25, 30, and 40
  slides. Any future regression that actually dropped notes past an
  arbitrary slide ordinal would be caught by these.

- verify: #939 ``_BulletFormat`` readers pinned. Issue #939
  (https://github.com/scanny/python-pptx/issues/939) asked for a way
  to *inspect* whether a paragraph has auto-numbered bullets enabled;
  the companion authoring surface landed as ``_BulletFormat.auto_number``
  via #100. This fork already ships the full symmetric reader on
  ``_Paragraph.bullet``: ``.type`` returns ``"char"`` / ``"autonum"``
  / ``"none"`` / |None| (the last meaning "no explicit setting;
  inherits from the style hierarchy"); ``.char`` returns the literal
  bullet character for a character bullet; ``.number_scheme`` returns
  a :class:`PP_AUTO_NUMBER_SCHEME` member for an autonum bullet; and
  ``.start_at`` returns the ECMA-376 ordinal (defaulting to 1 when
  ``a:buAutoNum/@startAt`` is absent). ``tests/test_issue_939_autonum_inspect_verify.py``
  adds a ``DescribeIssue939AutonumInspectVerify`` suite that pins
  each reader against the real public API, covering the three
  explicit bullet types, the inherited-bullet case (every reader
  returns |None|), the ECMA default for ``start_at``, the
  ``PP_AUTO_NUMBER`` alias, and a :meth:`Presentation.save` + reopen
  round-trip of all three configurations.

- fix: #1026 ``TextFrame.fit_text()`` no longer crashes when the text
  contains characters the chosen font cannot measure — arrows (``→`` /
  ``←`` / ``↑`` / ``↓``), CJK ideographs, emoji, or any codepoint that
  causes Pillow's ``ImageFont.getbbox`` to raise. The measurement path
  in ``src/pptx/text/layout.py`` now falls back to per-character
  measurement, substituting the width of ``?`` (the OpenType-required
  ``.notdef`` fallback glyph) for any char that still fails, so a sensible
  integer point size is returned instead of an uncaught
  ``UnicodeEncodeError`` / ``OSError`` / ``ValueError``. Genuine font-load
  failures continue to raise ``TextLayoutError`` per #168.

- fix: #1127 rename ``ERCENT_40`` to ``PERCENT_40`` in
  ``MSO_PATTERN_TYPE`` (alias retained for backcompat). The original
  member name was missing its leading ``P``; the canonical name is now
  ``PERCENT_40`` and existing code using ``ERCENT_40`` continues to
  resolve to the same member (value ``6``, ``xml_value="pct40"``). The
  alias is documented as deprecated and will be removed in a future
  major release.

- Add #96 ``SlideShapes.clear(preserve_placeholders=True)`` and
  convenience wrapper ``Slide.clear_shapes(preserve_placeholders=True)``
  that bulk-remove every top-level shape from a slide. Placeholders are
  preserved by default so the slide continues to inherit its layout;
  pass ``preserve_placeholders=False`` to remove everything. Per-shape
  ``.delete()`` side effects run in the normal way (a removed picture
  drops its image relationship, a removed chart drops its embedded
  chart part, a removed group is detached together with every shape
  inside it). See ``docs/user/understanding-shapes.rst``.

- feat: #377 add ``Slides.get_by_slide_id(slide_id, default=None)``.
  Issue #377 (https://github.com/scanny/python-pptx/issues/377) asked for
  an explicit-named lookup method on the slide collection that mirrors
  :meth:`.SlideMaster.get_layout` / :meth:`.SlideLayouts.get_by_id`. The
  new method walks ``p:sldIdLst/p:sldId`` for a matching ``@id`` and
  resolves the ``r:id`` relationship to a |Slide|, returning ``default``
  (``None`` by default) when no entry matches. It is an alias for the
  pre-existing :meth:`.Slides.get` chosen for API-shape parity with its
  layout-side siblings; callers that stashed a stable ``slide.slide_id``
  can round-trip the id back to the |Slide| object regardless of
  reordering.

- feat: #435 add :meth:`.SlideShapes.iter_leaf_shapes` — a leaf-only
  variant of :meth:`.SlideShapes.descendants` that yields every non-
  :class:`.GroupShape` descendant in document (z-order) sequence. Useful
  when you want to touch every drawable element on a slide once (collect
  alt-text, audit fills, export a shape-by-shape report) without having
  to special-case the group containers. Walks into nested groups at every
  depth; for a slide that contains no groups it yields the same sequence
  as ``iter(shapes)``.

- feat: #134 ``TextFrame.add_paragraph(text, *, bold, italic, size, color,
  font_name)`` and ``_Paragraph.add_run(text, *, bold, italic, size, color,
  font_name)`` accept the text and common run-level formatting in one call,
  collapsing the previous three-line ``p = tf.add_paragraph(); p.text = ...``
  / ``run = p.add_run(); run.text = ...; run.font.bold = True`` idiom. All
  formatting kwargs are keyword-only; ``color`` accepts an ``RGBColor`` or a
  member of ``MSO_THEME_COLOR``. Passing a run-level kwarg to
  ``add_paragraph`` without ``text`` raises ``ValueError`` (there is no run
  to apply it to). The bare no-argument calls remain fully backwards
  compatible.

- feat: #753 add ``_Paragraph.write_rich(*parts)`` rich-text authoring
  helper. Issue #753
  (https://github.com/scanny/python-pptx/issues/753) asked for a more
  ergonomic way to build a paragraph that mixes bold, italic, coloured,
  and plain runs than the per-run ``add_run`` / ``run.font.*`` loop. The
  new method accepts a positional sequence of *parts* — each a ``str``,
  a ``(text, formatting)`` 2-tuple, or a mapping with a ``"text"`` key —
  and appends one ``a:r`` run per part with the formatting applied to the
  run's ``font``. Recognised formatting keys are ``bold``, ``italic``,
  ``underline``, ``size`` (a |Length|, e.g. ``Pt(18)``), ``color`` (an
  ``RGBColor``), and ``font_name``; unknown keys raise ``ValueError`` so
  typos surface at authoring time. The method returns the paragraph to
  support chaining. It is a thin loop over
  :meth:`._Paragraph.add_run`, so the emitted XML is identical to the
  per-run idiom and round-trips through PowerPoint unchanged. See the
  "Rich text in one call" section in ``docs/user/text.rst``.

- feat: #649 hide individual legend entries. Issue #649
  (https://github.com/scanny/python-pptx/issues/649) asked for a way
  to suppress specific legend entries while keeping their series
  plotted — the programmatic equivalent of PowerPoint's right-click
  *Format Legend Entry → Delete* on a single selected entry. The fix
  models PowerPoint's ``c:legend/c:legendEntry/c:delete`` override:
  :meth:`.Legend.exclude_entry` inserts (or promotes) a
  ``c:legendEntry`` for the given 0-based index with an explicit
  ``c:delete val="1"`` child; :meth:`.Legend.include_entry` reverses
  the operation, removing the entire ``c:legendEntry`` when no other
  overrides (e.g. a ``c:txPr`` formatting block) are present;
  :attr:`.Legend.hidden_entries` returns the current set of hidden
  indices as a tuple in document order. The underlying ``c:ser`` is
  untouched so the series continues to contribute to the plot.
  ``tests/test_issue_649_hide_legend_entries.py`` pins the contract —
  4-series chart, hide two non-contiguous entries, all series still
  plotted, ``Legend.hidden_entries`` reports (1, 3), round-trips
  through :meth:`.Presentation.save` + reopen. Documented in the user
  guide under *Working with charts → Hiding individual legend
  entries*.

- feat: #809 add ``Slide.effective_background`` for inheritance-aware
  reads. Issue #809
  (https://github.com/scanny/python-pptx/issues/809) reported that a
  slide whose layout carries a solid-color ``p:bg`` could not be
  interrogated for its rendered color — ``slide.background.fill``
  destructively materializes a ``p:bgPr/a:noFill`` subtree on first
  read, silently clobbering the inheritance link. The new
  ``Slide.effective_background`` (plus ``SlideLayout.effective_background``
  and ``SlideMaster.effective_background``) walks *slide → layout →
  master* and returns a side-effect-free
  ``pptx.slide._EffectiveBackground`` view of the first ancestor
  carrying an explicit ``p:bg``. The proxy exposes ``.source``
  (``"slide"`` / ``"layout"`` / ``"master"``), ``.owner``,
  ``.bg_element``, and ``.fill`` (``None`` for a ``p:bgRef`` style
  reference). Reading through the proxy does not mutate the underlying
  XML, so inheritance stays intact.

- Add #582 ``shape.custom_props`` — dict-like mapping for attaching
  application-specific string metadata to any shape. Values persist in
  the shape's ``p:cNvPr/a:extLst`` under the URI
  ``{urn:loadfix-pptx:custom-props:v1}`` and round-trip through both
  python-pptx and PowerPoint. Available on every shape type (AutoShape,
  Picture, GraphicFrame, GroupShape, Connector); supports ``__getitem__``,
  ``__setitem__``, ``__delitem__``, ``__contains__``, ``get``, ``clear``,
  iteration, and insertion-order preservation. Issue #582.

- docs: #950 add ai-use-cases page. Issue #950
  (https://github.com/scanny/python-pptx/issues/950) asked whether
  python-pptx will "include Generative AI". The new
  ``docs/user/ai-use-cases.rst`` page explains that AI orchestration
  (prompts, model calls, auth, safety) is caller-side, sketches the
  patterns that compose well with python-pptx (LLM-authored JSON →
  deck, content translation, accessibility alt-text, data-to-slide
  pipelines, deck summarisation), and points at
  :ref:`rendering-to-pdf-video-or-image-formats` for the downstream
  ``.pptx`` → PDF/PNG/MP4 step.

- docs: #634 clarify ``MSO_THEME_COLOR_INDEX`` ↔ ``<a:schemeClr val>``
  mapping. Issue #634
  (https://github.com/scanny/python-pptx/issues/634) reported that the
  XML tag names (``tx1`` / ``bg1`` / ``dk1`` / ``lt1`` / etc.) don't
  visibly match the enum member names (``TEXT_1`` / ``BACKGROUND_1`` /
  ``DARK_1`` / ``LIGHT_1``). The mapping is intentional — OOXML carries
  two parallel sets of scheme-color names connected by the slide
  master's ``<p:clrMap>``: slide-level (``tx1`` / ``bg1`` / ``tx2`` /
  ``bg2``) and theme-level (``dk1`` / ``lt1`` / ``dk2`` / ``lt2``).
  ``MSO_THEME_COLOR_INDEX`` exposes a member for every value of OOXML
  ``ST_SchemeColorVal`` (except ``phClr``), correctly writing the
  slide-level tag PowerPoint itself emits when callers assign
  ``TEXT_1`` / ``BACKGROUND_1`` / ``TEXT_2`` / ``BACKGROUND_2`` and the
  theme-level tag when they assign ``DARK_1`` / ``LIGHT_1`` / ``DARK_2``
  / ``LIGHT_2``. The enum's class docstring, each ambiguous member's
  description, and ``docs/api/enum/MsoThemeColorIndex.rst`` now spell
  out the two sets plus the default ``<p:clrMap>`` that connects them.
  ``tests/test_issue_634_theme_color_mapping_verify.py`` pins the
  two-way mapping (enum → ``xml_value`` and ``from_xml`` → enum) plus
  round-trip behavior through ``font.color.theme_color`` for both the
  ``TEXT_1`` / ``tx1`` and ``DARK_1`` / ``dk1`` cases and confirms the
  MS-API integer values.

- verify: #316 picture placeholder renders in LibreOffice. Issue #316
  (https://github.com/scanny/python-pptx/issues/316) reported that a
  picture inserted into a layout's picture placeholder showed blank in
  LibreOffice Impress even though it rendered correctly in PowerPoint;
  the thread went cold without an opc-diff. The Wave 13
  ``_copy_inherited_spPr_decorations`` fix for #907 landed a ``p:pic``
  structure that carries every element LibreOffice's stricter reader
  requires (``p:nvPicPr`` with ``p:cNvPr`` / ``p:cNvPicPr`` (with
  ``a:picLocks``) / ``p:nvPr`` (with ``p:ph``), ``p:blipFill`` with
  ``a:blip@r:embed`` and ``a:stretch/a:fillRect``, and ``p:spPr``).
  A manual LibreOffice render (``libreoffice --headless --convert-to
  pdf``) confirms #316 is resolved on the current master.
  ``tests/test_issue_316_libreoffice_placeholder_pic.py`` pins the XML
  structure for three insertion variants — the specialised PICTURE
  placeholder, a generic OBJECT/content placeholder, and the
  ``crop=False`` fit path — both in memory and after
  ``Presentation.save`` + reopen, so a future regression of the
  placeholder-promotion XML would re-break LibreOffice compatibility
  visibly in CI.

- verify: #339 strikethrough regression test added. Issue #339
  (https://github.com/scanny/python-pptx/issues/339) asked for a public
  way to read and write the ``a:rPr/@strike`` attribute. The feature is
  shipped as :attr:`Font.strikethrough` (tri-state |True| / |False| /
  |None| plus the :class:`MSO_TEXT_STRIKE_TYPE` enum, see
  ``FEATURES.md``); ``tests/test_issue_339_strikethrough_verify.py``
  adds a ``DescribeIssue339Strikethrough`` suite that pins the
  inherit-by-default read, the ``True`` / ``False`` / enum-assignment
  setter paths (including ``DOUBLE_LINE``), ``None``-clears-attribute
  semantics, the exact ``@strike`` attribute spelling on ``a:rPr``, a
  :meth:`Presentation.save` + reopen round-trip across every flavour,
  and strikethrough preservation through
  :meth:`TextFrame.replace_text`.

- verify: #440 resolved by #828 (``Chart.series_in_rows``). Issue #440
  (https://github.com/scanny/python-pptx/issues/440) asked for a
  supported way to distinguish a chart whose source data is laid out
  with series **in columns** (PowerPoint's default) from one whose
  data has been switched to series **in rows** (via *Chart Design >
  Switch Row/Column* in the PowerPoint UI). OOXML does not persist a
  dedicated ``@switchRowCol`` attribute — the orientation is implicit
  in the shape of the ``c:f`` cell-range references under each
  ``c:ser``. The resolution shipped in Wave 11 as
  :attr:`.Chart.series_in_rows` (``feat(chart): #828 add
  Chart.series_in_rows read-only accessor``), which parses the first
  series's category and series-name references and returns ``False``
  (series-in-columns), ``True`` (series-in-rows), or ``None`` when
  the orientation can't be determined (no series, XY/scatter or
  bubble chart with no ``c:cat`` or parseable ``c:tx``, or inline
  literals). Read-only because PowerPoint rewrites every ``c:ser``
  sub-reference on the UI toggle and emulating that without a live
  workbook would silently diverge from what PowerPoint re-authors on
  next save — callers needing to flip orientation should re-author
  via :meth:`.Chart.replace_data` with the data transposed.
  ``tests/test_issue_440_data_alignment_verify.py`` pins the #440
  reporter's perspective: default-layout chart reads ``False``,
  post-switch XML reads ``True``, single-cell ``c:cat`` falls
  through to the ``c:tx`` row check, XY/scatter and bubble charts
  with no parseable refs read ``None`` (while freshly authored
  XY/bubble charts fall through to ``False`` via their row-1
  ``c:tx`` ref), empty chart reads ``None``, and the orientation
  signal survives a ``Presentation.save`` + reopen round-trip.
  Cross-references the fine-grained XML-shape unit tests in
  ``tests/chart/test_chart.py::DescribeChart``.

- verify: #452 resolved by #971 (``BaseShape.is_hidden``). Issue #452
  (https://github.com/scanny/python-pptx/issues/452) asked for a
  supported way to toggle a shape's visibility without deleting it —
  the PowerPoint "Selection Pane eye-icon" idiom. The #971 feature
  shipped ``BaseShape.is_hidden``, a read/write ``bool`` mapping the
  ``hidden`` attribute on the shape's ``cNvPr`` element, available on
  every shape kind (``p:sp``, ``p:pic``, ``p:cxnSp``, ``p:grpSp``,
  ``p:graphicFrame``). ``tests/test_issue_452_shape_hidden_verify.py``
  pins the #452 reporter's end-to-end workflow: default visible state,
  setter emits ``cNvPr/@hidden="1"``, unhide clears the attribute,
  ``Presentation.save`` + reopen round-trip, per-shape independence
  across sibling shapes, and coverage across all five ``nv*Pr``
  parent variants.

- verify: #622 per-slide audio narration recipe. Issue #622
  (https://github.com/scanny/python-pptx/issues/622) asked for an
  ``add_narration(...)`` wrapper that would embed an audio clip,
  auto-play it with the slide, hide the speaker icon, and advance the
  slide when playback ended. Every primitive is already public:
  ``SlideShapes.add_movie(mime_type="audio/*", autoplay=True)`` embeds
  the audio and wires the ``withPrevious`` start condition,
  ``BaseShape.is_hidden`` flips the ``cNvPr`` hidden attribute on the
  returned shape, and ``Slide.transition.advance_after_time`` carries
  the auto-advance delay (in milliseconds). No wrapper API is added;
  instead, a new "Recipe: per-slide audio narration" section in
  ``docs/user/media.rst`` documents the composition. A
  ``DescribeIssue622AudioNarrationRecipe`` suite
  (``tests/test_issue_622_audio_narration_verify.py``) pins each leg
  and a save + reopen round-trip of the full recipe, and a matching
  behave scenario in ``features/shp-shapes.feature`` runs the
  composition end-to-end so the three legs stay wired together.

- build: #327 document test deps (msoffcrypto-tool, pyparsing). Annotate
  ``requirements-test.txt`` so contributors can see at a glance which
  optional third-party package each line unlocks, and add
  ``@pytest.mark.skipif`` guards to ``tests/opc/test__crypto.py`` and
  ``tests/test_password.py`` so the suite no longer errors out with
  ``ModuleNotFoundError: No module named 'msoffcrypto'`` when the
  optional dependency is absent — previously 5 failures + 4 errors in
  the crypto test modules, now cleanly skipped.

- fix: #168 ``TextFrame.fit_text()`` surfaces opaque exceptions for several
  common edge cases (missing platform font directories, no system font
  matching ``font_family``/``bold``/``italic``, unreadable or non-TrueType
  ``font_file``, non-positive ``max_size``, margins exceeding shape
  dimensions). All of these now raise :class:`pptx.exc.TextLayoutError`
  with an actionable message pointing the caller at the supported
  workaround (supply an explicit ``font_file`` or set
  ``text_frame.auto_size``). The original exception is chained as
  ``__cause__`` for debuggability. Regression coverage in
  ``tests/test_issue_168_fit_text_exceptions.py``.

- Add #393 ``XySeries.x_values`` / ``BubbleSeries.bubble_sizes`` —
  read the cached X values of a scatter-series and the bubble-size
  values of a bubble series as tuples of floats (a ``None`` element
  for each blank cell), complementing the existing ``values``
  (Y-value) accessor. Streaming variants
  ``XySeries.iter_x_values()`` and ``BubbleSeries.iter_bubble_sizes()``
  are also provided. Resolves a long-standing ask to enumerate
  scatter-plot (x, y) coordinates from python-pptx without dropping
  into the ``c:xVal`` element tree by hand
  (https://github.com/scanny/python-pptx/issues/393).

- feat: #481 add :attr:`.CategoryAxis.label_align` — read/write
  :ref:`XlTickLabelAlignment` property mapping to
  ``c:catAx/c:lblAlgn/@val`` with values ``CENTER`` / ``LEFT`` / ``RIGHT``
  (``ctr`` / ``l`` / ``r``). Exposes PowerPoint's "Label alignment"
  drop-down on the *Format Axis* pane. Returns ``CENTER`` (the PowerPoint
  default) when the ``c:lblAlgn`` element is absent; assigning ``CENTER``
  removes the backing element so the XML stays minimal. Assigning a
  non-member raises :class:`ValueError`. Also adds the new
  ``pptx.enum.chart.XL_TICK_LABEL_ALIGNMENT`` enum.

- feat: #1043 ``_BaseSeries.delete()`` excludes a column from a chart.
  Issue #1043 (https://github.com/scanny/python-pptx/issues/1043) asked
  for a supported way to mark an embedded-workbook column as excluded
  from the chart — the programmatic equivalent of PowerPoint's
  right-click *Select Data* dialog unchecking a series. Series selection
  is modelled at the ``c:ser`` level (a column that should not be
  plotted simply has no corresponding ``c:ser``), so the fix is
  series-level deletion: :meth:`.Series.delete` detaches the underlying
  ``c:ser`` from its parent ``c:{x}Chart`` and leaves the embedded
  xlsx untouched. PowerPoint tolerates (and silently re-normalises) the
  resulting gaps in remaining series' ``c:idx`` / ``c:order``.
  ``tests/test_issue_1043_series_delete.py`` pins the contract from the
  reporter's perspective — 5-series chart, delete two non-contiguous
  series, remaining names/values intact across a
  :meth:`.Presentation.save` + reopen round-trip. Documented in the
  user guide under *Working with charts → Excluding a series from a
  chart*.

- Add #764 ``Chart.set_title(text)`` and ``Chart.set_axis_title(axis, text)``
  convenience helpers that collapse the common "ensure title present,
  rewrite text" idiom into a single chainable call. Passing ``None`` (or
  ``""``) removes the corresponding title. *axis* is one of
  ``"category"``, ``"value"``, or ``"secondary_value"``.

- Add #447 ``BaseShape.theme_style_refs`` for reading and writing a
  shape's theme-style preset. Returns a :class:`ThemeStyleRefs`
  named-tuple (``line_ref``, ``fill_ref``, ``effect_ref``, ``font_ref``)
  read from the shape's ``<p:style>`` child and the four ``<a:*Ref>``
  grandchildren; setter accepts the named-tuple (or a plain 4-tuple) to
  author a fresh ``<p:style>`` or |None| to drop it. Works on
  autoshapes, text-boxes, connectors, and pictures; graphic-frame and
  group-shape parents raise :class:`ValueError` because those element
  kinds do not carry a ``<p:style>``. See
  ``docs/user/autoshapes.rst`` for the end-to-end snippet and
  :class:`~pptx.shapes.base.ThemeStyleRefs` for the data-class
  signature.

- feat: #410 embedded 3D-model passthrough (detection + round-trip).
  PowerPoint 365 (Office 2016+) "Insert > 3D Models" authoring produces
  a ``p:graphicFrame`` whose ``a:graphicData/@uri`` is the Microsoft
  extension ``http://schemas.microsoft.com/office/drawing/2016/12/model3D``
  and whose single child is an ``am3d:model3D`` element pointing at an
  embedded ``.glb`` / ``.obj`` / ``.fbx`` part. This wave ships a
  read-only MVP: :attr:`GraphicFrame.has_model_3d` detects the frame,
  :attr:`GraphicFrame.model_3d_xml` returns the raw XML of the
  ``am3d:model3D`` element, and :attr:`GraphicFrame.model_3d` returns a
  ``Model3D`` proxy exposing :attr:`~pptx.shapes.model3d.Model3D.embedded_rel_id`,
  :attr:`~pptx.shapes.model3d.Model3D.ext`, and
  :attr:`~pptx.shapes.model3d.Model3D.media_blob`. The
  ``am3d:`` namespace (``.../2017/model3d``) and the
  :data:`~pptx.spec.GRAPHIC_DATA_URI_MODEL_3D` spec constant are
  registered so 3D-model-containing decks survive save+reopen verbatim.
  Authoring 3D models (camera, lighting, scene-graph, animation
  selection) is deliberately deferred; see
  ``docs/dev/analysis/model-3d.rst`` for the roadmap.

- verify: #419 ``FillFormat.blip_fill`` on shapes regression test.
  Issue #419 (https://github.com/scanny/python-pptx/issues/419) asked
  for a supported way to apply PowerPoint's "Picture or texture fill"
  to an auto-shape. The feature is shipped as
  ``FillFormat.blip_fill(image_file)`` (see ``FEATURES.md``);
  ``tests/test_issue_419_shape_picture_fill_verify.py`` adds a
  breadth-first ``DescribeIssue419ShapePictureFill`` suite that pins
  the reporter's workflow end-to-end: apply a picture fill to a
  rectangle and to an oval, confirm ``shape.fill.type`` is
  ``MSO_FILL.PICTURE`` both in-memory and after a ``Presentation.save``
  + reopen round-trip, accept a ``BytesIO`` stream and a
  ``pathlib.Path`` (via ``str(path)``) as the image file, and assert
  the ``<a:blipFill>`` structure directly under the shape's
  ``<p:sp>/<p:spPr>``. Complements the existing #234 suite, which
  covers image-part reuse and the no-part error path.

- verify: #586 extended_properties company/manager coverage. Pins
  :attr:`Presentation.extended_properties` against silent breakage of
  the #586 ask — reading and writing ``company`` and ``manager`` under
  ``/docProps/app.xml`` — with a round-trip through
  :meth:`Presentation.save`, confirmation that the neighbouring
  ``application`` / ``app_version`` / ``slide_count`` fields remain
  accessible, and cross-part isolation from
  :attr:`Presentation.core_properties`. See
  ``tests/test_issue_586_company_manager_verify.py``.

- verify: #768 resolved by #720 (name_ea / name_cs slots). Issue #768
  ("Change font name not working for asian characters") is marked upstream
  as a duplicate of #720 — both hit the OOXML font-slot dispatch rule
  (``a:rPr`` carries independent ``a:latin`` / ``a:ea`` / ``a:cs`` children
  and PowerPoint routes each Unicode script to the matching slot).
  ``Font.name_ea`` and ``Font.name_cs``, shipped by Wave 1 #337, make the
  East-Asian and complex-script slots addressable from Python, which is
  the fix the #768 reporter needed. ``tests/test_issue_768_asian_font_verify.py``
  pins the three-slot contract from the #768 perspective: single-slot
  Latin, single-slot EA, single-slot CS, all-three-together, a
  ``Presentation.save`` + reopen round trip, and a public-surface
  cross-reference to #720.

- verify: #821 resolved by Font.use_theme_hyperlink_color + run-level color
  override. The #821 thread ("How to change hyperlink text color") asked for
  a first-class surface to recolor hyperlinked text from python-pptx — an
  ask that was previously frustrated by PowerPoint's theme ``a:hlink``
  entry overriding any ``a:rPr/a:solidFill`` on a run carrying an
  ``a:hlinkClick``. The #940 feature shipped
  :attr:`.Font.use_theme_hyperlink_color` — a tri-state property that
  writes a python-pptx-owned ``a:extLst/a:ext`` marker under the run's
  ``a:hlinkClick`` to record the caller's intent that the run's explicit
  colour be preferred over the theme. Combined with
  ``run.font.color.rgb = ...`` the caller now has the full "red hyperlink"
  recipe. ``tests/test_issue_821_hyperlink_text_color_verify.py`` pins
  the canonical recipe, the default (theme-applies) read path, the
  toggle-back-and-forth idempotency, ``Presentation.save`` + reopen
  round-trip, underline preservation across the colour change, and
  independence between hyperlinked and plain sibling runs in the same
  paragraph.

- verify: #756 legacy comments regression test added. Issue #756
  (https://github.com/scanny/python-pptx/issues/756) asked for read/write
  access to PowerPoint's review-comment feature — the
  *Insert > Comment* yellow-sticky-note annotations attached to a slide.
  The feature is resolved by the #487 legacy-comments work which shipped
  ``Slide.comments`` / ``Slide.has_comments`` over the ECMA-376 Part 1
  ``p:cmLst`` + ``p:cmAuthorLst`` schema with part topology per §13.3.3
  (author registry on the presentation part, per-slide comments parts).
  ``tests/test_issue_756_comments_verify.py`` adds a
  ``DescribeIssue756Comments`` suite pinning the reporter-facing
  scenarios: authoring a new comment with text / author / position /
  timestamp, reading back through the ``Comments`` iterator, ordering
  across multiple comments, author-registry deduplication across
  ``get_or_add`` calls, removing a comment via the underlying
  ``p:cmLst`` helper, EMU round-trip on ``Comment.position``, full
  ``Presentation.save`` + reopen round-trip of every comment attribute
  and the package-level author registry, and the
  ``presentation.slides[0].comments`` public API navigation path.

- verify: #765 resolved by #938 + #378 (effective_* chain). Issue #765
  (https://github.com/scanny/python-pptx/issues/765) collected the
  long-running "effective font" discussion: a request for read-side
  API that returns the size / color / bold / italic / typeface
  PowerPoint would actually render for a run, walking the full OOXML
  inheritance chain (run ``a:rPr`` → paragraph ``a:pPr/a:defRPr`` →
  text-body ``a:lstStyle`` → slide-master ``p:txStyles`` →
  presentation ``p:defaultTextStyle``) and resolving theme/scheme
  colors against the master's theme. The fork closed that ask in two
  shipments: Wave 6 #938 added :attr:`.Font.effective_color`, and
  Wave 12 #378 added its size / bold / italic / name siblings
  (:attr:`.Font.effective_size`, :attr:`.Font.effective_bold`,
  :attr:`.Font.effective_italic`, :attr:`.Font.effective_name`).
  Together the five properties cover the complete "effective font"
  surface the #765 thread asked for. Adds a verify-close suite
  ``DescribeIssue765EffectiveFont`` under
  ``tests/test_issue_765_effective_font_verify.py`` that pins the
  placeholder → master size walk, the theme-scheme color resolution,
  paragraph-level bold/italic inheritance, the minor-Latin typeface
  fall-through, explicit-run-overrides-inherited semantics, the
  all-|None| "nothing declared" case, a save + reopen round-trip of
  every effective_* on an inherited-formatting placeholder, and a
  bare-textbox chain walk to the master's ``bodyStyle``.

- verify: #1063 resolved by the #378/#938/#765 effective_* chain. Issue
  #1063 (https://github.com/scanny/python-pptx/issues/1063) framed the
  classic "walk a deck, read the font properties PowerPoint would
  render" workflow: the reporter wanted ``run.font.color`` /
  ``run.font.size`` / ``run.font.name`` on an arbitrary run and got
  |None| (or a raised ``AttributeError`` on ``color.rgb``) whenever the
  value came from the paragraph's ``a:pPr/a:defRPr``, the shape's
  ``a:lstStyle``, or the slide-master's ``p:txStyles`` fallback — the
  default for every placeholder-authored deck. The fix is the read-side
  :attr:`.Font.effective_color` (#938, Wave 6) plus
  :attr:`.Font.effective_size`, :attr:`.Font.effective_bold`,
  :attr:`.Font.effective_italic`, and :attr:`.Font.effective_name`
  (#378, Wave 12), which together walk the full inheritance chain (run
  ``a:rPr`` → paragraph ``a:pPr/a:defRPr`` → text body
  ``a:lstStyle/a:lvl{N}pPr/a:defRPr`` → master ``p:txStyles`` →
  presentation ``p:defaultTextStyle``) and return the value PowerPoint
  would render, with theme/scheme colours resolved against the master's
  theme and ``a:lumMod``/``a:lumOff`` tint/shade applied. #1063 overlaps
  #378 (same four non-colour properties) and #765 ("read after
  loading"), and the read-side mechanism is the same as #938's colour
  fix. Adds a regression suite
  ``DescribeIssue1063EffectiveStyle`` under
  ``tests/test_issue_1063_effective_style_verify.py`` that pins the
  #1063 reporter's end-to-end workflow — walking a multi-slide deck and
  reading effective colour/size/name/bold/italic on every run (title
  placeholders, body placeholders, textboxes, and table cells), with
  explicit-wins-over-inherited sanity checks, theme-coloured run
  resolution (on both direct textbox runs and inherited placeholder
  runs), and save+reopen round-trips of placeholder, table-cell, and
  theme-coloured scenarios to cover #765's "read after loading" framing.

- fix: #740 ``_Paragraph.font.size`` (and every other paragraph-level default
  run-property assignment) no longer lands in the wrong position when the
  paragraph already contains an ``a:br`` line-break. The ``a:pPr`` insertion
  walked the declared successor list tag-by-tag, returning the first
  successor type that was present rather than the earliest-positioned
  child. When ``paragraph.text = "\nHello"`` ran before
  ``paragraph.font.size = Pt(24)``, the paragraph already contained
  ``[a:br, a:r]``; ``a:pPr`` would then be inserted *after* ``a:br`` (the
  first ``a:r`` being found before the first ``a:br`` in the successor
  tuple), producing ``<a:p><a:br/><a:pPr/>…</a:p>``. PowerPoint silently
  discards the out-of-order ``a:pPr`` on load, so the caller's font size
  vanished. ``BaseOxmlElement.insert_element_before`` now picks the
  earliest existing child whose tag is in the successor set, preserving
  schema order regardless of which successor types are present.

- fix: #789 ``DataLabels.position`` now validates against a per-chart-type
  whitelist and raises ``ValueError`` with a clear message when an
  incompatible position is assigned, instead of silently producing a file
  PowerPoint refuses to open. The classic trigger is
  ``plots[0].data_labels.position = XL_LABEL_POSITION.OUTSIDE_END`` on a
  stacked column / bar chart. The whitelist follows PowerPoint's UI: line
  / scatter accept ``CENTER`` / ``LEFT`` / ``RIGHT`` / ``ABOVE`` /
  ``BELOW``; clustered bar / column additionally accept ``OUTSIDE_END``;
  stacked and 100%-stacked bar / column accept only ``CENTER`` /
  ``INSIDE_END`` / ``INSIDE_BASE``; pie / doughnut accept ``CENTER`` /
  ``INSIDE_END`` / ``OUTSIDE_END`` / ``BEST_FIT``; area accepts
  ``CENTER``. Validation applies at both the plot-level
  (``plot.data_labels``) and series-level (``series.data_labels``) entry
  points. Assigning ``None`` remains universally legal and clears the
  ``c:dLblPos`` element; a rejected assignment leaves the XML untouched.

- fix: #907 ``PicturePlaceholder.insert_picture`` (and the same method on
  the generic ``SlidePlaceholder``) now preserves styling decorations that
  the slide layout's picture placeholder defines on its ``p:spPr`` — in
  particular ``a:prstGeom`` (including preset-geometry adjustments),
  ``a:ln`` outlines, and ``a:effectLst`` entries like soft edges. Prior
  behavior promoted the placeholder ``p:sp`` to a ``p:pic`` with an empty
  ``p:spPr``, so any outline / soft-edge / clipped-geometry styling the
  layout designer intended for inserted pictures silently disappeared.
  Direct decorations on the slide-level placeholder win over inherited
  ones, and fill-related children (``a:blipFill`` / ``a:solidFill`` / ...)
  are intentionally not copied — the picture supplies its own fill.

- fix: #947 ``TextFrame.text`` and ``_Paragraph.text`` now surface the
  plain-text rendering carried by an ``mc:Fallback`` subtree when a
  paragraph contains an ``mc:AlternateContent`` wrapper (typically an
  inline OMML math equation PowerPoint auto-inserts when it recognises
  mathematical syntax such as ``m^3``). Previously the walker only
  iterated ``a:r`` / ``a:br`` / ``a:fld`` direct children of ``a:p``, so
  the fallback text was invisible to text-extraction callers and
  ``.text`` returned an empty string for equation-only paragraphs. The
  live ``mc:Choice`` side is intentionally skipped to avoid
  double-counting the OMML characters already mirrored under
  ``mc:Fallback``.

- Add #321 ``Pie3DPlot`` support — charts of type
  ``XL_CHART_TYPE.THREE_D_PIE`` / ``THREE_D_PIE_EXPLODED`` (backing element
  ``c:pie3DChart``) now deserialize into a new
  :class:`pptx.chart.plot.Pie3DPlot` instance rather than raising
  ``ValueError("unsupported plot type ... pie3DChart")`` from
  ``PlotFactory``. The class mirrors |PiePlot|: it inherits the full
  ``_BasePlot`` API (series, categories, data labels, vary-by-categories)
  and ``Chart.chart_type`` reports ``THREE_D_PIE`` /
  ``THREE_D_PIE_EXPLODED`` based on ``c:ser/c:explosion``. This unblocks
  ``Chart.replace_data`` and ``replace_data_preserve_formulas`` on
  3D-pie charts. A new ``CT_Pie3DChart`` oxml class follows the
  ECMA-376 ``EG_PieChartShared`` tag sequence (``c:varyColors`` /
  ``c:ser`` / ``c:dLbls`` / ``c:extLst``).

- feat: #373 add ``Chart.has_data_table`` / ``Chart.data_table`` for the
  chart data-table (``c:plotArea/c:dTable``) rendered beneath the plot
  area. Assigning ``True`` to ``has_data_table`` writes a default
  ``c:dTable`` with all four show-\* flags on (horizontal and vertical
  borders, outline, legend keys); assigning ``False`` removes it. The
  new ``_DataTable`` proxy exposes read/write
  ``show_horz_border`` / ``show_vert_border`` / ``show_outline`` /
  ``show_keys`` booleans and a ``format`` |ChartFormat| for fill / line /
  effect styling of the data-table itself. Implemented against the
  ``CT_DTable`` grammar in ``dml-chart.xsd``.

- feat: #532 add ``Slide.shape_tree_flat`` and ``SlideShapes.descendants()``
  — a Selection-Pane-equivalent flat iterator that yields every shape on a
  slide, including the children of any :class:`.GroupShape` at every
  nesting depth, in document (z-order) sequence. Group shapes themselves
  are yielded before their contents so callers see the container
  alongside its members, matching PowerPoint's Selection Pane listing.
  :meth:`SlideShapes.get_by_name` and :meth:`SlideShapes.find_all_by_name`
  gained an ``include_descendants=False`` keyword that flips their lookup
  to the same flat traversal for finding shapes named inside a group.

- feat: #695 add ``_BaseSeries.source_range``, ``_BaseSeries.category_range``,
  ``_BaseSeries.name_range``, and ``_BaseSeries.values_sheet_reference`` —
  read access to the ``c:f`` formula on ``c:val/c:numRef`` (values),
  ``c:cat/c:strRef`` or ``c:cat/c:numRef`` (categories), and
  ``c:tx/c:strRef`` (series name). ``source_range`` has a setter that
  rewrites the formula text in place (creating a ``c:f`` child when
  missing); the cached values under ``c:numCache`` are left untouched —
  call ``Chart.update_cached_values()`` afterward to reconcile. A new
  ``SheetReference(sheet_name, a1_range)`` namedtuple is returned from
  ``values_sheet_reference`` with quoted sheet names unquoted
  (``'My Sheet'!$A$1`` → ``sheet_name="My Sheet"``).

- feat: #413 Add :meth:`.SlideMaster.add_layout` — create a brand-new
  slide layout on an existing master. This is the same-master cousin of
  :meth:`.SlideMaster.add_layout_from` (shipped for #1028 to *import*
  a layout from a different master) and mirrors the *Insert Layout*
  action exposed by PowerPoint's *Slide Master* view. The caller supplies
  a ``name`` (required, unique within the master) and optionally a
  ``based_on`` layout on *this* master to seed the new layout from.
  With ``based_on=None`` the new layout is built from a minimal blank
  template (``@type="cust"``, no placeholders, a
  ``p:clrMapOvr/a:masterClrMapping`` so the color map inherits from the
  master). A fresh ``p:sldLayoutId`` entry is appended to the master's
  ``p:sldLayoutIdLst`` with a newly-allocated id. Raises
  :class:`ValueError` on a name collision, on an empty name, or when
  ``based_on`` belongs to a different master (use
  :meth:`.SlideMaster.add_layout_from` for cross-master import). The
  supporting ``SlideLayoutPart.new_blank`` class method and
  ``CT_SlideLayout.new_blank`` oxml factory are reusable building blocks
  for code that needs a minimal custom-layout XML scaffold.

- feat: #1126 add ``Slide.background.bg_element`` — a side-effect-free
  read-only property that returns the underlying ``p:bg`` element, or
  |None| when the slide's background is inherited from its master or
  layout. The existing ``_Background._element`` still returns ``p:cSld``
  for backwards compatibility. Also adds
  ``Slide.copy_background_from(source_slide)``, which deep-copies the
  source slide's ``p:bg`` subtree onto this slide (or strips the explicit
  background and restores inheritance when the source inherits). Makes
  raw-XML background copy workflows a one-liner.

- feat: add ``Slide.follow_master_background()`` method to revert a custom
  slide background to master inheritance (closes gap noted during Wave 12 #366
  verify). The existing ``Slide.follow_master_background`` attribute is now a
  dual bool-like/callable proxy: reading returns ``True``/``False`` for the
  inheritance state (unchanged), while calling it (``slide.follow_master_background()``)
  drops the slide's ``p:bg`` child — PowerPoint's *Reset Background* button
  equivalent — and returns the slide for chaining.

- docs: #655 add a "Numbered lists" recipe to ``docs/user/text.rst``
  documenting the loop-over-``text_frame.paragraphs`` idiom for turning a
  text frame into a numbered list via the existing
  ``paragraph.bullet.auto_number(scheme, start_at=None)`` API (shipped under
  #100). The recipe covers the happy path, skipping a heading paragraph,
  starting mid-sequence with ``start_at``, and interleaving numbered and
  un-numbered paragraphs. No API change — the building block already
  existed; the gap was documentation. A regression test under
  ``tests/test_issue_655_numbered_lists_verify.py`` asserts the recipe
  round-trips through ``Presentation.save`` and reads back the expected
  ``a:buAutoNum`` elements on each paragraph.
- feat: #378 add ``Font.effective_size``, ``Font.effective_bold``,
  ``Font.effective_italic``, and ``Font.effective_name`` — read-only
  properties that walk the inheritance chain (run ``a:rPr`` → paragraph
  ``a:defRPr`` → text body ``a:lstStyle`` → slide-master ``p:txStyles`` →
  presentation ``p:defaultTextStyle``) and return the value PowerPoint
  would actually render, rather than |None| when the run itself declares no
  explicit value. Mirrors the existing ``Font.effective_color``.

- feat: #194 ``Slides.add_slide(slide_layout, index=None)`` now accepts an
  optional ``index`` keyword so callers can insert a new slide at a
  specific zero-based position rather than only appending. Index semantics
  match :meth:`Slides.move_slide` and :meth:`Slides.duplicate`: ``None``
  appends (the historical behavior), a negative index counts from the end,
  and an index beyond the end is clamped to the last position.

- verify: #657 resolved by prior line-end arrow work. Issue #657 asked to
  expand ``MSO_CONNECTOR_TYPE`` with arrow-head variants (*Connector with
  Arrow*, *Double Arrow*, *Curved with Arrow*, *Elbow with Arrow*) that
  PowerPoint's Insert Shapes gallery lists as separate tiles. In OOXML
  those are not distinct connector geometries — they are an existing
  ``MSO_CONNECTOR`` preset (``STRAIGHT`` / ``ELBOW`` / ``CURVE``) whose
  ``a:ln`` carries an ``a:headEnd`` / ``a:tailEnd`` decoration. That
  machinery is already exposed for fork-era releases via
  ``LineFormat.begin_arrow`` / ``LineFormat.end_arrow`` (with ``.type``,
  ``.width``, ``.length`` sub-properties), reachable from a connector as
  ``connector.line.begin_arrow`` / ``connector.line.end_arrow``. Adding
  new ``MSO_CONNECTOR_TYPE`` members would have been a spec-breaking
  change — those three are the only geometries defined by ISO/IEC 29500.
  Ships a verify-close unit test on ``Connector`` and a
  ``shp-connector.feature`` scenario that exercises arrow configuration
  on a real connector, plus a documentation snippet in
  ``docs/user/autoshapes.rst`` / a note in ``FEATURES.md`` mapping the
  PowerPoint gallery tile names onto the real API.

- verify: #116 ``Picture.replace_image(...)`` regression test added.
  Issue #116 (https://github.com/scanny/python-pptx/issues/116) — the
  long-running "feature: Picture.replace_image()" request that drew 23
  comments — is resolved by the supported API shipped in this release
  line. ``tests/test_issue_116_replace_image_verify.py`` adds a
  breadth-first ``DescribeIssue116ReplaceImage`` suite that pins every
  preservation guarantee named in the method docstring (position, size,
  rotation, all four crop fractions, masking auto-shape, outline color
  and width, ``cNvPr/@name`` / ``@descr`` / ``@title``) across a
  ``Presentation.save`` + reopen round-trip, plus both file-path and
  ``BytesIO`` input forms and the ``ValueError`` branch for a picture
  with no embedded image. Complements the existing #834 (geometry +
  crop) and #819 (find-by-alt-text workflow) regression suites.

- verify: #470 combo charts shipped by ``feat: #338 combo charts``
  (``Chart.add_plot(chart_type, chart_data)``) together with
  ``feat: #141 secondary value axis read access``
  (``Chart.has_secondary_value_axis`` / ``Chart.secondary_value_axis``).
  The #470 reporter hit ``AttributeError: 'CategoryWorkbookWriter'
  object has no attribute 'x_values_ref'`` while trying to follow a
  pre-fork recipe for overlaying a line plot on a bar chart. On this
  fork the supported path is ``chart.add_plot(XL_CHART_TYPE.LINE,
  line_data)`` after creating the first (bar/column) plot — the
  overlay plot reuses the existing plot's ``c:axId`` pair, and its
  series data is emitted inline as ``c:numLit`` / ``c:strLit`` so the
  embedded-workbook writer (the source of the reporter's crash) is
  not in the call path. Adds a regression suite
  ``DescribeIssue470ComboCharts`` under
  ``tests/test_issue_470_combo_charts_verify.py`` pinning
  (a) bar + line combo with round-trip save/reopen asserting both
  ``BarPlot`` and ``LinePlot`` survive, (b) the optional secondary
  value axis case via ``Chart.secondary_value_axis`` (authoring the
  secondary ``c:valAx`` from scratch on a library-created chart
  remains a follow-up — the round-trip and read/write path on a
  chart that already carries one is what's covered), and
  (c) custom line color + ``XL_MARKER_STYLE`` marker style on a
  series created by ``add_plot`` surviving the round-trip.
  Follow-up: convenience "add secondary value axis" /
  ``add_plot(..., secondary=True)`` authoring shortcut remains
  deferred; see ``docs/dev/analysis/combo-chart.rst``.
- verify: #366 (set a slide's background — solid / gradient / picture
  — and revert to master-inherited background) resolved by the existing
  :attr:`.Slide.background` / :attr:`.Slide.follow_master_background`
  surface. The #366 reporter asked for a Python-side hook to author a
  per-slide background; the ``_Background`` proxy's ``.fill`` property
  already exposes the standard :class:`.FillFormat` API — ``solid()``,
  ``gradient()``, and ``blip_fill(image_file)`` — so a solid-color,
  gradient, or picture background is authored with the same
  ``fill.fore_color.rgb = RGBColor(...)`` / ``fill.gradient()`` /
  ``fill.blip_fill(...)`` idioms used for shapes and table cells.
  :attr:`.Slide.follow_master_background` is ``True`` when the slide
  has no ``p:bg`` override and ``False`` once a custom background
  is applied; removing the ``p:bg`` child restores master inheritance.
  Adds ``tests/test_issue_366_slide_background_verify.py`` with an
  end-to-end verify suite covering the default master-inherited state,
  solid / gradient / picture authoring with ``Presentation.save`` +
  reopen round-trips, and the revert-to-inherited transition.
- verify: #620 (dynamically replicating a slide) resolved by Wave 1 #132.
  Issue #620 (https://github.com/scanny/python-pptx/issues/620) asked for
  a way to duplicate an existing slide — one that carries text
  placeholders, a picture, a chart, and a table — and then mutate content
  on the copy without disturbing the source. Wave 1 #132 shipped
  :meth:`Slides.duplicate`, which returns a newly-added slide with a
  deep-copied shape tree, shares image/chart/OLE/media parts with the
  source by relationship reuse (no content is re-embedded), and allocates
  a fresh slide-id so the copy is fully independent. Adds a regression
  suite ``DescribeIssue620DuplicateAndEdit`` under
  ``tests/test_issue_620_duplicate_and_edit_verify.py`` that pins the
  reporter's workflow end-to-end: duplicate a rich slide carrying all
  four content kinds, edit the duplicate's title and confirm the source
  title is unchanged, round-trip the pair through ``Presentation.save`` +
  reopen, and duplicate the same source twice in a row without mutating
  its shape tree.
- verify: #941 (unable to open generated presentation) closed as
  non-reproducible. The #941 reporter copied the Hello-World quickstart
  snippet from the manual verbatim and reported that the resulting
  ``.pptx`` was rejected as corrupt/invalid by Microsoft PowerPoint,
  PowerPoint-online, and Google Slides; a follow-up comment narrowed
  the trigger to "any textbox being present". A second contributor
  (MartinPacker) ran the exact same snippet on macOS and PowerPoint
  opened the output cleanly, and no traceback, file, or environment
  detail was ever posted by the reporter. The symptom is a hallmark of
  environment-local failure — a corrupted or partial wheel install, a
  Windows-editor CRLF/EOL mangling of the uploaded ``.pptx``, a zip
  re-compressor in the transfer chain, or antivirus quarantining the
  file — and no bug in ``pptx`` matching the claim has ever been
  identified. Adds a defensive regression suite
  ``DescribeIssue941QuickstartOpensCleanly`` under
  ``tests/test_issue_941_verify.py`` that pins both the verbatim
  quickstart reproducer and the "add a textbox" follow-up as producing
  well-formed ``.pptx`` packages (valid zip container, every member's
  CRC verifies, all required OPC parts present) and surviving
  ``Presentation.save`` + reopen. Any future regression at the zip,
  content-types, relationships, or default-template layer that would
  break the manual's first code example for real will reproduce the
  #941 symptom and be caught here.
- verify: #956 (deleting slides leaves unused parts behind) — the reporter
  worried that :meth:`.Slides.delete` would leave orphan slide / image /
  chart / notes-slide parts inside the saved ``.pptx`` zip, letting a
  heavily-edited deck grow forever. It doesn't: package save is already
  a reachability-based GC. :meth:`.OpcPackage.save` passes
  ``tuple(self.iter_parts())`` to :class:`.PackageWriter`, and
  ``iter_parts()`` walks the rels graph depth-first from the package
  root — so any part no longer reachable from the presentation part is
  silently dropped on save. :meth:`.Slides.delete` correctly severs
  reachability by removing the ``p:sldId`` entry and then calling
  :meth:`.XmlPart.drop_rel` on the presentation part, which in turn
  releases the slide's notes slide, its uniquely-referenced image /
  chart / embedded-xlsx / media / tags parts, etc. Parts shared with
  another surviving slide (e.g. an image on the duplicate of a deleted
  slide) are correctly preserved. Adds verify suite
  ``DescribeIssue956DeleteSlideCleanup`` under
  ``tests/test_issue_956_delete_slide_cleanup.py`` pinning the
  reporter's concern across five axes: unique-image GC, shared-image
  preservation, slide-part GC, notes-slide GC, chart + embedded-xlsx
  GC, save-shrinks-the-zip byte-size check, and idempotency across
  many repeated add/delete cycles.
- fix: #1058 connector coordinates accepting ``float`` values no longer
  produce a corrupt ``.pptx`` PowerPoint refuses to open. Floats passed
  to ``slide.shapes.add_connector(...)`` and to the
  ``Connector.begin_x`` / ``begin_y`` / ``end_x`` / ``end_y`` setters
  are now coerced with ``int(round(...))`` before they reach the
  int-typed ``a:off/@x``, ``a:off/@y``, ``a:ext/@cx`` and ``a:ext/@cy``
  XML attributes, matching the treatment already applied by
  ``CT_GraphicalObjectFrame.new_table_graphicFrame``.
- fix: #1111 ``Font.color`` getter no longer mutates the run's XML. Previously,
  reading ``run.font.color`` (or any attribute on it — ``.type``, ``.rgb``,
  ``str()``, etc.) eagerly inserted an empty ``<a:solidFill/>`` child into the
  run's ``<a:rPr>``, which PowerPoint interprets as "override inherited color
  with nothing" and permanently breaks the theme-color inheritance chain. The
  getter now returns a read-only proxy that reports ``.type is None`` when no
  explicit color is set; ``<a:solidFill>`` is created only when the caller
  writes ``color.rgb = ...`` or ``color.theme_color = ...``.
- fix: #650 per-point ``DataLabel.font.color.rgb`` (or any customization that
  auto-creates a ``c:dLbl`` via the per-point factory) no longer clobbers
  categories, values, and other ``c:show*`` settings inherited from the
  series-level ``c:dLbls``. ``CT_DLbl.new_dLbl`` used to hard-code
  ``c:showCatName val="0"`` (and the other show-flag children) on every new
  per-point ``c:dLbl`` container; children under ``c:dLbl`` *override* the
  series-level defaults, so author-visible categories would vanish the moment
  a color override touched a single point. The factory now omits the show-flag
  children, letting series-/plot-level settings inherit as intended. The
  series-level ``CT_DLbls.new_dLbls`` factory is unchanged (its show-flag
  defaults are the meaningful fallback when no series-level toggles exist).
- fix: #450 preserve series-level shadow (and other non-fill visual
  overrides) when assigning a color to a single ``CategoryPoint`` on a
  chart series. PowerPoint treats a per-point ``c:spPr`` as a complete
  replacement for the series-level shape properties, so any
  ``a:effectLst`` / ``a:ln`` / ``a:scene3d`` / ``a:sp3d`` authored on
  the series was silently dropped from that point's rendering the moment
  the user overrode its fill. ``CT_DPt._new_spPr`` now seeds a
  newly-created point ``c:spPr`` with deep copies of those non-fill
  children from the series ``c:spPr``, so overriding a point's color
  keeps the series shadow (and outline) on that point. Adds unit
  coverage under :class:`DescribePoint` and an end-to-end regression
  suite ``tests/test_issue_450_point_shadow_preservation.py`` that
  round-trips an authored shadow through save-and-reopen.
- Add #86 ``Table.add_row()`` and ``Table.add_column()`` one-liner convenience
  methods that delegate to the existing ``_RowCollection.add`` /
  ``_ColumnCollection.add`` (same defaults, same semantics). Saves a hop
  through the collection property for the common "append a row/column"
  case.
- feat: #269 Add :meth:`.SlideMaster.get_layout` and
  :meth:`.SlideLayouts.get_by_id` for layout lookup by the
  presentation-stable ``p:sldLayoutId/@id``, together with the read-only
  :attr:`.SlideLayout.slide_layout_id` property that exposes it. The id is
  preserved across layout reordering, so code that captures a layout
  reference by id in one run continues to resolve the same layout in a
  later run regardless of position changes — a more robust alternative
  to ``slide_master.slide_layouts[index]``.
- Add #663 :attr:`.BaseShape.text_frame_rect` — read-only property
  returning a :class:`~pptx.text.text.TextFrameRect` namedtuple
  ``(left, top, width, height)`` of |Length| values in EMU giving the
  slide-relative rectangle PowerPoint allocates for rendering a shape's
  text, i.e. the shape's bounding box shrunk by its four
  :attr:`.TextFrame.margin_left` / ``margin_top`` / ``margin_right`` /
  ``margin_bottom`` insets. Raises :class:`ValueError` on a shape that
  has no text frame. Useful for picking a font-size that will not
  overflow a shape at render time. See the user-guide section
  "Inspecting the text-rendering rectangle" in ``docs/user/text.rst``.
- feat: #1077 Add :attr:`._Hyperlink.target_slide` so a run of text can jump
  to another slide in the same presentation (the run-level equivalent of
  :attr:`ActionSetting.target_slide`). Assigning a |Slide| writes an
  ``a:hlinkClick`` with ``action="ppaction://hlinksldjump"`` and a slide
  relationship; assigning |None| or using ``del`` removes any hyperlink on
  the run. :attr:`._Hyperlink.address` now returns |None| for a slide-jump
  hyperlink so the URL and slide-jump surfaces remain strictly separate.
- Add #1028 :meth:`.SlideMaster.add_layout_from` — clone a slide layout
  from any other master (in this presentation or another) into this
  master. Addresses the "apply a layout from one master in another
  master" request from the #1028 reporter. The cloned layout is
  appended to the destination master's ``slide_layouts`` with a fresh
  ``p:sldLayoutId``; image relationships and external hyperlinks are
  materialised in the target package (images are content-deduplicated
  against existing parts). The clone inherits its theme (color / font
  / effect scheme) from the *destination* master, matching PowerPoint's
  behavior when a layout is dragged between masters in Slide Master
  view. Raises :class:`ValueError` when the destination master already
  contains a layout with the same name; rename one of them first. The
  accompanying oxml helpers include an ``add_sldLayoutId`` method on
  :class:`.CT_SlideLayoutIdList` that auto-allocates a fresh
  ST_SlideLayoutId (minimum 2147483648) for each new entry, and a
  :meth:`SlideLayoutPart.clone_from` classmethod that composes with
  the F1 relationship cloner to move images and other layout-owned
  parts across packages.

- docs: #960 add a "Check placeholder state before inserting a picture"
  recipe to ``docs/user/placeholders-using.rst`` showing how to use
  ``placeholder.placeholder_format.type`` (``PP_PLACEHOLDER.PICTURE`` /
  ``OBJECT``) together with ``placeholder.shape_type`` (flips to
  ``MSO_SHAPE_TYPE.PICTURE`` once populated) to guard ``insert_picture``
  against mis-typed placeholders and against re-populating a slot that
  already contains a picture. No API change — the accessors already exist;
  the gap was documentation. The same pattern generalizes to
  ``insert_table`` / ``insert_chart``.

- verify: #607 resolved by Wave 2 #544. Issue #607
  (https://github.com/scanny/python-pptx/issues/607) asked how to add
  vertical (Y-direction) error bars to the data points of an XY scatter
  chart — the reporter could see ``c:errBars`` in the schema but there
  was no Python-API hook to configure it on a series. Wave 2 #544
  (commit ``05072c75``, "feat(chart): #544 add error-bar support for
  chart series") shipped the missing API on :class:`_BaseSeries`:
  ``has_error_bars``, read/write ``error_bars`` returning an
  |ErrorBars| proxy (assign ``None`` to remove), and
  ``set_error_bars(type_, value, include, direction)`` which attaches a
  fresh ``c:errBars`` block. The new :class:`XL_ERROR_BAR_DIRECTION`
  enum exposes ``X`` and ``Y`` members so clients of XY-scatter and
  bubble series — which have two value axes rather than a category axis
  plus a value axis — can explicitly select the vertical whiskers shown
  in the #607 screenshot. Adds an end-to-end verify suite
  ``DescribeIssue607XyErrorBarsVerify`` under
  ``tests/test_issue_607_xy_error_bars_verify.py`` that pins the
  scenario from the user's perspective across all five XY-scatter
  variants plus bubble, including a ``Presentation.save`` + reopen
  round-trip to confirm the ``c:errBars`` subtree survives
  serialization.

- verify: #884 (can't change text without changing format + ``AttributeError``
  on inherited color) resolved by the #836 / #420 / #308 / #938 composition.
  The #884 reporter hit two coupled failures while trying to rewrite a
  templated placeholder: ``text_frame.text = "new"`` clobbered every
  run-level style, and ``font.color.rgb`` raised ``AttributeError: no .rgb
  property on color type '_NoneColor'`` when they tried to read the old
  color to re-apply it after the rewrite. Wave 6 #836 added
  :meth:`.TextFrame.replace_text` / :meth:`._Paragraph.replace_text`,
  which keep the origin run's full ``a:rPr`` while absorbing the
  replacement text — bold, italic, underline, font size, font name and
  color all survive a rewrite with no re-apply step (Wave 8 #285 already
  pins the preserve-format contract). Wave 1 #420 and Wave 3 #308 added a
  universal :meth:`.ColorFormat.to_rgb` that returns |None| for a
  ``_NoneColor`` instead of raising and resolves scheme colors against
  :attr:`.SlideMaster.theme_colors` with ``a:lumMod`` / ``a:lumOff``
  applied, and Wave 9 #938 added :attr:`.Font.effective_color`, which
  walks the full inheritance chain (``a:rPr`` → paragraph ``a:defRPr`` →
  text-body ``a:lstStyle`` → slide-master ``p:txStyles``) and returns the
  RGB PowerPoint would actually render. Together these give the #884
  reporter both halves of what they asked for — rewrite the text without
  losing formatting, and read the current color of an inherited- or
  theme-colored run without the ``_NoneColor`` exception. Adds a
  regression suite ``DescribeIssue884ChangeTextAndReadColor`` under
  ``tests/test_issue_884_change_text_color_verify.py`` that pins the
  reporter's combined workflow end-to-end — single-run rewrite,
  split-token rewrite, inherited-color read via ``to_rgb()`` /
  ``effective_color``, and explicit theme-color resolution — all
  round-tripped through ``Presentation.save`` + reopen.

- verify: #1106 entrance/exit animations re-verified end-to-end through the
  #102 public ``BaseShape.set_animation`` API. Extends
  ``tests/test_issue_1106_entrance_exit_verify.py`` with a
  ``DescribeIssue1106ReporterNarrativeDemo`` suite that authors
  ``MSO_ANIMATION_TYPE.FADE_IN`` on one shape and
  ``MSO_ANIMATION_TYPE.FADE_OUT`` on a second shape, saves to a real
  on-disk ``tmp_path`` file (and a ``tempfile.NamedTemporaryFile``
  variant), reopens the deck, and asserts the entrance + exit effects
  survive through every documented public-API surface a user would
  reach for —
  :attr:`pptx.shapes.base.BaseShape.animation`,
  :attr:`pptx.slide.Slide.has_animations`, and
  :attr:`pptx.slide.Slide.animation_sequence`.

- verify: #1112 (``SlideShapes.add_picture`` + SVG input) resolved by
  Wave 2 #652. Issue #1112
  (https://github.com/scanny/python-pptx/issues/1112) reported that
  calling ``slide.shapes.add_picture(<some.svg>, ...)`` surfaced an
  opaque ``PIL.UnidentifiedImageError`` — the caller got no actionable
  signal that python-pptx cannot embed SVG natively and no hint that a
  pre-rasterize workaround exists. Wave 2 #652
  (``feat/issue-652-svg-placeholder-handling``, commit ``66f53a9b``)
  addressed the root cause at the package-level image-part factory by
  sniffing the incoming image bytes for an ``<svg`` root element and
  raising the new :class:`pptx.exc.UnsupportedImageTypeError` with a
  message that names the SVG format as the cause, recommends
  pre-rasterizing to PNG (``cairosvg``, ``svglib`` + ``reportlab``,
  Pillow + librsvg, etc.), and cross-refs the
  "Inserting SVG images" section of
  ``docs/user/placeholders-using.rst``. Because
  :meth:`SlideShapes.add_picture` funnels through
  ``slide.part.get_or_add_image_part`` →
  ``package.get_or_add_image_part`` →
  ``_ImageParts.get_or_add_image_part``, the #652 sniff automatically
  covers the :meth:`~.SlideShapes.add_picture` entry-point the #1112
  reporter used on top of the :meth:`~.PicturePlaceholder.insert_picture`
  path #652's own behave scenario exercised. #1112 is therefore a
  duplicate of #652. Adds a regression suite
  ``DescribeIssue1112SvgAddPictureVerify`` under
  ``tests/test_issue_1112_svg_picture_verify.py`` that pins the
  reporter's exact workflow end-to-end: path and ``BytesIO`` SVG inputs
  both raise ``UnsupportedImageTypeError`` (not ``PIL.UnidentifiedImageError``),
  the message carries the pre-rasterize workaround hint documented at
  ``docs/user/placeholders-using.rst``, the failed call is atomic (no
  ``p:pic`` is appended to the slide's ``p:spTree``), and the
  pre-rasterized-PNG happy path still round-trips through
  :meth:`Presentation.save` + reopen so the documented workaround is
  exercised through the same public API.

- verify: #696 (move a slide from one .pptx to another) resolved by Wave 1
  #1036 + Wave 7 #934. Issue #696
  (https://github.com/scanny/python-pptx/issues/696) asked how to graft a
  slide found by text search in one deck into another deck (the reporter's
  "Nuestra experiencia en su industria" industry-experience slide); the
  shipped recipe composes :meth:`Slides.add_slide_from_external` (Wave 1/2
  #1036, promoted to the full-fidelity cloning pipeline by Wave 7 #934) +
  :meth:`Slides.move_slide` + :meth:`Slides.delete` to pull a slide from a
  library deck, position it where a placeholder slide was, and drop the
  placeholder. For the "append every slide of ppt2 to ppt1" shortcut,
  :meth:`Presentation.merge` does it in one call. Adds a regression suite
  ``DescribeIssue696MoveSlideCrossPptx`` under
  ``tests/test_issue_696_move_slide_cross_pptx_verify.py`` that pins the
  reporter's text-keyed slide-location traversal, the graft-and-replace
  workflow via ``add_slide_from_external`` + ``move_slide`` + ``delete``,
  picture-bearing save/reopen round-trip, source-not-mutated, whole-deck
  merge as the one-shot alternate path, and same-presentation self-copy.

- verify: #944 (TreeMap and ScatterPlot chart support) — *scatter half*
  resolved; treemap half cross-referenced to the existing design doc.
  Issue #944 asked for two chart kinds. The **scatter** half has
  shipped all along: :meth:`Shapes.add_chart` accepts all five
  ``XL_CHART_TYPE.XY_SCATTER*`` members
  (``XY_SCATTER``, ``XY_SCATTER_LINES``, ``XY_SCATTER_LINES_NO_MARKERS``,
  ``XY_SCATTER_SMOOTH``, ``XY_SCATTER_SMOOTH_NO_MARKERS``) and
  dispatches to ``_XyChartXmlWriter`` which emits a schema-valid
  ``c:scatterChart`` whose ``c:scatterStyle/@val``
  (``"lineMarker"``/``"smoothMarker"``) and optional
  ``c:marker/c:symbol="none"`` select the visual variant; the read
  path is covered by :class:`XyPlot` / :class:`XySeries`. Adds a
  regression suite ``DescribeIssue944ScatterVerify`` under
  ``tests/test_issue_944_scatter_verify.py`` that pins each of the
  five enum members through the real :meth:`Shapes.add_chart` API,
  asserts the ``scatterStyle`` distinction and the *_NO_MARKERS*
  marker-suppression shape, and round-trips every variant through
  ``Presentation.save`` + reopen. The **treemap** half of #944 is a
  duplicate of `#371`_ and is blocked on the F4 chartex-foundation
  work; see ``docs/dev/analysis/chartex-treemap.rst`` for the full
  design analysis (which now cross-references #944).

.. _#371: https://github.com/scanny/python-pptx/issues/371

- verify: #641 (read header/footer data from a pptx) resolved by Wave 1
  #201 + the existing placeholder-access API. Issue #641
  (https://github.com/scanny/python-pptx/issues/641) asked for a
  supported way to introspect the header/footer configuration of an
  opened ``.pptx`` — both the master-/layout-level ``p:hf`` visibility
  toggles (slide number, header, footer, date on/off) and the actual
  text content of header/footer/date/slide-number placeholders on
  individual slides. Wave 1 #201 landed :class:`pptx.slide._HeaderFooter`
  as a getter/setter proxy returned by
  :attr:`.SlideMaster.header_footer` and :attr:`.SlideLayout.header_footer`,
  exposing ``slide_number_visible`` / ``header_visible`` /
  ``footer_visible`` / ``date_visible`` backed by the ``sldNum`` /
  ``hdr`` / ``ftr`` / ``dt`` attributes of the ``p:hf`` child (with the
  XSD "absent => True" default). Reading the text inside a
  footer/date/slide-number placeholder on a slide requires no new API
  — ``slide.placeholders[idx].text_frame.text`` with ``idx`` in
  ``{10, 11, 12}`` (date / footer / slide number, the PowerPoint
  convention) is sufficient. Adds a regression suite
  ``DescribeIssue641HeaderFooterRead`` under
  ``tests/test_issue_641_hf_read_verify.py`` that pins the two reads
  end-to-end against a round-tripped package: master/layout ``p:hf``
  flag reads for every flag + default-inheritance behaviour, and
  idx-based and type-based reads of the actual footer / date /
  slide-number placeholder text content.

- verify: #1013 ("color is empty when set via placeholder") resolved by
  Wave 6 #938. Issue #1013
  (https://github.com/scanny/python-pptx/issues/1013) reported that
  reading ``run.font.color`` on a placeholder whose colour was set
  indirectly — e.g. at the paragraph level via ``paragraph.font.color.rgb
  = …``, or inherited unchanged from the slide-master's ``p:bodyStyle`` —
  came back "empty" (``color.type is None`` and ``color.rgb`` raising),
  even though PowerPoint rendered the placeholder in a visible colour.
  The symptom is the legacy ``font.color`` contract: it inspects only the
  run's own ``a:rPr/a:solidFill`` and reports ``None`` for any run that
  inherits its colour from the paragraph, the text-body ``a:lstStyle``,
  or the slide-master ``p:txStyles`` fallback. Wave 6 #938 (commit
  ``cec3edc1``, "feat(text): #938 add Font.effective_color with
  inheritance walk") shipped the read-side fix by adding
  :attr:`.Font.effective_color`, a read-only property that returns the
  |RGBColor| PowerPoint would actually render — walking the full
  inheritance chain (run ``a:rPr`` → paragraph ``a:pPr/a:defRPr`` → text
  body ``a:lstStyle/a:lvl{N}pPr/a:defRPr`` → master
  ``p:txStyles``/``bodyStyle``-``otherStyle``-``titleStyle``), resolving
  ``a:schemeClr`` against the slide-master's theme, and applying any
  ``a:lumMod``/``a:lumOff`` tint/shade siblings via the Wave-1 #420 +
  Wave-3 #308 :meth:`.ColorFormat.to_rgb` code path. The legacy
  ``font.color`` contract is intentionally unchanged, so callers that
  rely on it to distinguish "explicit" from "inherited" still can; #1013
  adds a new read path, it does not mutate the old one. Adds a regression
  suite ``DescribeIssue1013PlaceholderColorVerify`` under
  ``tests/test_issue_1013_placeholder_color_verify.py`` that pins the
  placeholder-flavoured scenarios the #1013 framing emphasises — an
  untouched title or body placeholder inheriting the master's default
  colour (explicit-RGB and scheme-colour variants), a placeholder whose
  colour was set at the paragraph-``defRPr`` level, a ``tx1`` body-style
  luminance-tinted with ``a:lumMod``/``a:lumOff``, the
  explicit-run-colour-wins precedence, a save-and-reopen round-trip for
  both the paragraph-level and master-inherited cases, and pins the
  unchanged legacy ``font.color`` contract so a silent behavioural drift
  is caught.

- verify: #115 (``Chart.update_cached_values()`` from linked Excel
  workbook) resolved by Wave 2 #381. Issue #115 asked for a
  programmatic way to refresh a chart's displayed values after its
  backing data was updated out-of-band — the reporter's workflow was
  "weekly presentation with lot of charts" whose embedded ``.xlsx``
  was rewritten from an ODBC source, without the manual per-chart
  *Refresh Data* (``Alt+F5``) click in PowerPoint. Wave 2 #381 shipped
  :meth:`Chart.update_cached_values`, which re-reads the embedded
  workbook, resolves every ``<c:f>`` cell reference under the chart,
  and rewrites the sibling ``<c:numCache>`` / ``<c:strCache>`` subtrees
  so the chart renders the workbook's current values with no
  PowerPoint interaction. Adds a regression suite
  ``DescribeIssue115UpdateCachedValuesVerify`` under
  ``tests/test_issue_115_update_cached_values_verify.py`` pinning the
  reporter's end-to-end workflow (author chart → rewrite embedded
  xlsx → refresh → save → reopen), the "refresh every chart on every
  slide" bulk case, the docstring-level edge cases (no-op on charts
  without an embedded workbook; unresolved cell references left
  untouched; idempotent), and XY/scatter coverage.

- fix: #636 ``_Cell.merge()`` now collapses a full table-span merge
  instead of emitting a merged-cell range PowerPoint silently drops.
  When the selected range covers every column of the table, the rows
  beneath the origin row are removed and their heights absorbed by the
  origin row; symmetrically, a range covering every row collapses into
  the leftmost column of the range, with the removed columns' widths
  absorbed. The surviving row (or column) retains any residual
  horizontal (or vertical) merge within it, matching PowerPoint's own
  rendering of such merges. The graphic-frame dimensions are preserved
  because the absorbed rows / columns contribute their original size to
  the surviving axis. Regression covered by the new ``_Cell.merge() in
  single-column table deletes merged rows (#636)`` and ``_Cell.merge()
  in single-row table deletes merged columns (#636)`` scenarios in
  ``features/tbl-cell.feature`` plus
  ``Describe_Cell.and_it_collapses_a_full_column_span_merge`` and
  ``and_it_collapses_a_full_row_span_merge`` unit tests.

- fix: #864 make slide layouts authored in Google Slides identifiable.
  Google Slides exports ``.pptx`` files whose ``p:sldLayout`` elements
  frequently lack a ``@type`` attribute and leave ``p:cSld/@name`` empty —
  stripping the two handles python-pptx used to give each layout a stable
  identity. :attr:`.SlideLayout.name` now falls back to a positional
  ``"Layout N"`` (1-based index within the parent master's
  ``slide_layouts``) when ``@name`` is empty, mirroring the existing
  :attr:`.SlideMaster.name` behaviour landed for issue #679. A new
  :attr:`.SlideLayout.slide_layout_type` property exposes the underlying
  ``ST_SlideLayoutType`` token (``"title"``, ``"obj"``, ``"twoObj"``,
  ``"titleOnly"``, ``"blank"``, ``"secHead"``, ``"picTx"``, ``"cust"``, ...;
  defaults to ``"cust"`` per the schema when the attribute is absent), and
  a new :meth:`.SlideLayouts.get_by_type` helper provides a
  name-independent lookup path for decks originating in Google Slides.

- Add #845 :attr:`Slide.show_master_shapes` — read/write ``bool`` mirroring
  PowerPoint's *Hide Background Graphics* checkbox (*Design > Format
  Background*). Assigning ``False`` writes ``p:sld/@showMasterSp="0"``
  so the master's non-placeholder shapes (company logo, footer bar,
  etc.) no longer render underneath the slide; master placeholders
  continue to inherit normally. Assigning ``True`` (the schema default)
  removes the attribute for a minimal round-trip. Adds the
  ``showMasterSp`` ``OptionalAttribute`` on ``CT_Slide`` plus unit,
  acceptance, and documentation coverage.

- feat: #886 Add :attr:`.CategoryAxis.tick_mark_skip` and
  :attr:`.CategoryAxis.tick_label_skip` read/write ``int`` properties
  exposing ``c:catAx/c:tickMarkSkip/@val`` and ``c:catAx/c:tickLblSkip/@val``.
  These thin out dense category axes by drawing a major tick / tick-label
  only every Nth category (``1`` — the default — draws every one; ``2``
  draws every other, and so on). Assigning ``1`` removes the backing
  element so the XML stays minimal; a value less than ``1`` raises
  :class:`ValueError` (ECMA-376 ``ST_Skip`` mandates a positive integer).

- feat: #828 Add :attr:`.Chart.series_in_rows` read-only property that
  reports whether the chart's source data is organised by rows or by
  columns — the state PowerPoint's *Chart Design > Switch Row/Column*
  toggles. OOXML does not persist a dedicated ``@switchRowCol``
  attribute, so the property infers orientation from the shape of the
  ``c:f`` cell references under the first ``c:ser``: categories running
  vertically (single column, multiple rows) mean series-in-columns
  (``False``); categories running horizontally (single row, multiple
  columns) mean series-in-rows (``True``). Returns ``None`` when the
  orientation cannot be determined — for instance when the chart has
  no series, uses inline literals instead of workbook references, or is
  an XY/scatter / bubble chart without a category axis. Read-only; the
  supported way to flip orientation programmatically is to re-author
  via :meth:`.Chart.replace_data` with the data transposed.

- Add #236 :attr:`BaseShape.hover_action` — an :class:`ActionSetting` proxy
  parallel to :attr:`BaseShape.click_action` but bound to the shape's
  ``a:hlinkMouseOver`` element instead of ``a:hlinkClick``. A hover
  action fires as the slideshow viewer's mouse pointer passes over the
  shape without a click. The returned |ActionSetting| exposes the same
  API surface as the click variant — ``action``, ``hyperlink.address``,
  ``target_slide``, ``screen_tip``, and
  ``set_sound()`` / ``remove_sound()`` / ``sound`` — so hover hyperlinks,
  slide-jumps, tooltips, and sounds can all be authored through the
  existing proxy.

- Add #730 :meth:`GroupShape.ungroup` which dissolves a group shape in
  place, hoisting each direct child onto the slide's top-level shape
  tree at its slide-relative effective rectangle and removing the
  now-empty ``p:grpSp``. The freed shapes are returned in the same
  z-order they held inside the group. Nested groups are handled
  transparently via the Wave 3 #925 ``effective_*`` cascade -- each
  child's slide rectangle is composited through every enclosing group
  transform before the hoist so shapes render at their original slide
  positions regardless of nesting depth; a freed child group keeps its
  internal ``a:chOff`` / ``a:chExt`` intact (with its own ``a:off`` /
  ``a:ext`` rewritten to its slide rectangle) so its descendants still
  render exactly where they did before.

- Add #675 :attr:`Font.highlight_color` — read/write
  :class:`~pptx.dml.color.ColorFormat` for the text-highlight
  (text-background) swatch PowerPoint exposes on the Home ribbon.
  Corresponds to ``a:rPr/a:highlight`` and supports both RGB and
  theme-color settings. Companion :meth:`Font.clear_highlight_color`
  removes any explicit highlight, restoring inheritance. Registers
  ``a:highlight`` as ``CT_Color`` in the oxml layer and adds a
  ``highlight`` ``ZeroOrOne`` descriptor to
  ``CT_TextCharacterProperties``.

- verify: #1095 (apply a POTX / PPTX template to existing slides) resolved
  by composing #1070 (POTX open) + #310 (:meth:`Presentation.strip_slides`)
  + #934 (:meth:`Presentation.merge`). ``Presentation("brand.potx")
  .strip_slides()`` yields a deck carrying only the template's masters,
  layouts, theme, and embedded fonts; ``merge(content_deck)`` then grafts
  the content slides onto it at full fidelity. Adds an
  ``Applying a POTX/PPTX template to existing content`` section to
  ``docs/user/presentations.rst`` documenting the recipe, the layout-
  binding behaviour, and the caveat that :meth:`Presentation.save` always
  writes a regular ``.pptx`` content type.

- verify: #403 resolved by #934. Issue #403 asked for "a merge function
  that combines the two filled [.pptx] file[s] into a single [.pptx]",
  which is exactly the workflow :meth:`Presentation.merge` now provides:
  ``a = Presentation("a.pptx"); a.merge(Presentation("b.pptx"));
  a.save("c.pptx")``. Wave 7 #934 shipped the merge API with full-fidelity
  cloning of pictures, media, charts (distinct embedded workbook per
  merged chart), OLE objects, and external hyperlinks in the target
  package, and guarantees the source presentation is not mutated by the
  merge — so each input template can be filled separately, combined
  without side-effects, and served to additional merges. Adds a
  regression suite ``DescribeIssue403MergePresentationFile`` under
  ``tests/test_issue_403_merge_presentation_verify.py`` that pins the
  "combine two filled .pptx files into a single .pptx" workflow
  end-to-end (save + reopen round-trip, distinct-workbook check,
  three-deck chaining, source-not-mutated, self-merge rejection).

- verify: #954 resolved by Wave 5. Issue #954
  (https://github.com/scanny/python-pptx/issues/954) reported that
  ``SlideShapes.add_movie()`` corrupted a ``.pptx`` whenever the
  slide already carried a pre-existing ``p:timing`` wrapped inside an
  ``mc:AlternateContent``/``mc:Choice`` block — the form PowerPoint
  emits when timing content references a 2010+ extension such as a
  ``p14:morph`` trigger. Wave 5 (commit ``95030bb7``, "fix(movie):
  merge p:video into wrapped p:timing (#954)") addressed this by
  teaching ``CT_Slide.get_or_add_childTnLst`` to locate the existing
  ``p:timing`` whether plain or ``mc:AlternateContent``-wrapped, and
  to merge the new ``p:video`` into its ``p:childTnLst`` in place
  rather than appending a duplicate sibling. Adds an end-to-end
  verify suite ``DescribeIssue954MovieCorruptionVerify`` under
  ``tests/test_issue_954_movie_corruption_verify.py`` that pins the
  scenario from the user's perspective — authoring a slide whose
  timing carries a ``p14:morph`` trigger, calling
  :meth:`SlideShapes.add_movie`, and round-tripping the saved pptx
  to confirm the file reopens cleanly with exactly one wrapped
  ``p:timing`` containing both the original extension trigger and
  the newly merged ``p:video``.

- verify: #1016 ``table.columns.add()`` / ``table.rows.add()`` resolved by
  #832 + #895 (with #837 bracketing the delete path). The #1016 reporter,
  migrating an Excel-to-PowerPoint table generator from python-pptx 0.6.23
  to v1.0.0, asked why ``table.columns.add()`` / ``table.rows.add()`` now
  raised ``AttributeError: '_ColumnCollection' object has no attribute
  'add'`` and whether the feature had been removed, citing PowerPoint's
  ``Rows.Add`` / ``Columns.Add`` methods as the expected API shape. The
  API had never actually shipped in 0.6.x; the fork landed it across
  three independent PRs:
  :meth:`._RowCollection.add` (#832, Wave 2,
  ``feat/issue-832-table-row-add``) appends a row at the bottom of the
  table (one empty cell per column, height inherited from the last
  existing row or defaulting to 370,840 EMU) and recomputes the
  graphic-frame height; :meth:`._ColumnCollection.add` (#895, Wave 6,
  ``feat/issue-895-table-column-mutation``) appends an ``a:gridCol`` and
  an empty ``a:tc`` in every row (width inherited from the last existing
  column or defaulting to 914,400 EMU = 1 inch) and recomputes the
  graphic-frame width; and the delete counterparts land via #837 (Wave
  3, :meth:`._Row.delete` / :meth:`._RowCollection.remove`) and #895
  (:meth:`._Column.delete` / :meth:`._ColumnCollection.remove`). #1016
  is therefore a duplicate of the #832 + #895 feature PRs. Adds a
  regression suite ``DescribeIssue1016TableAddRowsColumns`` under
  ``tests/test_issue_1016_table_add_rows_columns.py`` that pins the
  reporter's Excel-driven "grow procedurally" workflow in both directions
  (``rows.add`` and ``columns.add``), the default and explicit
  height/width overloads, the graphic-frame invariant, a mixed
  grow-then-shrink sequence that exercises the #837 / #895 delete
  counterparts, and save + reopen round-trips for both directions.

- verify: #235 (add 3-D pie chart type) resolved by Wave 7 #266. The #235
  reporter asked for ``Shapes.add_chart`` to accept
  :attr:`XL_CHART_TYPE.THREE_D_PIE`; on pre-fork python-pptx that call
  raised ``NotImplementedError: XML writer for chart type THREE_D_PIE
  (-4102) not yet implemented`` because ``ChartXmlWriter`` had no
  3D-chart builders. #266 (``feat/issue-266-3d-chart-types``, Wave 7)
  shipped ``_Pie3DChartXmlWriter`` (together with the other 12
  ``THREE_D_*`` members — area/bar/column/line variants) so the factory
  now emits schema-valid ``c:pie3DChart`` XML plus a ``c:view3D``
  sibling on ``c:chart`` with PowerPoint's default 3D angles
  (``rotX=15``, ``rotY=20``, ``rAngAx=1``, ``depthPercent=100``). Both
  :attr:`XL_CHART_TYPE.THREE_D_PIE` and
  :attr:`XL_CHART_TYPE.THREE_D_PIE_EXPLODED` are covered — the exploded
  variant injects a ``c:explosion val="25"`` child on the series. Adds
  a regression suite ``DescribeIssue235ThreeDPieVerify`` under
  ``tests/test_issue_235_3d_pie_verify.py`` that pins the authoring
  path through ``Shapes.add_chart``, the default ``c:view3D`` angles,
  the plain-vs-exploded explosion distinction, series-data carriage,
  and a save + reopen round-trip for both enum members. Read-access
  for ``c:pie3DChart`` (``chart.plots`` iteration) remains the same
  known limitation as the other non-area 3D chart types per the #266
  MVP write-only contract and is not in scope for #235.

- verify: #1068 (edit individual data labels) resolved by
  ``series.points[i].data_label`` — the per-point |DataLabel| proxy
  wrapping the ``c:dLbl`` element addressed by ``c:idx`` matching the
  point index. The capability has existed since ``Point.data_label``
  was first added; the Wave-era work that closed the adjacent gaps —
  ``feat/issue-716-datalabel-border`` (Wave 1) added
  ``DataLabel.format`` (a |ChartFormat| wrapping the individual
  ``c:dLbl``) with ``.fill`` / ``.line`` / ``.shadow``, and
  ``feat/issue-560-data-label-colors`` (Wave 7) added the series-level
  analogue ``DataLabels.format`` — round out the per-label API. The
  #1068 reporter's screenshot (different label contents on one point
  only) is achieved by assigning
  ``series.points[i].data_label.text_frame.text = "..."``, which
  writes a ``c:dLbl/c:tx/c:rich`` subtree PowerPoint renders in place
  of the series-level ``show_value`` / ``show_series_name`` flags for
  that point. Adds a regression suite
  ``DescribeIssue1068IndividualDataLabels`` under
  ``tests/test_issue_1068_individual_data_labels.py`` that pins
  ``Series.points[i].data_label`` on category, XY, and bubble series,
  round-trips per-point ``text_frame``, ``position``, ``font``, and
  ``format.fill`` through ``Presentation.save`` + reopen, and asserts
  per-point independence (customizing one label does not touch its
  siblings and writes exactly one ``c:dLbl`` per customized point
  with a matching ``c:idx/@val``).

- verify: #805 (replace audio / video in existing slide) resolved by
  Wave 6 #784. The #805 reporter asked for a way to swap the audio or
  video binary behind an already-authored media shape without dropping
  and re-adding it — the media analogue of ``Picture.replace_image``
  (#116). Wave 6 #784 (``feat/issue-784-replace-audio``) shipped
  :meth:`Movie.replace_media` ``(new_path_or_file, mime_type=None)``
  which loads the replacement bytes into a fresh |MediaPart|, rewires
  the shape's ``a:videoFile`` / ``a:audioFile`` ``@r:link`` and
  ``p14:media`` ``@r:embed`` rIds, and drops the old relationships so
  the previous media part is eligible for garbage-collection on save.
  Position, size, cropping, poster frame, hyperlink, and ``p:timing``
  entries are preserved; the shape's media-element tag
  (``a:videoFile`` vs ``a:audioFile``) is also preserved — modality
  switching (audio ↔ video) remains a documented follow-up requiring
  :meth:`SlideShapes.add_movie` + :meth:`Shape.delete`. Adds a
  verify-and-close regression suite
  ``DescribeIssue805ReplaceMediaVerify`` under
  ``tests/test_issue_805_replace_media_verify.py`` that pins audio and
  video in-place swaps, geometry preservation, the a:audioFile /
  a:videoFile modality guard, the not-a-media-pic error path, and a
  full ``Presentation.save`` + reopen round-trip confirming the
  rewired relationships survive packaging.

- verify: #819 (image replacing) resolved by #834 / #116
  :meth:`.Picture.replace_image`. The #819 reporter wanted to locate a
  picture on a later program run via its ``cNvPr/@descr`` alt-text and
  swap its pixel bytes without disturbing its position, size, or
  identifying attributes. Their sketched workaround ``deleteShape`` +
  :meth:`.SlideShapes.add_picture` failed because ``add_picture`` mints
  a fresh ``descr`` from the new source image's filename (the reporter's
  observed "replaced with image.png" symptom), so the next run's
  find-by-descr loop could no longer locate the shape. Their second
  attempt reached into :class:`.ImagePart` internals via a non-existent
  ``SlidePart.relatedpart`` method and rewrote ``_blob`` directly — a
  path that cannot work correctly because image parts are shared across
  shapes via content-hash deduplication. Wave 6 #834
  (``feat/issue-834-picture-replace-image``) shipped
  :meth:`.Picture.replace_image` which answers #819 directly: it adds
  (or reuses, via :meth:`.SlidePart.get_or_add_image_part`) an image
  part, rebinds ``p:pic/p:blipFill/a:blip/@r:embed`` to the new rId,
  and drops the previous image relationship — leaving position, size,
  rotation, cropping, masking shape, outline, and crucially the
  ``cNvPr/@name`` + ``cNvPr/@descr`` identity of the shape untouched,
  so the reporter's find-by-descr loop keeps working across every
  subsequent program run. #819 is therefore a duplicate of #834 / #116.
  Adds a regression suite ``DescribeIssue819ImageReplacing`` under
  ``tests/test_issue_819_picture_replace.py`` that pins the reporter's
  exact workflow — find a picture by :attr:`.BaseShape.alt_text`,
  :meth:`.Picture.replace_image` its pixels, save + reopen, confirm the
  descr still locates the shape, and repeat the cycle — alongside the
  file-like-replacement form matching the reporter's ``open(path,
  "rb")`` snippet.

- verify: #782 (programmatic access to click-action sounds) resolved by
  Wave 2 #734. The shipped surface —
  :attr:`~pptx.action.ActionSetting.sound`,
  :meth:`~pptx.action.ActionSetting.set_sound`, and
  :meth:`~pptx.action.ActionSetting.remove_sound`, together with the
  :class:`~pptx.action.Sound` read-view and the
  :class:`~pptx.media.Audio` value object — covers every item the
  reporter asked for: read the sound name and embedded audio bytes from
  an existing shape, attach a new ``<a:snd>`` to a click or hover
  action, and remove one without dropping the enclosing hyperlink.
  Adds a ``Describe_Issue782_SoundAccess`` regression suite in
  ``tests/test_action.py`` pinning the read / set / replace / remove /
  hover / missing-sound branches at the proxy-object level so the
  user-visible contract cannot silently regress.

- fix: #287 Bubble chart data points now render in LibreOffice (and other
  strict OOXML readers). The ``c:bubbleChart`` element emitted for a freshly
  authored bubble chart was missing the optional ``c:bubble3D`` child
  specified by ``CT_BubbleChart`` between ``c:dLbls`` and ``c:bubbleScale``.
  PowerPoint treats the chart-level value as the default for every series
  and ignored its absence, but older LibreOffice versions silently skipped
  drawing the bubbles (only the axes rendered). ``_BubbleChartXmlWriter``
  now emits ``<c:bubble3D val="0"/>`` for ``BUBBLE`` (and ``val="1"`` for
  ``BUBBLE_THREE_D_EFFECT``) in the schema-declared position, matching the
  per-series ``c:bubble3D`` already written on each ``c:ser``.

- fix: #852 ``chart.category_axis.visible = False`` (and the same property on
  any value or date axis) no longer produces a ``<c:delete/>`` element with a
  missing ``val`` attribute. Although the OOXML schema specifies the default
  as ``true`` for ``CT_Boolean``, PowerPoint does not honor that default on
  ``c:delete`` (analogous to the existing ``c:overlay`` workaround via
  ``CT_Boolean_Explicit``), so python-pptx now always writes ``val="1"`` or
  ``val="0"`` explicitly when assigning to :attr:`_BaseAxis.visible`. Hidden
  axes authored by python-pptx are now correctly rendered as hidden by
  PowerPoint.

- Add #298 ``Chart.plot_area`` returning a ``PlotArea`` proxy with a
  ``format`` property that yields a ``ChartFormat`` object. This exposes
  ``.fill``, ``.line``, and ``.shadow`` on the chart's ``c:plotArea`` /
  ``c:spPr`` shape-properties block, mirroring the ``format`` pattern
  already in place on ``ChartTitle`` and ``DataLabel`` / ``DataLabels``.
  Callers can now colour or outline the plot-area rectangle directly —
  e.g. ``chart.plot_area.format.line.color.rgb = RGBColor(0xFF, 0, 0)`` —
  without dropping into XML. Adds a ``ZeroOrOne("c:spPr")`` descriptor on
  ``CT_PlotArea`` so the ``c:spPr`` child is inserted in schema-declared
  order (before any ``c:extLst``).

- Add #546 text shadow — :class:`pptx.text.text.Font` now exposes a
  ``shadow`` property (and a broader ``effect_format`` property for the
  full ``a:effectLst`` family: shadow / glow / reflection / soft-edge).
  ``run.font.shadow`` returns a real
  :class:`~pptx.dml.effect.ShadowFormat` with the familiar
  ``inherit`` / ``blur_radius`` / ``distance`` / ``direction`` / ``color``
  API, writing its XML under ``a:rPr/a:effectLst/a:outerShdw`` — the
  ECMA-376 §20.1.8 + §21.1.2.3.8 location for a text-run shadow. Adds a
  ``ZeroOrOne`` ``effectLst`` descriptor on ``CT_TextCharacterProperties``
  (``a:rPr`` / ``a:defRPr`` / ``a:endParaRPr``) and a regression suite
  ``DescribeIssue546TextShadow`` under
  ``tests/test_issue_546_text_shadow.py`` pinning the reporter's exact
  workflow (set a shadow, round-trip through
  :meth:`Presentation.save` + reopen, assert every knob survived).

- feat: #846 Add :attr:`.BarPlot.has_series_lines` read/write boolean
  and :attr:`.BarPlot.series_lines` accessor returning a new
  :class:`~pptx.chart.plot.SeriesLines` proxy. Series lines
  (``c:serLines``) connect segment tops across adjacent series in a
  stacked bar or stacked column chart, helping readers compare
  segment-to-segment change. ``BarPlot.series_lines.format`` returns a
  :class:`~pptx.dml.chtfmt.ChartFormat` so the line color, width, and
  dash style can be customised via the familiar
  ``.format.line`` / ``.format.fill`` / ``.format.shadow`` handles.
  Toggling off removes the element entirely, keeping the XML minimal.
  The new element is wired into ``CT_BarChart`` at the XSD-mandated
  position (after ``c:overlap`` and before ``c:axId``), and ``c:serLines``
  is registered as :class:`~pptx.oxml.chart.axis.CT_ChartLines` so the
  existing ``spPr`` descriptor carries the format parent-chain.

- feat: #859 Add :attr:`.Chart.display_blanks_as` read/write property
  exposing PowerPoint's "Hidden and Empty Cells" / "Show empty cells as"
  setting (``c:dispBlanksAs``) as a member of the new
  :ref:`XlDisplayBlanksAs` enumeration (``GAPS``, ``ZERO``,
  ``INTERPOLATED``). The #859 reporter can now suppress zero-valued bars
  in a stacked bar chart by assigning
  ``XL_DISPLAY_BLANKS_AS.GAPS`` — blank cells are then drawn as gaps
  rather than as zero-height bars. Assigning
  :attr:`~XL_DISPLAY_BLANKS_AS.ZERO` (the XSD default) removes any
  existing ``c:dispBlanksAs`` element so the XML stays minimal.

- feat: #299 Add chart-series trendlines. ``Series.add_trendline(type,
  ...)`` attaches a fitted-curve overlay (linear, logarithmic, polynomial
  of order 2–6, power, exponential, or moving-average) to a chart series
  and returns a :class:`~pptx.chart.series.Trendline` proxy;
  ``Series.trendlines`` enumerates any existing trendlines and
  ``Trendline.delete()`` removes one. Each trendline exposes
  ``trendline_type``, ``order`` (polynomial), ``period``
  (moving-average), ``forward`` / ``backward`` extrapolation, custom
  ``intercept``, ``display_equation`` / ``display_r_squared`` on-chart
  annotation flags, a user-visible ``name``, and a
  :class:`~pptx.dml.chtfmt.ChartFormat` ``format`` handle. New enum
  :class:`pptx.enum.chart.XL_TRENDLINE_TYPE` carries the six regression
  types and round-trips ``c:trendlineType/@val``.

- Add #94 :attr:`Presentation.view_props` (editor-view settings that
  PowerPoint restores on re-open — :attr:`.ViewProps.view_type`,
  :attr:`.ViewProps.show_comments`, :attr:`.ViewProps.show_formatting`,
  plus per-view zoom properties) and :attr:`Presentation.first_slide_num`
  (``p:presentation/@firstSlideNum``, the "Number slides from" setting).
  Backed by a new :class:`.ViewPropsPart` (``ppt/viewProps.xml``) which
  is created lazily when the package does not already contain one. Adds
  :class:`pptx.enum.presentation.PP_VIEW_TYPE` (value-space of
  ``p:viewPr/@lastView``) and companion design notes at
  ``docs/dev/analysis/view-props.rst``.

- verify: #175 (add slide / slide layout from other presentation) resolved
  by #934 + :meth:`Slides.add_slide_from_external`. Wave 7 #934 shipped
  :meth:`Presentation.merge` for whole-deck full-fidelity copy, and
  promoted :meth:`Slides.add_slide_from_external` (originally the #1036
  "basic" copy path) to the same full-fidelity cloning pipeline — so the
  #175 reporter's first ask (adding a foreign slide to a presentation) is
  fully served. The second ask (adding a foreign *layout* to a
  presentation) is served by a documented workaround: open the
  layout-source deck as the base, :meth:`Presentation.merge` the
  content-carrying decks into it, then :meth:`Slides.delete` any unwanted
  starter slides. The target keeps its full layout list (delete only
  removes slides, not layouts) while gaining the content of every merged
  deck. Adds a regression suite
  ``DescribeIssue175CrossPresentationSlideLayoutCopy`` under
  ``tests/test_issue_175_cross_presentation_slide_layout_copy.py`` that
  pins per-slide copy, whole-deck merge, and the layout-import workaround
  end-to-end (including save + reopen round-trip). Refreshes the
  ``Copying a slide from one presentation to another`` section of
  ``docs/user/slides.rst`` — the pre-Wave-7 claim that chart /
  OLE-object / media slides raise ``NotImplementedError`` is no longer
  accurate — and adds ``Merging every slide of another presentation``
  and ``Importing slide layouts from another presentation`` sections
  documenting :meth:`Presentation.merge` and the layout-import recipe.

- fix: #1084 avoid ``AttributeError: 'Part' object has no attribute 'image'``
  when opening a ``.pptx`` whose ``[Content_Types].xml`` declares an image
  with a non-canonical MIME type or casing. Registers ``image/tif`` as an
  alias for ``image/tiff`` (mirroring the existing ``image/jpg`` alias for
  ``image/jpeg``), and makes ``PartFactory._part_cls_for`` lookup
  case-insensitive so content-types such as ``Image/Tiff`` or ``IMAGE/PNG``
  also resolve to :class:`ImagePart` rather than falling through to the
  generic :class:`Part`. Per RFC 2046 §4.1 MIME type tokens are
  case-insensitive.

- fix: #844 restore ``from pptx.oxml import qn`` backwards compatibility.
  ``qn()`` had been relocated to :mod:`pptx.oxml.ns` during an early
  refactor without being re-exported from :mod:`pptx.oxml`, breaking
  long-standing ``from pptx.oxml import qn`` call sites. The name is
  now re-exported from :mod:`pptx.oxml` (and listed in ``__all__``) so
  downstream code continues to work unchanged; ``pptx.oxml.ns.qn``
  remains the canonical location.

- Add #310 ``Presentation.strip_slides()`` — remove every slide (and any
  sections) from an opened presentation while preserving its slide
  masters, slide layouts, theme, embedded fonts, and other template-
  level resources. Enables the "use an existing ``.pptx`` as a blank
  template" workflow requested in issue #310: open a branded deck, strip
  its slides, and reuse the remaining presentation as the starting
  point for a new deck. Returns ``self`` so it chains after the
  ``Presentation()`` factory call.

- feat(movie): #427 add ``autoplay`` convenience kwarg on
  :meth:`SlideShapes.add_movie`. When |True|, the newly added movie is
  configured to start automatically with the slide
  (``start_condition="withPrevious"``); the default of |False| preserves
  PowerPoint's click-to-play behavior. Layered on top of the Wave 5
  ``Movie.start_condition`` / ``Movie.start_time`` properties (#811);
  callers that need ``"afterPrevious"`` or a delayed start can still
  assign those directly on the returned shape after the
  ``add_movie(..., autoplay=True)`` call.

- feat: #508 Add :attr:`.BaseShape.alt_text` and :attr:`.BaseShape.title`
  read/write properties exposing the accessibility description (``descr``)
  and title (``title``) attributes on every shape's ``cNvPr`` element
  (auto-shape, picture, graphic frame, group-shape, and connector).
  Both properties return the empty string when the attribute is absent
  (PowerPoint's default); assigning the empty string clears the attribute.
  The values are surfaced in PowerPoint's Alt Text pane and consumed by
  screen readers for accessibility.

- feat: #329 add read/write ``TickLabels.rotation`` exposing the
  clockwise rotation (degrees) applied to axis tick labels. Maps to
  ``c:txPr/a:bodyPr/@rot`` on each axis type (``c:catAx`` / ``c:valAx``
  / ``c:dateAx``). Accepts ``int`` or ``float``; stored as 60000ths of
  a degree per OOXML. Reads ``0.0`` when no rotation is set; assigning
  ``0`` removes the attribute; negative values are normalized into
  ``[0, 360)``.

- Add #351 chart user-shapes (annotation shapes) read access. A new
  :class:`pptx.parts.chartdrawing.ChartDrawingPart` models the
  ``c:userShapes`` relationship part that PowerPoint writes when a user
  draws annotation shapes (arrows, call-outs, text boxes) on top of a
  chart, and new :attr:`pptx.chart.chart.Chart.user_shapes` /
  :attr:`~pptx.chart.chart.Chart.has_user_shapes` properties expose the
  part — the former returns the :class:`ChartDrawingPart` (or ``None``
  when no ``chartUserShapes`` relationship exists), the latter is a
  non-destructive boolean probe. The part offers
  :meth:`~pptx.parts.chartdrawing.ChartDrawingPart.iter_anchor_elements`
  and :attr:`~pptx.parts.chartdrawing.ChartDrawingPart.anchor_count` to
  enumerate the raw ``cdr:relSizeAnchor`` / ``cdr:absSizeAnchor`` anchor
  elements. This MVP is read-only (raw-``lxml`` passthrough); a proxy
  hierarchy and authoring API for chart-drawing shapes are deferred —
  see ``docs/dev/analysis/chart-user-shapes.rst`` for the full design
  note. Registering :class:`ChartDrawingPart` against the
  ``chartshapes+xml`` content-type also fixes a subtle round-trip issue
  where user-shapes parts were previously loaded as anonymous byte-blobs
  via the generic :class:`~pptx.opc.package.Part` fallback.

- Add #801 ``Movie.blob`` / ``Movie.ext`` / ``Movie.content_type`` for
  reading an embedded movie's media bytes, file extension, and MIME
  type. Mirrors the ``Picture.image.blob`` / ``.ext`` / ``.content_type``
  surface so audio / video clips can be extracted from a presentation
  without touching the package internals. All three properties return
  |None| when the shape has no associated media part.

- docs: #391 add a "Reading slide properties" section to
  ``docs/user/slides.rst`` covering ``slide.slide_id``, ``slide.name``,
  ``slide.slide_layout``, ``slide.shapes.title``, and iterating
  placeholders by ``placeholder_format.idx`` / ``type``. No API change —
  all of the referenced accessors already exist; the gap was
  documentation.

- rfctr: Resolve ``AnimationEffect`` class-name collision between
  :mod:`pptx.animation` (authoring API, Wave 5 #102) and :mod:`pptx.slide`
  (read-only introspection proxy for :attr:`.Slide.animation_sequence`,
  Wave 5 #256). :class:`pptx.animation.AnimationEffect` is the canonical
  public authoring class and is unchanged. The introspection proxy on
  ``pptx.slide`` has been renamed to :class:`.AnimationEffectView`;
  ``pptx.slide.AnimationEffect`` remains as a deprecated alias
  (``AnimationEffect = AnimationEffectView``) for one release so
  existing ``from pptx.slide import AnimationEffect`` imports keep
  working. New code should import :class:`~pptx.slide.AnimationEffectView`
  for introspection or :class:`pptx.animation.AnimationEffect` for
  authoring.

- verify: #720 resolved by #337 EA/CS font slots

- verify: #343 resolved by #337 EA/CS font slots

- verify: #838 (bounding boxes wrongly parsed for shapes inside GroupShape)
  resolved by #925's ``effective_*`` properties on :class:`.BaseShape`. The
  #838 reporter's raw ``shape.left`` / ``shape.top`` / ``shape.width`` /
  ``shape.height`` returned "wrong-looking" values for children of a resized
  group because those values are expressed in the enclosing group's child
  coordinate system (``a:chOff``/``a:chExt``) — when PowerPoint resizes the
  group it diverges ``a:ext`` from ``a:chExt`` and leaves children's raw
  offsets unchanged. #925 (``fix/issue-925-group-shape-transform``, Wave 3)
  added :attr:`.BaseShape.effective_left` / :attr:`~.BaseShape.effective_top`
  / :attr:`~.BaseShape.effective_width` / :attr:`~.BaseShape.effective_height`
  which walk every enclosing ``p:grpSp`` ancestor and apply its
  ``a:chOff``/``a:chExt`` → ``a:off``/``a:ext`` linear transform so the
  returned values are the slide-relative geometry the shape renders at. The
  pre-existing raw properties are unchanged for backwards compatibility.
  Adds a regression suite ``DescribeIssue838GroupChildBoundingBox`` under
  ``tests/test_issue_838_group_bbox_verify.py`` that authors a two-child
  group, simulates a PowerPoint 2:1 horizontal resize by diverging the
  group's ``a:ext`` from its ``a:chExt``, pins both the raw (group-local)
  and effective (slide-relative) values on each child, exercises
  compositing through nested groups, and round-trips the geometry through
  ``Presentation.save`` + reopen.

- verify: #534 (shape position/size report relative to parent GroupShape)
  resolved by #925's ``effective_*`` properties on :class:`.BaseShape`. #534
  is the sibling group-local ask to #838 — the reporter wanted a way to
  read a shape's position and size in slide-relative coordinates (composed
  through the enclosing group's transform) rather than in the enclosing
  group's local ``a:chOff``/``a:chExt`` frame that :attr:`.BaseShape.left` /
  :attr:`~.BaseShape.top` / :attr:`~.BaseShape.width` /
  :attr:`~.BaseShape.height` return. #925
  (``fix/issue-925-group-shape-transform``, Wave 3) added
  :attr:`.BaseShape.effective_left` / :attr:`~.BaseShape.effective_top` /
  :attr:`~.BaseShape.effective_width` / :attr:`~.BaseShape.effective_height`
  which walk every enclosing ``p:grpSp`` ancestor and apply its
  ``a:chOff``/``a:chExt`` → ``a:off``/``a:ext`` linear transform so the
  returned values are slide-relative — exactly the API #534 asked for. Raw
  and effective values coexist on the same shape object, so callers can
  read both the parent-relative and slide-relative rectangles. Adds a
  regression suite ``DescribeIssue534GroupRelativePositionSize`` under
  ``tests/test_issue_534_group_relative_verify.py`` that pins the
  complementary scenarios the #534 framing emphasized — an asymmetric 2D
  corner-handle resize (distinct x and y scale factors), a group whose
  ``a:chOff`` has been translated away from the slide origin, a
  picture-shape group child (to verify the transform is shape-kind
  agnostic), and the dual-report contract where raw and effective readings
  are simultaneously correct — and round-trips the composited values
  through ``Presentation.save`` + reopen.

- verify: #1106 entrance/exit animations resolved by #102 set_animation API

- verify: #1020 resolved by #62 alpha + #234 blip_fill + #515 preset geometry

- verify: #791 resolved by #837. The #791 reporter asked for a way to delete
  a specific row from a table ("I want to delete last 7th row from table")
  and sketched the workaround ``table._tbl.remove(row._tr)`` inherited from
  the #192 thread. That workaround bypassed the graphic-frame height
  invariant (cy = sum of authored row heights). #837
  (``feat/issue-837-table-row-delete``, Wave 3) shipped the supported API
  in its place: :meth:`._Row.delete` detaches the row's ``a:tr`` from its
  parent ``a:tbl`` and triggers the parent ``Table`` to recompute the
  containing graphic-frame height, and :meth:`._RowCollection.remove` is
  the collection-level form that raises :class:`ValueError` when handed a
  row from a different table. #791 is therefore a duplicate of #837. Adds
  a regression suite ``DescribeIssue791DeleteTableRow`` under
  ``tests/test_issue_791_table_row_delete.py`` that pins the reporter's
  exact "delete the 7th row of a 7-row table" scenario, the interior-row
  and collection-level variants, the foreign-row ``ValueError``, the
  graphic-frame height shrinkage, the ``a:tr`` detachment from ``a:tbl``,
  and a save + reopen round-trip (including the multi-delete case a real
  caller typically needs).

- verify: #635 (remove an image from a PowerPoint) resolved by #41
  ``BaseShape.delete`` + ``Picture.delete`` override. The #635 reporter
  walked every picture shape to extract bytes, then asked "Is there any
  library through i can remove the image from pptx?" Wave 1 #41
  (``feat/issue-41-shape-delete``) shipped :meth:`.BaseShape.delete` —
  the primitive that detaches the shape element from its parent
  ``p:spTree`` — plus a :meth:`.Picture.delete` override that additionally
  drops the ``r:embed`` relationship to the backing
  :class:`.ImagePart`, so the image bytes are garbage-collected on save
  rather than left dangling inside the saved ``.pptx`` zip. #635 is
  therefore a duplicate of #41. Adds a regression suite
  ``DescribeIssue635RemoveImage`` under
  ``tests/test_issue_635_remove_image_verify.py`` that pins the
  reporter's exact iter-pictures-and-delete-a-subset flow (both a
  direct-reference and a visitor-style variant), the orphan ImagePart
  being dropped from the package on save (the tangible "file shrinks"
  benefit over a raw element removal), the ``p:pic`` detachment from
  ``p:spTree``, and a save + reopen round-trip.

- verify: #303 resolved by #141 secondary axis + #338 combo charts

- docs: #963 "save slide as image" — close as out-of-scope (python-pptx
  does not render slides). Extend the LibreOffice headless example in the
  ``Rendering to video, PDF, or image formats`` section of
  ``docs/user/use-cases.rst`` with the ``--convert-to png`` /
  ``pdftoppm`` / ImageMagick ``convert`` recipe for exporting every slide
  as an image (``--convert-to png`` alone only writes the first slide).
- verify: #357 (customize pie-chart slice colors) resolved by the
  pre-existing ``Point.format`` API on ``CategoryPoints``. Pie-chart
  per-slice color control has been available since commit
  ``de58d605`` ("cht: add Point.format") because ``PieSeries``
  inherits from ``_BaseCategorySeries`` and therefore exposes
  ``CategoryPoints`` via ``PieSeries.points``. The idiomatic call
  pattern is
  ``chart.plots[0].series[0].points[n].format.fill.solid();
  chart.plots[0].series[0].points[n].format.fill.fore_color.rgb =
  RGBColor(...)``, which writes a ``c:dPt/c:spPr/a:solidFill/a:srgbClr``
  subtree under the ``c:pieChart/c:ser``. Adds a regression suite
  ``DescribeIssue357PieChartColors`` under
  ``tests/test_issue_357_pie_chart_colors.py`` that authors a 4-slice
  pie with distinct per-slice fills, inspects the emitted ``c:dPt``
  elements (one per slice, ``c:idx`` matching position, ``c:spPr``
  carrying the fill), round-trips through save + reopen, and also
  covers per-slice line color via ``points[i].format.line``.
- verify: #571 (number_format for category axis tick labels) resolved by
  pre-existing infrastructure. ``CategoryAxis.tick_labels.number_format``
  already writes ``c:catAx/c:numFmt/@formatCode`` (and sets
  ``@sourceLinked="0"``) via the shared ``_BaseAxis.tick_labels`` /
  ``TickLabels`` plumbing, because ``CT_CatAx._tag_seq`` has always
  included ``c:numFmt`` as a valid child. The #571 reporter's confusion
  was over PowerPoint's rendering: category-axis tick labels are plain
  strings (the category names), so a numeric ``formatCode`` has no
  visible effect unless the categories are themselves numeric (in which
  case PowerPoint treats the axis as a date / numeric scale). Adds a
  regression suite ``DescribeIssue571CategoryAxisNumberFormat`` under
  ``tests/test_issue_571_catax_number_format.py`` that pins the get /
  set / ``sourceLinked`` semantics and round-trips the assignment
  through ``Presentation.save`` + reopen so any future change that
  drops ``c:numFmt`` from ``CT_CatAx`` or decouples
  ``CategoryAxis.tick_labels`` from the shared ``TickLabels`` plumbing
  will be caught here.
- verify: #662 (data-label background color) resolved by
  ``feat/issue-716-datalabel-border`` (Wave 1) together with
  ``feat/issue-560-data-label-colors`` (Wave 7). #716 added
  ``DataLabel.format`` (a ``ChartFormat`` wrapping the individual
  ``c:dLbl``) with ``.fill`` / ``.line`` / ``.shadow``, and #560 added
  the series-level analogue ``DataLabels.format`` (wrapping
  ``c:dLbls``) plus the ``c:spPr`` registration on ``CT_DLbls`` so
  the element is inserted between ``c:numFmt`` and ``c:txPr`` in
  schema order. Together these close the #662 reporter's request to
  set a background fill on data labels — ``series.data_labels.format
  .fill.solid()`` colors every label on a series, and
  ``series.points[i].data_label.format.fill.solid()`` colors a single
  label. Leader lines (the other half of the #662 thread) remain
  out of scope. Adds a regression suite
  ``DescribeIssue662DataLabelBackground`` under
  ``tests/test_issue_662_data_label_background.py`` that round-trips
  both scopes through ``Presentation.save`` + reopen and pins
  ``c:spPr`` placement under ``c:dLbls``.
- verify: #825 (change points color in scatter plot) resolved by
  ``feat/issue-825-scatter-point-colors``. The per-point marker-color
  API has existed since ``Point.marker`` was first added (mirroring the
  series-level ``.marker`` on ``XySeries`` / ``LineSeries`` /
  ``RadarSeries``); the correct idiom on scatter charts is
  ``series.points[i].marker.format.fill`` — which writes to
  ``c:dPt/c:marker/c:spPr``, the element PowerPoint consults for the
  marker swatch. (The superficially-similar ``point.format.fill`` writes
  ``c:dPt/c:spPr``, which PowerPoint renders as the data-point's *line*
  color on marker-type charts, hence the reporter's observation that
  fill assignments had no visible effect.) Adds a regression suite
  ``DescribeIssue825ScatterPointColors`` under
  ``tests/test_issue_825_scatter_point_colors.py`` that pins the XML
  shape (``c:dPt/c:marker/c:spPr`` with no sibling ``c:dPt/c:spPr``)
  across XY_SCATTER / XY_SCATTER_LINES / XY_SCATTER_SMOOTH, locks in
  per-point independence (three points, three distinct colours, no
  cross-contamination and no duplicate ``c:dPt`` with the same
  ``c:idx``), and round-trips the authored colours through
  ``Presentation.save`` + reopen.
- verify: #830 (embed custom ``.otf`` font) resolved by
  ``feat/issue-355-font-embedding`` (Wave 3). The
  :meth:`.Presentation.embed_font` API treats its ``font_file``
  argument as an opaque byte blob and writes it into a ``FontPart``
  with content-type ``application/x-fontdata`` — the same OOXML
  envelope PowerPoint uses for both TrueType (``.ttf``) and
  OpenType (``.otf``) embedded-font containers. OpenType files
  therefore embed, save, and round-trip byte-identically under the
  same API, with no ``.otf``-specific code path required. Adds a
  regression suite ``DescribeIssue830OtfFont`` under
  ``tests/test_issue_830_otf_font.py`` that pins the path + stream
  intake for ``.otf`` files, the ``application/x-fontdata``
  content-type, a save + reopen round-trip (bytes preserved
  verbatim), and multi-style (regular + bold) OTF embedding under
  a single typeface.
- verify: #937 (arrange tickers side-by-side in table cells) resolved by
  ``feat/issue-71-cell-borders`` (Wave 3) on top of the pre-existing
  ``_Cell.margin_*`` / ``_Cell.merge`` / ``_Paragraph.add_run`` /
  ``paragraph.alignment`` APIs. Two recipes now cover the reporter's
  "tickers side-by-side" layout without any new feature surface: (A)
  one sub-cell per ticker with ``cell.border_<side>.fill.background()``
  hiding the internal edges and ``cell.merge`` spanning a header above
  the group, or (B) a single cell whose paragraph holds one
  ``paragraph.add_run()`` per ticker so the runs flow side-by-side on
  one line with independent per-run font / bold / colour. Adds a
  regression suite ``DescribeIssue937TableCellLayout`` under
  ``tests/test_issue_937_table_cell_layout.py`` that pins both recipes
  plus the ``_Cell.margin_*`` knobs (the reporter's secondary complaint
  about slide overflow) through ``Presentation.save`` + reopen.
- feat: #319 add ``Slide.is_hidden`` read/write ``bool`` property mapping
  to the ``p:sld/@show`` attribute. Reading returns ``True`` for a slide
  explicitly marked hidden (``show="0"``) and ``False`` otherwise (the
  schema default for ``@show`` is ``true``). Assigning ``True`` writes
  ``show="0"``; assigning ``False`` removes the attribute so the XML
  round-trips to the default-visible state. Hidden slides are skipped by
  PowerPoint during a normal slide-show run but remain in the package
  and in :attr:`.Presentation.slides`.
- feat: #971 add ``BaseShape.is_hidden`` read/write bool mapping to the
  ``hidden`` attribute on the shape's ``cNvPr`` element. When |True|,
  PowerPoint skips the shape during slide-show and print while leaving
  it visible in the editing surface. Available on all shape types
  (``p:sp``, ``p:pic``, ``p:cxnSp``, ``p:graphicFrame``, ``p:grpSp``).
- feat: #425 add read/write ``ActionSetting.screen_tip`` exposing the
  tooltip (ScreenTip) text that PowerPoint displays on mouse-over of a
  shape carrying a click or hover action. Backed by
  ``a:hlinkClick/@tooltip`` (or ``a:hlinkHover/@tooltip`` on hover
  actions); returns |None| when no hyperlink element is present, and
  assigning |None| or the empty string clears the attribute without
  removing the hyperlink element itself so an accompanying URL, sound,
  or slide-jump target survives.
- feat: #1030 manually set chart title position. ``ChartTitle.position``
  is a new read/write property returning ``(x, y)`` as a tuple of floats
  in the 0.0-1.0 relative-to-chart coordinate space defined by
  ``c:title/c:layout/c:manualLayout/c:x`` + ``c:y``, or ``None`` when
  the title is laid out automatically (the PowerPoint default).
  Assigning a 2-tuple writes ``c:layout/c:manualLayout`` with both
  ``c:xMode`` and ``c:yMode`` set to ``"factor"`` (the same element the
  PowerPoint UI writes when a user drags the title); assigning ``None``
  removes the ``c:layout`` entirely, restoring auto-layout.
- feat: #62 add read/write ``ColorFormat.alpha`` for per-color transparency.
  Exposes the OOXML ``<a:alpha val="N"/>`` child on any color-choice element
  (``a:srgbClr`` / ``a:schemeClr`` / ``a:sysClr`` / ``a:prstClr`` /
  ``a:hslClr`` / ``a:scrgbClr``) as a float in ``[0.0, 1.0]`` where ``1.0``
  is fully opaque and ``0.0`` is fully transparent — consistent with the
  existing ``brightness`` (float) and ``GradientStop.position`` (float in
  ``[0.0, 1.0]``) conventions. Reading an absent ``<a:alpha>`` returns
  ``1.0`` (the OOXML default). Assigning ``None`` removes any existing
  ``<a:alpha>`` so the color inherits opacity. Setting alpha on a
  ``MSO_COLOR_TYPE is None`` color raises ``ValueError`` (set ``.rgb`` or
  ``.theme_color`` first). The new property coexists with ``brightness``
  (``a:lumMod`` / ``a:lumOff``) on the same color element so transparent
  theme-color tints / shades round-trip correctly. Adds a regression
  suite ``DescribeIssue62FillAlpha`` under
  ``tests/test_issue_62_fill_alpha.py`` that round-trips alpha at
  ``[0.0, 0.25, 0.5, 0.75, 1.0]`` through ``Presentation.save`` + reopen
  on both an explicit RGB color and a theme color, and confirms
  coexistence with a brightness adjustment.
- feat: #438 add ``Presentation.save_ppsx(file)`` for saving a
  presentation as a PowerPoint Show. The serialized package is identical
  to a regular ``.pptx`` except for the content-type override on the
  presentation part, which is written as
  ``application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml``
  rather than ``…presentation.main+xml``. Named with a ``.ppsx``
  extension, the resulting file opens in PowerPoint directly into
  slide-show playback (no authoring UI). Complements #1070 (Wave 3),
  which added ``.ppsx`` read support by whitelisting
  ``CT.PML_SLIDESHOW_MAIN`` on load. Accepts the same ``zip_date_time``
  and ``password`` keywords as :meth:`Presentation.save`; the
  in-memory presentation's content-type is not mutated, so a follow-up
  ``save()`` writes a plain ``.pptx``.
- feat: #243 add :meth:`Chart.apply_template` that applies a ``.crtx``
  chart-template package to an existing chart. The template (a ZIP
  package containing a ``c:chartSpace`` XML part) is opened and its
  formatting elements are copied onto the target chart, leaving the
  chart's own series data, embedded workbook, and axis ``axId``
  references untouched. What crosses over: the chart-style index
  (``c:style`` or ``mc:AlternateContent`` / ``c14:style``),
  chart-space ``c:spPr`` / ``c:txPr``, the ``c:legend`` element, and
  (per axis of matching type) ``c:majorGridlines`` /
  ``c:minorGridlines`` / ``c:numFmt`` / ``c:majorTickMark`` /
  ``c:minorTickMark`` / ``c:tickLblPos`` / ``c:spPr`` / ``c:txPr``.
  ``c:ser`` data and ``c:externalData`` are preserved; ``c:title``
  from the template is adopted only when the target has no title.
  Accepts a path, bytes, or any file-like object
  :class:`zipfile.ZipFile` can open. Does not change chart type;
  pair with ``SlideShapes.add_chart`` when the template is for a
  different plot family than the target.
- feat: #575 add non-placeholder shapes to slide master / slide layout.
  :class:`MasterShapes` and :class:`LayoutShapes` now subclass
  ``_BaseGroupShapes`` and therefore expose the same shape-creation API as
  :class:`SlideShapes` — ``add_shape``, ``add_picture``, ``add_textbox``,
  ``add_connector``, ``add_group_shape``, and ``build_freeform``. This
  lets callers bake branding elements (client logo, background image,
  company watermark, fixed text) into a master or layout so they appear
  on every slide that inherits from it, without having to edit each
  slide individually. Shape-creation APIs that require chart / OLE
  part-creation helpers that currently live only on ``SlidePart``
  (``add_chart``, ``add_ole_object``, ``add_movie``) remain
  slide-only for now.
- fix: #991 normalize XML-declaration quoting to double quotes. lxml's
  ``etree.tostring`` emits the XML declaration with single quotes
  (``<?xml version='1.0' encoding='UTF-8' standalone='yes'?>``) while
  element attributes use double quotes, producing mixed quoting that
  trips some strict validators. ``serialize_part_xml``,
  ``oxml_to_encoded_bytes``, ``oxml_tostring``, and
  ``FlatOpcWriter.xml_bytes`` now rewrite the declaration to use double
  quotes so the entire saved package has consistent quoting matching
  Microsoft Office's own output.
- feat: #934 add ``Presentation.merge(other_presentation)`` for
  full-fidelity deck merging. Every slide in ``other_presentation`` is
  appended to the receiver via
  :meth:`Slides.add_slide_from_external`, now promoted from the
  restricted (#1036 "basic") copy path to full-fidelity cloning via
  Foundation F1 (``PartRelationshipCloner``) + F5
  (``clone_embedded_xlsx``). Charts come across with a *distinct*
  :class:`EmbeddedXlsxPart` so PowerPoint's "Edit Data" dialog keeps
  working on both the original and the merged copy; image and media
  parts are content-deduplicated against the target package;
  OLE-object and embedded-package rels get a shallow-clone fallback;
  external hyperlinks are preserved verbatim. Each cloned slide is
  bound to the layout at the same index in the target master's layout
  list as the source slide's layout occupied in the source master
  (with "last layout" fallback when the target has fewer layouts).
  Notes-slide relationships -- which carry a back-reference to their
  owning slide -- are intentionally dropped on the copies. See
  :meth:`Presentation.merge` for the convenience API and
  :meth:`Slides.add_slide_from_external` for per-slide control.
- verify: #449 (``insert_chart`` on Content Placeholder) resolved by
  ``feat/issue-333-content-placeholders`` (Wave 3) on top of
  ``feat/issue-199-placeholder-insert-chart`` (Wave 2). #199 lifted
  ``insert_chart`` from ``ChartPlaceholder`` onto
  ``_BaseSlidePlaceholder`` so every slide placeholder inherits it;
  #333 extended the same treatment to ``insert_picture`` and
  ``insert_table`` and updated ``SlidePlaceholder`` to advertise the
  full chart/picture/table insertion API. A Content placeholder has
  ``ph_type == PP_PLACEHOLDER.OBJECT`` and is not mapped to a
  specialized subclass by ``_SlidePlaceholderFactory``, so it lands on
  generic ``SlidePlaceholder`` and picks up the inherited
  ``insert_chart``. Adds a regression suite
  ``DescribeIssue449ContentPlaceholderInsertChart`` under
  ``tests/test_issue_449_content_placeholder_insert_chart.py`` that
  pins the factory mapping, the happy path on a "Title and Content"
  layout, a save-and-reload round trip, and coverage across every
  default-template layout that carries a Content placeholder
  (``Title and Content``, ``Two Content``, ``Comparison``,
  ``Content with Caption``).
- docs: #398 Jinja2 / templating of text placeholders — close as out-of-scope
  (templating is the caller's concern; python-pptx provides the plumbing via
  :meth:`.TextFrame.replace_text` / :meth:`._Paragraph.replace_text` shipped
  by #836). Add a ``Templating text`` section to ``docs/user/text.rst`` that
  documents the recommended read-render-write pattern with concrete Jinja2
  examples (whole-deck rendering paragraph-by-paragraph and per-token
  replacement), calls out the ``replace_text`` scoping rules (no crossing
  of paragraph / ``a:br`` / ``a:fld`` boundaries, origin-run formatting
  wins), and cross-references third-party libraries built on top of
  python-pptx (``python-pptx-templater``, ``pptx-template``,
  ``python-pptx-interface``) for callers who need looped slides,
  conditional inclusion, or chart-data substitution.
- verify: #275 resolved by F2 + #446 + #130 + #705
  (``feat/issue-275-shadows-verify``). The #275 reporter asked whether
  autoshapes inserted via ``slide.shapes.add_shape(MSO_SHAPE.…)`` had any
  option to "control the shadow" — at the time, ``BaseShape.shadow`` was a
  skeletal ``ShadowFormat`` exposing only ``.inherit``. Foundation F2
  (Wave 1) added the full ``a:effectLst`` family with read/write
  ``blur_radius`` / ``distance`` / ``direction`` / ``color`` on
  ``pptx.dml.effect.ShadowFormat``; #446 (Wave 2) tightened
  ``inherit = False`` to also zero any sibling
  ``p:style/a:effectRef/@idx`` so the explicit (empty) ``a:effectLst``
  really does suppress the theme-inherited shadow a caller would otherwise
  see on an MSO_SHAPE autoshape; #130 (Wave 3) lifted the same
  ``ShadowFormat`` onto ``ChartFormat.shadow``; and #705 (Wave 7) pinned
  the end-to-end API on ``Shape.shadow`` / ``GroupShape.shadow``. Adds a
  regression suite ``DescribeIssue275AutoshapeShadows`` under
  ``tests/test_issue_275_autoshape_shadows.py`` that walks the #275
  reporter's exact code path (``add_shape(MSO_SHAPE.…)`` → set blur /
  distance / direction / color) and round-trips every knob through
  ``Presentation.save`` + reopen.

- verify: #285 (change text without editing the formatting) resolved by
  ``feat/issue-836-replace-text-across-runs`` (Wave 6). The reporter
  asked for a way to rewrite the text of a placeholder / token without
  having to re-apply every run-level formatting attribute afterwards.
  :meth:`.TextFrame.replace_text` and :meth:`._Paragraph.replace_text`
  now do exactly that: the run where the match starts keeps its full
  ``a:rPr`` (bold, italic, underline, font size, font family, color,
  highlight, …) and absorbs the replacement text, runs fully inside the
  match are dropped, and any surviving suffix on a trailing run keeps
  that run's own rPr. The cross-run case PowerPoint produces whenever
  it re-flows a token after an edit (``"{NA"`` / ``"ME}"`` across two
  ``a:r`` elements with different formatting) is handled too — the
  origin run's rPr wins and the #285 caller no longer has to reach into
  runs and reapply styles. Adds a regression suite
  ``DescribeIssue285ReplaceTextPreservesFormatting`` under
  ``tests/test_issue_285_replace_text_preserve_format.py`` that
  round-trips bold / italic / font-size / font-name / RGB color through
  ``Presentation.save`` + reopen for the one-run, cross-run, and
  ends-mid-run cases.
- verify: #705 resolved by F2 + #130. Foundation F2 (Wave 1) shipped the
  full ``a:effectLst`` family on ``pptx.dml.effect`` — ``ShadowFormat``
  now exposes read/write ``blur_radius`` / ``distance`` / ``direction``
  / ``color`` (the four knobs the #705 reporter asked for) on top of
  the pre-existing ``.inherit`` stub, and #130 (Wave 3) lifted the same
  ``ShadowFormat`` onto ``ChartFormat.shadow``. Issue #446 (Wave 2)
  tightened the ``inherit = False`` semantics to also zero any sibling
  ``p:style/a:effectRef/@idx`` so an empty ``a:effectLst`` really does
  suppress the theme-inherited shadow. Adds a regression suite
  ``DescribeIssue705ShapeShadows`` under
  ``tests/test_issue_705_shape_shadows.py`` that pins the full API
  (blur / distance / direction / color on both ``Shape.shadow`` and
  ``GroupShape.shadow``) and round-trips every knob through
  ``Presentation.save`` + reopen.
- verify: #808 (replaced chart data triggers PowerPoint "repair needed" on
  open) resolved by ``fix/issue-396-replace-data-type-mismatch`` together
  with ``fix/chart-replace-data-490`` (both Wave 3). The two root-causes
  that put callers in the #808 state — passing a ``CategoryChartData`` to
  an XY or bubble chart (or vice-versa), which wrote ``c:cat/c:val`` into
  a plot expecting ``c:xVal/c:yVal`` / ``c:bubbleSize``, and a chart
  whose embedded-workbook relationship was stripped so
  ``ChartWorkbook.xlsx_part`` raised ``KeyError`` — are both handled up
  front now (``ValueError`` with a name-the-expected-type message for
  #396, transparent re-attach of a fresh ``EmbeddedXlsxPart`` for #490).
  The reporter's exact scenario (2-series 100% stacked-bar chart,
  replaced with 6 series and 5 categories via ``CategoryChartData``)
  produces valid category-chart XML with unique ``c:idx`` / ``c:order``,
  matched ``c:numRef`` / ``c:strRef`` caches, and a live
  ``c:externalData`` pointing at a freshly-written embedded ``.xlsx``;
  round-trips through save + reload with all 6 series intact. Adds a
  regression suite ``DescribeIssue808ReplaceChartDataRepair`` under
  ``tests/test_issue_808_replace_chart_data_repair.py`` that pins the
  happy path, the #396 type-mismatch guard, and the #490 missing-rel
  recovery so any future regression of either path would reproduce #808.
- fix: #767 strip file extension from movie shape name to match PowerPoint.
  ``SlideShapes.add_movie()`` previously named the newly-inserted ``p:pic`` shape
  with the full movie filename including its extension (e.g. ``"intro.mp4"``),
  so the selection pane showed ``"intro.mp4"``. PowerPoint itself strips the
  extension when inserting a video, producing the stem (``"intro"``).
  ``_MoviePicElementCreator._shape_name`` now returns
  ``os.path.splitext(self._video.filename)[0]`` so the shape name matches
  PowerPoint's convention.
- fix: #434 ``Picture.image`` raised ``InvalidXmlError`` when the backing
  ``p:pic`` element had no ``p:blipFill`` child (a malformed-but-observed
  shape produced by tools that strip image data). ``CT_Picture.blipFill``
  is now ``ZeroOrOne`` and every dependent accessor (``blip_rId``,
  ``_srcRect_x``, ``Picture.image``) handles the missing child safely —
  ``.image`` returns ``None``, crop accessors return ``0.0``, and shape-
  factory construction succeeds unchanged. Adds a regression suite at
  ``tests/test_issue_434_picture_missing_blipfill.py`` covering direct
  element access and a save-then-reopen round-trip.
- fix: #619 add data labels to XY Scatter chart.
  ``plot.has_data_labels = True`` previously raised ``AttributeError:
  'CT_ScatterChart' object has no attribute 'dLbls'`` because the
  oxml class was missing its ``dLbls`` ``ZeroOrOne`` descriptor — the
  slot was reserved in ``_tag_seq`` but never wired up. Adds the
  descriptor so scatter plots use the same ``c:dLbls`` subtree as the
  other Cartesian plot types (same default-off show-flags emitted by
  ``CT_DLbls.new_dLbls``; PowerPoint picks sensible per-plot-type
  defaults from those). Adds ``tests/test_issue_619_xy_scatter_data_
  labels.py`` covering all five ``XL_CHART_TYPE.XY_SCATTER_*`` variants
  and a save/reload round-trip.
- docs: #504 negative bars render white when a point fill is set. This is
  expected PowerPoint behavior: the point-level ``c:invertIfNegative``
  element defaults to |True|, so an authored solid fill on a negative bar
  is *inverted* to white on render. The workaround already landed as the
  #776 setter (``Point.invert_if_negative = False``); this change expands
  the ``Point.invert_if_negative`` docstring with the recipe (set the
  fill, then assign ``invert_if_negative = False`` on the same point) and
  adds a regression test (``tests/chart/test_point.py::DescribePoint::
  it_can_preserve_a_solid_fill_color_on_negative_bars_issue_504``) that
  asserts the combined XML carries both ``c:invertIfNegative val="0"`` and
  the authored ``a:srgbClr``. No code change is required.
- fix: #666 ``Chart.replace_data`` preserves author-set ``c:formatCode`` on
  ``c:val`` / ``c:xVal`` / ``c:yVal`` / ``c:bubbleSize`` and numeric ``c:cat``
  elements instead of silently resetting every series's number format back to
  ``"General"``. An explicit ``number_format=`` on the replacement
  ``CategoryChartData`` / ``XyChartData`` / ``BubbleChartData`` still wins —
  the default ``"General"`` is now treated as "no opinion" so the existing
  format survives. Implemented in
  ``_BubbleSeriesXmlRewriter._rewrite_ser_data`` /
  ``_CategorySeriesXmlRewriter._rewrite_ser_data`` /
  ``_XySeriesXmlRewriter._rewrite_ser_data`` (see
  ``src/pptx/chart/xmlwriter.py``) which capture the existing
  ``c:numCache/c:formatCode`` before the element is removed and reapply it to
  the freshly generated replacement. Regression coverage in
  ``tests/test_issue_666_replace_data_preserve_format.py``.
- feat: #309 look up slide shapes by name. Adds
  ``SlideShapes.get_by_name(name)`` which returns the first shape in the
  slide whose ``@name`` matches (or |None| when none do) and
  ``SlideShapes.find_all_by_name(name)`` which returns every matching
  shape in z-order. An explicit method pair was chosen over overloading
  ``__getitem__`` with a string key to keep indexed access unambiguous.
- feat: #133 add ``TextFrame.rotation`` read/write float property for the
  ``a:bodyPr/@rot`` attribute. Rotates the text *inside* the text frame in
  degrees clockwise (distinct from ``Shape.rotation`` which rotates the whole
  shape via ``p:spPr/a:xfrm/@rot``). Returns ``0.0`` when the attribute is
  absent; negative assignments are normalized to the equivalent
  ``[0, 360)`` value. Values are stored by PowerPoint in 60000ths of a
  degree; the existing ``ST_Angle`` converter handles that translation.
- feat: #547 expose ``BaseShape.flip_horizontal`` / ``BaseShape.flip_vertical``
  as read/write bool properties mapping to ``a:xfrm/@flipH`` / ``@flipV``, and
  add ``BaseShape.flip_horizontally()`` / ``BaseShape.flip_vertically()``
  convenience methods that toggle the corresponding attribute for parity with
  the PowerPoint ``Flip Horizontal`` / ``Flip Vertical`` UI commands. Setting
  a flip on a shape that has no ``a:xfrm`` creates one as needed; clearing a
  flip attribute reverts the element to schema-default (i.e. ``flipH=False``
  is represented by omitting the attribute).
- feat: #266 emit valid chart XML for 3D chart types
  (``THREE_D_AREA``/``_STACKED``/``_STACKED_100``,
  ``THREE_D_BAR_CLUSTERED``/``_STACKED``/``_STACKED_100``,
  ``THREE_D_COLUMN``/``_CLUSTERED``/``_STACKED``/``_STACKED_100``,
  ``THREE_D_LINE``, ``THREE_D_PIE``/``_EXPLODED``). The
  ``ChartXmlWriter`` factory previously raised ``NotImplementedError`` for
  these enum members; it now dispatches to new
  ``_Area3DChartXmlWriter`` / ``_Bar3DChartXmlWriter`` /
  ``_Line3DChartXmlWriter`` / ``_Pie3DChartXmlWriter`` builders that emit
  the ECMA-376 ``c:area3DChart`` / ``c:bar3DChart`` / ``c:line3DChart`` /
  ``c:pie3DChart`` wrappers plus a ``c:view3D`` sibling on ``c:chart``
  (PowerPoint's defaults: rotX=15, rotY=20, rAngAx=1, depthPercent=100).
  MVP — per-chart rotation/perspective customization deferred to a
  follow-up.
- feat: #144 add ``_Run.delete()`` and ``_Paragraph.delete()`` to remove a
  single run (``a:r``) from its paragraph and a single paragraph (``a:p``)
  from its text frame. The paragraph variant preserves the PowerPoint
  invariant that every text frame contains at least one ``<a:p>`` — when the
  paragraph being deleted is the last one, a fresh empty ``<a:p/>`` is added
  in its place. Both the ``p:txBody`` (shape) and ``a:txBody`` (table-cell)
  forms are handled. Subsequent use of a deleted ``_Run`` or ``_Paragraph``
  object is undefined.
- feat: #528 add ``_Paragraph.add_math_equation(omml_xml)`` so callers can
  insert an OMML equation (typically the output of Microsoft's
  ``MML2OMML.XSL``) directly into a text-frame paragraph. The fragment is
  wrapped in the ``mc:AlternateContent/mc:Choice[Requires="a14"]/a14:m``
  scaffolding PowerPoint emits for an inline equation, accompanied by an
  ``mc:Fallback/a:r`` run carrying the OMML reduced to its visible text
  (concatenated ``m:t`` children) so pre-2010 consumers render something
  readable. Existing runs, line-breaks and fields in the paragraph are
  preserved -- the equation is appended before any ``a:endParaRPr``. The
  companion read side is ``BaseShape.math_equation_xml`` /
  ``has_math_equation`` (#126). Converting between OMML and LaTeX /
  MathML remains out of scope -- the caller is responsible for producing
  the OMML. Registers the ``a14`` namespace (``http://schemas.microsoft.com
  /office/drawing/2010/main``) in ``pptx.oxml.ns``.
- feat: #234 ``FillFormat.blip_fill(image_file)`` — picture (image) fill.
  Adds a new ``blip_fill(image_file)`` method to |FillFormat| that embeds
  *image_file* as an ``ImagePart`` on the containing part (reusing an
  existing image part when the bytes match) and rewrites the underlying
  ``EG_FillProperties`` as ``<a:blipFill><a:blip r:embed="…"/>
  <a:stretch><a:fillRect/></a:stretch></a:blipFill>``. Works for auto-shape
  fills, table-cell fills, slide background fills, font fills, and line
  fills — any |FillFormat| whose factory call passes through a part
  reference. Raises |ValueError| when called on a |FillFormat| created
  without a part context (e.g. ``ChartFormat.fill``). Adds an end-to-end
  regression suite ``DescribeIssue234BlipFill`` under
  ``tests/test_issue_234_blip_fill.py`` that authors a rectangle shape,
  applies ``shape.fill.blip_fill("tests/test_files/python-powered.png")``,
  round-trips the presentation through save + reopen, and asserts both
  the ``MSO_FILL.PICTURE`` fill type and the embedded image bytes survive.
- feat: #259 custom document properties. Adds
  ``Presentation.custom_properties`` as a dict-like accessor for the
  ``/docProps/custom.xml`` OPC part (the "Custom" tab of PowerPoint's
  Document-Properties dialog and the store a ``{ DOCPROPERTY }`` field
  code reads from). Supports get / set / delete / ``__contains__`` /
  ``__iter__`` / ``keys`` / ``items`` / ``values`` / ``update`` /
  ``clear`` / ``pop`` / ``setdefault``; value types ``str`` /
  ``int`` (32-bit signed) / ``float`` / ``bool`` / ``datetime.datetime``
  are serialized as ``vt:lpwstr`` / ``vt:i4`` / ``vt:r8`` / ``vt:bool``
  / ``vt:filetime`` respectively. New ``CustomPropertiesPart``
  (``src/pptx/parts/customprops.py``) and ``CT_CustomProperties``
  (``src/pptx/oxml/custprops.py``) mirror the Wave-2 #131
  ``ExtendedPropertiesPart`` pattern. Reads against a package with no
  custom-properties part are no-ops; writes create the part lazily so a
  presentation that doesn't use custom properties round-trips without
  gaining a stray ``docProps/custom.xml`` entry. Regression suite in
  ``tests/test_issue_259_custom_properties.py`` plus unit coverage in
  ``tests/oxml/test_custprops.py`` and ``tests/parts/test_customprops.py``.
- docs: #244 enumerate every ``XL_CHART_TYPE`` member in ``docs/user/charts.rst``
  with its current support level (create / read / round-trip / not yet),
  including the ``UNSUPPORTED_CHARTEX`` sentinel introduced by #386 and a
  reference to that issue's chartex roadmap.
- #479/#651 waterfall chart: design analysis (F4-pending). Adds
  ``docs/dev/analysis/chartex-waterfall.rst`` covering the
  waterfall-specific ``cx:series`` shape with ``layoutId="waterfall"``,
  the ``cx:layoutPr/cx:subtotals/cx:subtotal`` block that marks bars as
  subtotal anchors, and the ``cx:visibility/@connectorLines`` toggle for
  the step-lines between bars. Documents the pair of ``cx:axis``
  children waterfall requires (unlike funnel) and proposes a minimal
  authoring-API surface (``CategoryChartData.subtotal_indices``).
  Documents the issue as blocked on full F4 (chartex foundation) with a
  concrete unblock-and-ship checklist; #479 and #651 are the same ask
  and will be resolved together.
- #371 treemap chart: design analysis (F4-pending). Adds
  ``docs/dev/analysis/chartex-treemap.rst`` covering the
  treemap-specific ``cx:series`` shape with ``layoutId="treemap"``, the
  hierarchical ``cx:chartData`` shape (``cx:strDim type="cat"`` with
  multiple leaf-first ``cx:lvl`` children encoding parent-child tile
  nesting), and the ``cx:layoutPr/cx:parentLabelLayout`` attribute that
  selects how PowerPoint paints labels for non-leaf tiles
  (``banner`` / ``overlapping`` / ``none``). Proposes tuple-valued
  ``CategoryChartData.categories`` (root-first, uniform depth) as the
  MVP authoring-API surface, with ragged hierarchies and per-tile
  overrides deferred to follow-up work. Documents the issue as blocked
  on full F4 (chartex foundation) with a concrete unblock-and-ship
  checklist; notes that the resulting writer is ~80% of the sibling
  sunburst kind.
- #305 funnel chart: design analysis (F4-pending). Adds
  ``docs/dev/analysis/chartex-funnel.rst`` covering the ``cx:plotArea``
  subset, ``cx:series`` with ``layoutId="funnel"``, and the cached-data
  structure a funnel chart writer must emit. Documents the issue as
  blocked on full F4 (chartex foundation) with a concrete
  unblock-and-ship checklist.
- docs: #407 rewrite ``Chart.chart_style`` docstring to reflect the two-tier
  reality (plain 1-48 via ``c:style`` vs. extended 49-255 via
  ``c14:style`` wrapped in ``mc:AlternateContent``) shipped with #516. Adds
  examples of common PowerPoint-UI-displayed plain styles (style 6 = the
  Accent-5 monochrome variant, etc.) and a note that the numeric-to-visual
  mapping depends on the theme applied (the six ``a:accentN`` colours from
  ``theme1.xml``) as well as the PowerPoint version, so the reliable
  workflow for hitting a specific visual is to set the style in PowerPoint
  and read back ``Chart.chart_style`` to learn the integer to use in code.
- docs: #584 Save as PDF — close as out-of-scope (rendering requires a layout
  engine / font rasterizer python-pptx does not ship; same disposition as
  #1049 PPT → MP4). Extend the "Rendering to video, PDF, or image formats"
  section of the user guide to call out PDF export explicitly and point at
  the same integration paths (``libreoffice --headless --convert-to pdf``,
  PowerPoint COM automation, Aspose.Slides, python-pptx-interface) already
  documented for MP4 / PNG rendering, and add a
  ``rendering-to-pdf-video-or-image-formats`` cross-reference label so other
  sections can link here.
- verify: #539 (bar chart color duplicated when extending series count)
  resolved by ``feat/issue-529-chart-theme-colors``. The #529 fix
  (``_apply_accent_color_to_ser`` called from
  ``_BaseSeriesXmlRewriter._add_cloned_sers``) rewrites any
  ``a:srgbClr`` / ``a:schemeClr`` fill on a cloned ``c:ser`` to
  ``a:schemeClr val="accent{n}"`` cycling 1..6 on the new series's
  index, so growing a coloured chart from 1 to 6+ series via
  ``Chart.replace_data()`` no longer repeats the source colour on every
  new series — the #539 scenario exactly. Adds a regression suite
  ``DescribeIssue539BarChartColorCycle`` under
  ``tests/test_issue_539_chart_color_cycle.py`` that replays the
  reporter's code (paint the source series red, grow to six series)
  and pins the accent-cycling behaviour across an 8-series wrap case
  and a save-and-reopen round-trip.
- verify: #810 (chart legend colours not matching graph colours)
  resolved by ``feat/issue-529-chart-theme-colors``. The #810 reporter
  described the same structural bug as #539 from a different angle —
  after ``Chart.replace_data()`` cloned a coloured source series onto
  new series the duplicated plot-area fills no longer lined up with
  the legend swatches (which are drawn from each ``c:ser/c:spPr``
  too, so cloned bars and legend entries ended up showing unrelated
  colours once the chart grew). The #529 fix's
  ``_apply_accent_color_to_ser`` guarantees each cloned ``c:ser`` owns
  a single, distinct ``a:schemeClr val="accent{n}"`` fill, and since
  PowerPoint renders both the legend marker and the plot-area bar for
  a given series from that same ``c:spPr``, legend colours now match
  graph colours by construction. Adds a regression suite
  ``DescribeIssue810LegendColorsMatchGraph`` under
  ``tests/test_issue_810_legend_colors.py`` that pins three invariants:
  every series gets a distinct colour key after a 1-to-6 growth, each
  series has exactly one ``c:spPr/a:solidFill`` (the single source of
  truth feeding both legend and plot) with no ``c:legendEntry`` colour
  override, and the one-colour-per-series invariant survives
  ``Presentation.save`` + reopen.
- verify: #777 resolved by ``feat/issue-752-ole-embed-generic``. Embedding
  an HTML file as an OLE object now works via the generic
  ``add_ole_object(html_path, prog_id="MSHtml.MHT", ..., extension="html")``
  path delivered by #752 (or via the ``PROG_ID.HTML`` convenience member
  which uses the alternate ``"htmlfile"`` progId). Both variants
  round-trip through save + reopen with the HTML bytes preserved
  byte-for-byte and the embedded part written under
  ``/ppt/embeddings/oleObject*.html`` with the generic OLE content-type.
  Adds an end-to-end regression suite
  ``DescribeIssue777EmbedHtmlOleObject`` under
  ``tests/test_issue_777_html_ole_embed.py`` that pins the flow for both
  a str path and a ``BytesIO`` caller.
- verify: #640 (duplicate a chart) resolved by ``feat/issue-877-cross-slide-
  chart-copy``. Same-slide duplication is a structural subset of the
  cross-slide chart-copy primitive: calling
  ``chart.clone_to(same_slide.shapes, x, y, cx, cy)`` produces a working
  duplicate on the source slide with distinct ``ChartPart`` and
  ``EmbeddedXlsxPart``, a fresh shape id / name (no collision with the
  source graphic-frame), and survives a save + reload round-trip. Adds
  a regression suite ``DescribeIssue640RegressionChartDuplicate`` under
  ``tests/test_issue_640_chart_duplicate.py`` that pins this behaviour.
- verify: #1033 (set arrow type of LINE object) resolved by
  ``feat/issue-375-line-arrows`` (Wave 1) together with
  ``fix/issue-749-auto-shape-type-line`` (Wave 2). #375 introduced
  ``LineFormat.begin_arrow`` / ``LineFormat.end_arrow`` with read/write
  ``type`` / ``width`` / ``length`` sub-properties and the
  ``MSO_LINE_END_TYPE`` / ``MSO_LINE_END_WIDTH`` / ``MSO_LINE_END_LENGTH``
  enumerations; #749 added ``MSO_SHAPE.LINE`` so a straight-line
  auto-shape can be added via ``shapes.add_shape(MSO_SHAPE.LINE, ...)``.
  Adds a regression test (``tests/test_issue_1033_line_arrow.py``) that
  authors a LINE shape, sets ``line.end_arrow.type =
  MSO_LINE_END_TYPE.TRIANGLE`` (plus width / length), round-trips the
  presentation through save + reopen, and asserts every arrow attribute
  survives — covering the exact flow the #1033 reporter asked for.
- verify: #1017 resolved by ``feat/issue-946-connector-adjustments``. The
  ``Connector.adjustments`` collection shipped for #946 delivers the
  elbow-connector "bend" operation the #1017 reporter was asking for —
  ``connector.adjustments[0] = 0.25`` rewrites the ``a:gd`` child of
  ``a:avLst`` and round-trips through save / reopen. Adds an end-to-end
  regression suite ``DescribeIssue1017ElbowConnectorAdjust`` under
  ``tests/test_issue_1017_elbow_connector_adjust.py`` that exercises the
  reporter's flow against a saved package (bend, save, reopen, confirm
  the new ``a:gd[@fmla='val 25000']`` survived) and also covers
  reading a pre-authored ``a:avLst`` and partial assignment on a
  three-adjustment ``bentConnector5``.
- verify: #874 resolved by Foundation F3 (``mc:AlternateContent`` traversal)
  plus #126 (OMML equation read API). The original bug — a shape whose text
  contained ``"500-7,000 m^3  per day."`` was silently dropped from
  ``slide.shapes`` — was caused by PowerPoint wrapping the equation-bearing
  ``p:sp`` in an ``mc:AlternateContent`` envelope the pre-F3 shape iterator
  didn't know how to descend into. F3 teaches ``CT_GroupShape.iter_shape_elms``
  to walk into ``mc:Choice`` transparently (preserving ``mc:Fallback`` for
  round-trip) and #126 adds ``BaseShape.has_math_equation`` /
  ``math_equation_xml`` so callers can detect and read the embedded OMML.
  Adds a regression test ``DescribeIssue874EquationInShapeText`` under
  ``tests/test_issue_874_shape_list_equation.py`` that opens the reporter's
  exact ``Presentation2.pptx`` attachment, confirms all six shapes surface
  (including the one whose ``m^3`` superscript forced the wrapping), and
  pins the ``mc:Fallback`` round-trip invariant.
- feat: #246 ``BaseShape.replace_with(other_shape)`` — one-call swap of
  an existing shape with a just-added replacement. Copies this shape's
  ``left`` / ``top`` / ``width`` / ``height`` onto ``other_shape``, moves
  ``other_shape``'s XML into this shape's z-order slot, then deletes this
  shape (dispatching to subclass ``delete`` so a ``Picture`` still drops
  its image relationship). Combined with the #41 ``BaseShape.delete()``
  primitive and ``SlideShapes.add_picture``, this closes the #246 ask:
  "replace an image while preserving its position". Adds a regression
  suite ``DescribeIssue246RegressionShapeReplace`` in
  ``tests/test_issue_246_shape_replace.py`` that round-trips a picture
  replacement through save + reopen.
- fix: #679 ``SlideMaster.name`` now falls back to a positional name of the
  form ``"Master N"`` (1-based index in ``prs.slide_masters``) when the
  underlying ``p:cSld/@name`` is empty — PowerPoint-authored masters almost
  always leave that attribute blank, so ``prs.slide_masters[0].name`` no
  longer surprises callers with an empty string. A ``SlideMaster.name``
  setter is added that writes ``p:cSld/@name`` (round-tripping through
  save/load); assigning ``""`` or ``None`` clears the attribute and restores
  the positional fallback.
- feat: #472 add ``DateAxis.major_unit`` and ``DateAxis.minor_unit``
  read/write properties so callers can set the tick-spacing on a
  date-scaled category axis (e.g. ``date_axis.major_unit = 3`` together
  with a ``c:majorTimeUnit`` of ``months`` yields a major tick every
  three months). The properties mirror the existing ``ValueAxis``
  accessors: ``None`` removes the ``c:majorUnit`` / ``c:minorUnit``
  child (restoring PowerPoint's Auto behaviour) and assigning a
  numeric value adds or replaces the child.
- feat: #224 add ``Slide.find_shapes_by_xpath(xpath_expr)`` — evaluate an
  XPath expression against the slide's ``p:spTree`` and return matching
  elements as :class:`BaseShape` proxies (using the same
  :func:`SlideShapeFactory` that :attr:`Slide.shapes` uses). The standard
  Open-XML namespace map (``pptx.oxml.ns._nsmap``) is bound, so the usual
  prefixes (``p``, ``a``, ``r``, ``mc``, ``p14``, …) work without further
  setup. Matches that land on a child element (e.g. ``p:cNvPr`` or
  ``a:xfrm``) are resolved to the nearest shape ancestor so callers always
  get a shape proxy back; duplicates are deduplicated in XPath-result
  order. Addresses the "find shape by name" use case requested in the
  issue, e.g.
  ``slide.find_shapes_by_xpath(".//p:sp[p:nvSpPr/p:cNvPr/@name='Title 1']")``.

- feat: #806 add ``SlideShapes.add_picture_link(url, left, top, width=None,
  height=None)`` for inserting a picture shape that *links* to an external
  image URL instead of embedding its bytes. Creates an external relationship
  of type ``http://schemas.openxmlformats.org/officeDocument/2006/
  relationships/image`` with ``Target=url`` / ``TargetMode="External"`` on
  the slide part and emits ``<a:blip r:link="rIdX"/>`` in place of the usual
  ``r:embed``. No image bytes are read or stored in the package; PowerPoint
  fetches the URL at render time. When ``width`` / ``height`` are omitted
  they default to one inch each (no aspect-ratio computation is possible
  without inspecting the image bytes, so callers should supply explicit
  dimensions to match the target image). End-to-end coverage lives in
  ``tests/test_issue_806_linked_picture.py``.
- feat: #1059 save as Flat OPC (XML Presentation) single-file XML. Adds
  ``Presentation.save_flat_xml(path_or_stream)`` (with matching
  ``PresentationPart.save_flat_xml`` / ``OpcPackage.save_flat_xml`` on the
  lower layers) that serializes the entire package as one ECMA-376 Part 4
  ``<pkg:package>`` document: every part is emitted as a ``<pkg:part>``
  child with its content type carried inline, XML parts embedded inside
  ``<pkg:xmlData>`` and binary parts (images, fonts, OLE, media)
  base64-encoded inside ``<pkg:binaryData>``. The output is prefixed with
  the ``<?mso-application progid="PowerPoint.Show"?>`` processing
  instruction so PowerPoint opens the resulting ``.xml`` file in the same
  way as its native "Save As → XML Presentation" command. Implemented in a
  new ``pptx.opc.flat_opc`` module.
- feat: #834 (and #116) add ``Picture.replace_image(image_file)``. Swaps
  the embedded image on a picture shape while keeping position, size,
  rotation, cropping, masking shape, outline, and any other shape-level
  formatting intact. Internally: adds (or reuses) an image part via the
  existing ``SlidePart.get_or_add_image_part`` deduplication, rebinds
  ``p:pic/p:blipFill/a:blip/@r:embed`` to the new rId, and drops the
  previous image relationship so the old image part is
  garbage-collected on save when no longer referenced. Raises
  ``ValueError`` on a malformed ``p:pic`` that has no embedded image to
  replace. #116 was deferred from Wave 2 pending Foundation F1; F1
  landed on master and the minimal swap turned out not to need the
  cross-part cloner. End-to-end regression in
  ``tests/test_issue_834_picture_replace_image.py``.
- feat: #784 add ``Movie.replace_media(new_path_or_file, mime_type=None)``.
  Swaps the audio/video binary behind an existing media shape while
  preserving its position, size, poster frame, hyperlink, and
  ``p:timing`` entries — a direct parallel of ``Picture.replace_image``
  (#116). A new |MediaPart| is created for the replacement bytes and the
  shape's ``a:videoFile`` / ``a:audioFile`` ``@r:link`` and ``p14:media``
  ``@r:embed`` rIds are rewired at it; the old relationships are dropped
  when no longer referenced so the previous media part is eligible for
  garbage-collection on save. Introduces ``CT_Picture.media_video_rId``
  and ``CT_Picture.media_embed_rId`` read/write properties as the XML
  attachment points. The shape's media-element tag
  (``a:videoFile`` vs ``a:audioFile``) is preserved across the swap;
  switching modality (audio ↔ video) requires
  ``SlideShapes.add_movie`` + delete instead. Verified end-to-end by a
  new ``tests/test_issue_784_replace_audio.py`` regression suite that
  round-trips a rewritten audio shape through save + reload.
- feat: #560 customize colors of series data labels. Per-point font color
  (``series.points[i].data_label.font.color.rgb``), per-point fill / line
  (via ``data_label.format`` shipped in #716), and per-series font color
  (``series.data_labels.font.color.rgb``) all already worked; this change
  closes the last gap by exposing ``DataLabels.format`` (``ChartFormat``
  wrapping ``c:dLbls`` with ``.fill`` / ``.line`` / ``.shadow``) so
  authors can color every label on a series at once without dropping to
  oxml. Registers ``c:spPr`` on ``CT_DLbls`` so ``ChartFormat`` can add
  it in correct schema position between ``c:numFmt`` and ``c:txPr``.
  End-to-end round-trip regression suite lives in
  ``tests/test_issue_560_data_label_colors.py``.
- feat: #836 add ``TextFrame.replace_text(find, replace)`` and
  ``Paragraph.replace_text(find, replace)`` that match ``find`` against
  the flattened paragraph text so a keyword split across multiple
  ``a:r`` runs (e.g., PowerPoint broke ``"{NAME}"`` into two runs after
  an edit) is still replaced. The run containing the start of the match
  keeps its formatting and absorbs the replacement; runs fully inside
  the match are dropped, and when a match ends partway through a run,
  that run's surviving suffix keeps its own formatting. Matches do not
  cross ``a:br`` (line break) or ``a:fld`` (auto-refresh field)
  boundaries. Returns the number of replacements performed.
- feat: #938 add ``Font.effective_color`` read-only property that walks the
  run's style inheritance chain and returns the |RGBColor| PowerPoint would
  render — including cases where the run has no explicit color and its
  effective color comes from a paragraph ``a:defRPr``, a text body
  ``a:lstStyle``, or the slide master's ``p:txStyles`` (via the theme's
  ``a:clrScheme``). Builds on the ``ColorFormat.to_rgb()`` resolver (#420)
  and its tint/shade support (#308). Returns |None| when no color can be
  resolved (for example when the |Font| object is created without a
  part-aware parent, as is the case for chart ``a:defRPr`` text). Existing
  ``run.font.color.rgb`` behaviour (``AttributeError`` on an inherited-only
  run) is unchanged; use ``effective_color`` when you need the value
  directly.

- feat: #883 ``Presentation(pptx_format=...)`` accepts paper-size presets
  ``"letter"`` (US Letter, 10 x 7.5 in tagged ``letter``) and ``"a4"``
  (A4 landscape, 297 x 210 mm tagged ``A4``), in addition to the previous
  ``"4x3"`` / ``"16x9"`` aspect-ratio presets. A ``(cx, cy)`` tuple in EMU
  is also accepted for arbitrary slide sizes (written as ``p:sldSz/@type
  = "custom"``). The ECMA-376 ``p:sldSz/@type`` enumeration is now
  exposed on ``CT_SlideSize`` as an optional ``type`` attribute.
- feat: #895 add/delete a column on an existing table. Adds
  ``_ColumnCollection.add(width=None)`` which appends a new ``a:gridCol`` to
  the table's ``a:tblGrid`` and a new empty ``a:tc`` to every existing row
  (new column inherits width from the last existing column or defaults to
  914,400 EMU = 1 inch when the table has no columns). Adds
  ``_Column.delete()`` and ``_ColumnCollection.remove(column)`` as the
  delete counterpart: the target ``a:gridCol`` and the ``a:tc`` at the same
  column offset in every row are detached, and the containing graphic-frame
  width is recomputed. Mirrors the row-mutation API landed by #832 and #837.
- fix: #1085 ``GroupShape.duplicate()`` places the clone at the source's
  slide-relative rectangle even when the source is nested inside one or
  more enclosing groups. Overrides ``BaseShape.duplicate()`` on
  ``GroupShape`` to clone the whole ``p:grpSp`` subtree via F1
  ``PartRelationshipCloner`` (so inner picture / media / OLE
  relationships are re-materialised on the target slide part), reassigns
  every ``cNvPr/@id`` in the clone to fresh unique ids, and sets the new
  group's ``a:off``/``a:ext`` to the source's
  ``effective_left``/``effective_top``/``effective_width``/``effective_height``
  (which apply the #925 group-transform cascade). The clone's
  ``a:chOff``/``a:chExt`` child coord system is preserved from the source
  so inner shapes retain their local positions.
- fix: #974 ``Movie.delete()`` now drops the three media-related slide-part
  rels (``a:videoFile``/``a:audioFile``'s ``@r:link``, the ``p14:media``
  ``@r:embed``, and the poster-frame ``a:blip``'s ``@r:embed``) and removes
  the matching ``p:timing/...//p:video[p:cMediaNode/p:tgtEl/p:spTgt/@spid]``
  timing-tree entry that ``SlideShapes.add_movie`` adds. Previously
  ``shape.delete()`` on a movie left a dangling ``p:video`` targeting a
  now-missing shape id plus three orphan rels, producing "file is corrupt"
  errors when the saved deck was reopened in PowerPoint. The fix follows
  the Wave 1 ``Picture.delete()`` override pattern.
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
- feat: #861 read/write access to animation delays. Adds ``stCondLst`` /
  ``endCondLst`` ``ZeroOrOne`` descriptors on ``CT_TLCommonTimeNodeData``
  and a convenience ``CT_TLCommonTimeNodeData.delay`` property that
  reads/writes the ``@delay`` attribute of the first
  ``p:stCondLst/p:cond`` child as an ``int`` (milliseconds) or the
  sentinel string ``"indefinite"``. Introduces typed ``CT_TLTimeCondition``
  (``p:cond``) and ``CT_TLTimeConditionList`` (``p:stCondLst`` /
  ``p:endCondLst``) element classes. A user-facing ``AnimationEffect``
  proxy is deferred to whatever #102 / #1106 land; today's scope is the
  XML-layer primitives on the existing F8 foundation.
- #811 Start time for ``add_movie``: adds ``Movie.start_time`` (float
  seconds, or ``None`` for ``"indefinite"``) and ``Movie.start_condition``
  (``"onClick"`` / ``"withPrevious"`` / ``"afterPrevious"``) read/write
  accessors. Both are backed by the F8 timing subtree —
  ``CT_TLCommonTimeNodeData.stCondLst`` now surfaces the
  ``p:stCondLst/p:cond`` under a movie's ``p:video`` timing node, and
  ``CT_TLTimeConditionList`` / ``CT_TLTimeCondition`` are new typed
  element classes for ``p:stCondLst`` / ``p:endCondLst`` / ``p:cond``.
- #256 Adding Timings Programatically — MVP introspection surface. Adds
  ``Slide.animation_sequence``, a read-only tuple of ``AnimationEffect``
  describing each effect in the slide's main animation sequence (the
  ``p:seq`` with ``nodeType="mainSeq"`` under ``p:timing``). Each
  ``AnimationEffect`` exposes ``shape_id`` (``p:spTgt/@spid``),
  ``preset_class`` / ``preset_id`` / ``preset_subtype`` (the
  ``p:cTn/@presetClass`` / ``@presetID`` / ``@presetSubtype`` preset
  selectors), and ``delay`` (first ``p:cond/@delay``). Adds
  ``CT_TLShapeTargetElement`` for ``p:spTgt`` and the
  ``iter_main_sequence_effects`` / ``first_spTgt_spid`` helpers in
  ``pptx.oxml.timing``. Extends ``CT_TLCommonTimeNodeData`` with the
  three preset attributes. Authoring (add / remove / reorder effects,
  ``Shape.animation``) is deferred to downstream items #102 / #264 /
  #1106 which now have a stable read surface to layer on top of.
- #264 shape-animation introspection (read-only MVP). Adds
  ``Slide.iter_shape_animations()`` yielding a ``ShapeAnimation`` proxy
  for every shape-targeted effect (``p:anim`` / ``p:animEffect`` /
  ``p:animMotion`` / ``p:animRot`` / ``p:animScale`` / ``p:animClr`` /
  ``p:set``) in the slide's ``p:timing`` subtree. Each proxy exposes
  ``shape_id``, ``effect_type``, ``delay_ms`` (int ms /
  ``"indefinite"`` / ``None`` — directly answers the user's ask in
  #264), ``duration_ms``, and ``element`` for callers that need to
  hand-edit XML today. Write-side authoring (add / modify effects,
  motion paths, full trigger configuration) is deferred to the Wave-7
  lift that combines #102, #264, #861, #1106; see
  ``docs/dev/analysis/f8-animations-transitions.rst``.
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
- #102 shape animations (MVP). Adds ``Shape.animation`` read-only proxy
  and ``Shape.set_animation(effect_type, trigger='onClick', delay=0)``
  writer for the five common presets ``APPEAR``, ``FADE_IN``, ``FLY_IN``,
  ``PULSE``, and ``FADE_OUT`` via two new enums
  ``MSO_ANIMATION_TYPE`` / ``MSO_ANIMATION_TRIGGER`` in
  ``pptx.enum.animation``. Effects are authored into the slide's
  ``p:timing/p:tnLst/p:par[tmRoot]/p:cTn/p:childTnLst/p:seq[mainSeq]``
  sub-tree using the F8 foundation. ``trigger`` accepts
  ``'onClick'`` / ``'onPrev'`` string aliases or
  ``MSO_ANIMATION_TRIGGER.ON_CLICK`` / ``.AFTER_PREVIOUS``;
  ``delay`` is a non-negative int (milliseconds). Out of MVP scope and
  reserved for downstream issues: motion-paths (#264), broader preset
  sets and color-emphasis (#1106), ``withEffect`` /
  ``onMouseOver`` / interactive sequences (#264).
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
- #715 / #969 python-pptx fit text within text placeholder. ``TextFrame``
  gains read/write ``font_scale`` and ``line_space_reduction`` properties
  that expose the ``fontScale`` and ``lnSpcReduction`` attributes of
  ``a:normAutofit``. Assigning either attribute ensures an ``a:normAutofit``
  child is present on ``a:bodyPr``, replacing any existing ``a:noAutofit``
  or ``a:spAutoFit`` choice sibling. This lets callers pre-compute the
  autofit hints that PowerPoint reads on first render, so placeholder text
  appears correctly scaled without requiring the user to edit the box —
  important for viewers such as LibreOffice, Google Slides, and headless
  image renderers that honor the attributes literally rather than
  recomputing the fit. The existing ``TextFrame.fit_text()`` convenience
  method already works on placeholder text frames because placeholder
  width/height are inherited through the layout/master chain.
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
