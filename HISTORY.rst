.. :changelog:

Release History
---------------

Unreleased
++++++++++

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
