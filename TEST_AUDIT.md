# Test Suite Audit

This report surveys the state of the `loadfix/python-pptx` fork test suite
with three aims:

1. Document what is well covered and what is not.
2. Identify pre-existing latent defects and anti-patterns.
3. Propose concrete, prioritised follow-ups.

All tests were run against commit ``1a345f8f`` on ``master``. The baseline
is **4 791 passed, 0 skipped, 0 deselected** in ~21 s
(``PYTHONPATH=src pytest tests/ -W
ignore::pyparsing.warnings.PyparsingDeprecationWarning -q``).

Behave acceptance tests pass cleanly: **69 features / 1 208 scenarios /
3 810 steps passed, 0 failures** in ~6 s
(``PYTHONPATH=src behave --tags=-wip``).

---

## 1. Coverage summary

Run:

```
PYTHONPATH=src pytest tests/ --cov=src/pptx --cov-report=term-missing \
  -W ignore::pyparsing.warnings.PyparsingDeprecationWarning -p no:randomly -q
```

**Overall coverage: 97 %** (18 338 statements, 585 missed).

Counting tests, the suite comprises 439 ``Describe*`` classes across 141
test modules and ~48 961 lines of test code (``tests/`` tree, excluding
``__pycache__``). ``features/`` contributes 69 ``.feature`` files with
538 ``Scenario``/``Scenario Outline`` headers (1 208 scenarios when
outline rows are counted, as ``behave`` does above).

### 1.1 Lowest-coverage production modules

The five modules with the lowest coverage percentages and a terse note
about the missing lines.

| % | module | stmt / miss | uncovered (representative) |
|---:|---|---:|---|
| 0 | ``src/pptx/types.py`` | 9/9 | whole file — two runtime-irrelevant ``typing.Protocol`` classes (``ProvidesExtents``, ``ProvidesPart``). Every line is under ``if TYPE_CHECKING:`` or is a `...` stub; coverage is correctly not measured by executing the module, only by importing it (and the suite does not import ``pptx.types`` at runtime). |
| 72 | ``src/pptx/enum/base.py`` | 74/21 | ``DocsPageFormatter`` (lines 103-175) — RST doc-generation tool, exercised only by the ``docs/`` Sphinx build, never by tests. Identical to the python-docx situation; consider moving it into ``docs/`` or adding a smoke test. |
| 84 | ``src/pptx/chart/chart.py`` | 614/100 | two large dead regions. Lines 935-995 are the ``_as_number`` / ``_as_text`` / ``_set_cache_pt`` code paths for the single-cell ``_update_numRef``/``_update_strRef`` writer (used by the Wave-8 ``Chart._update_cell`` plumbing — happy path is exercised but no test feeds it a non-numeric value, boolean, missing ``c:ptCount``, or an index beyond the cache). Lines 1340-1479 are contiguous blocks inside the Wave-8 ``apply_template`` (#243) axis / legend / title copy-over — ``_insert_axis_child`` fall-through, ``_apply_legend`` on a chart without a ``c:plotArea``, ``_apply_title_if_absent`` when the target already has a title, and the ``_insert_in_schema_order`` ``ValueError`` and out-of-sequence branches. Worth filling given ``apply_template`` is a headline Wave-8 landing. |
| 86 | ``src/pptx/dml/effect.py`` | 205/29 | the ``= None`` setter branches on ``ShadowFormat.blur_radius`` / ``distance`` / ``direction`` (222-226, 244-248, 266-270) when ``_child`` is absent, plus the whole ``GlowFormat.size`` / ``color`` None-setter group (311, 343, 348-352, 362, 367-371). These exist for parity with ``brightness`` / ``alpha`` but are authored and never round-tripped through ``Presentation.save`` under test. |
| 91 | ``src/pptx/oxml/theme.py`` | 47/4 | absent-element branches in ``CT_OfficeStyleSheet.themeElements`` subtree walks (73, 80, 83, 86) — reached only from a theme part that is missing a ``<a:clrScheme>``. |

Counting statements, the highest-risk "dense red" modules by *range
size* rather than percentage are:

- ``src/pptx/chart/chart.py`` (84 %, 100 lines missed across two contiguous
  blocks — largest coverage hole in the repo).
- ``src/pptx/text/text.py`` (93 %, 45 missed): lines 721-734 and 757-772
  are the Wave-5 ``Font.effective_color`` (#938) master-``txStyles``
  fallback walk. The happy path is tested with a scheme color only;
  bodyStyle → otherStyle → titleStyle failure/fallback ordering is
  uncovered.
- ``src/pptx/slide.py`` (95 %, 36 missed): scattered but includes
  1452/1468/1484/1508/1514 (the animation-sequence read-only proxy on the
  *pptx.slide* module that was just renamed to ``AnimationEffectView`` —
  the new alias is exercised, the rebranded proxy's defensive ``is None``
  branches aren't) and 1817-1837 (``SlideTags.get`` / ``values`` /
  ``items`` on slides without a tags part).
- ``src/pptx/chart/xlsx.py`` (93 %, 35 missed): the ``xlsxwriter``
  rewrite-on-replace code in ``WorkbookWriter._write_cell`` handles
  float/int/bool/None well; uncovered branches are the error paths when
  the shared-strings table or the embedded workbook has unexpected
  shape — i.e. the "load a legacy chart workbook" recovery paths.

### 1.2 Modules at 92 % - 96 %

Several large modules sit in the 92-96 % band. The missing lines are
typically defensive None-returning branches and a few edge cases:

- ``src/pptx/oxml/shapes/shared.py`` (93 %, 20/284 missed): the anchor /
  placeholder ``CT_NvSpPr`` size-attribute read paths when the attribute
  is absent (109, 131, 142, 174, 418-453). Worth a couple of parametrised
  tests.
- ``src/pptx/oxml/comments.py`` (93 %, 9/129 missed): lines 160-161 and
  194-200 are the ``CT_CommentAuthorList`` lookup-by-id miss paths (#487
  Wave-? comments landing). Trivial to cover.
- ``src/pptx/oxml/chart/chart.py`` (93 %, 14/189 missed): defensive
  ``CT_ChartSpace._add_autoTitleDeleted`` branches reachable only from
  malformed chart XML.

### 1.3 Well-covered areas

Modules at 99-100 % include: ``pptx.animation`` (96 %), ``pptx.comments``
(99 %), ``pptx.chart.datalabel`` (100 %), ``pptx.chart.legend`` (100 %),
``pptx.chart.marker`` (100 %), ``pptx.chart.point`` (100 %),
``pptx.dml.fill`` (100 %), ``pptx.enum.chart`` (100 %),
``pptx.enum.dml`` (100 %), ``pptx.enum.lang`` (100 %),
``pptx.enum.shapes`` (100 %), ``pptx.enum.text`` (100 %),
``pptx.opc.constants`` (100 %), ``pptx.opc.flat_opc`` (100 %),
``pptx.opc.package`` (99 %), ``pptx.opc.packuri`` (100 %),
``pptx.opc.serialized`` (99 %), ``pptx.oxml.*`` generally (99-100 %),
``pptx.parts.chart`` / ``parts.comments`` / ``parts.coreprops`` /
``parts.extprops`` / ``parts.font`` / ``parts.media`` /
``parts.presentation`` / ``parts.tags`` (100 %),
``pptx.shapes.shapetree`` (99 %), ``pptx.table`` (98 %).

The recent Wave-8 landings (#319, #971, #425, #1030, #62, #438, #243,
#575) all ride on top of modules at ≥96 %.

---

## 2. Test-suite scale

### 2.1 ``tests/`` (pytest)

| metric | count |
|---|---:|
| ``*.py`` files | 141 |
| lines of code | 48 961 |
| ``Describe*`` classes | 439 |
| ``Test*`` classes | (not counted separately; the suite is ``Describe*``-dominant) |
| passing tests (baseline) | 4 791 |

### 2.2 ``features/`` (behave)

| metric | count |
|---|---:|
| ``*.feature`` files | 69 |
| ``Scenario`` / ``Scenario Outline`` headers | 538 |
| behave-counted scenarios (outline rows expanded) | 1 208 |
| steps (behave-counted, outline rows expanded) | 3 810 |
| step modules | 39 (under ``features/steps/``) |

Feature prefixes (grouped by subject — same convention as the step
modules): ``act-`` (actions/hyperlinks), ``cht-`` (charts, 22 files),
``cmt-`` (comments, 1 file — fork addition), ``dml-`` (DrawingML —
color/fill/line/effect), ``opc-`` (packaging), ``ph-`` (placeholders),
``prs-`` (presentation — 7 files including the fork ``prs-sections`` and
``prs-font-embedding``), ``shp-`` (shapes — 11 files including the fork
``shp-animation``), ``sld-`` (slides — includes fork ``sld-transition``
and ``sld-background``), ``tbl-`` (tables), ``txt-`` (text).

Fork-era additions that did ship a behave feature file:
``cmt-comments.feature`` (#487 legacy comments, 16 lines),
``cht-chartex.feature`` (#386 chartex, 30 lines),
``shp-animation.feature`` (#102 animation MVP, 36 lines),
``prs-font-embedding.feature`` (#355 font embedding, 19 lines),
``cht-clone-to.feature`` (#877 cross-slide chart copy, 25 lines),
``prs-sections.feature`` (Foundation F7 sections, 66 lines). These six
are the entire behave delta between upstream ``scanny/python-pptx`` and
the fork — see §5 for the full list of fork additions that never
received a behave scenario.

---

## 3. Latent defects / anti-patterns

### 3.1 Skipped / deselected tests

**None.** No ``@pytest.mark.skip`` / ``skipif``, no ``pytest.skip(...)``,
no ``--deselect`` in ``pyproject.toml`` or any Makefile / tox target.
Unlike python-docx (28 deselected ``CT_Border`` tests), this fork runs
100 % of its tests 100 % of the time. The suite is *healthier* than its
sibling in that respect.

### 3.2 ``pytest.raises`` / ``pytest.warns`` specificity

- 248 ``pytest.raises`` invocations, 0 ``pytest.warns``.
- 105 of the 248 include a ``match=`` regex (~42 %). Compare python-docx,
  which runs a similar ratio.
- ``pytest.raises(Exception)`` / ``pytest.raises(BaseException)`` count:
  **zero**. All invocations name a specific subclass (``ValueError`` /
  ``TypeError`` / ``KeyError`` / ``IndexError`` / ``NotImplementedError``
  / ``OverflowError`` / ``InvalidXmlError`` /
  ``UnsupportedImageTypeError`` / ``EncryptedPackageError`` /
  ``PackageNotFoundError`` / ``TextLayoutError`` / ``AttributeError`` /
  ``RuntimeError`` / ``PackageTooLargeError``).
- Uses of a *specific subclass without* ``match=`` are legitimate when the
  raised exception is unambiguous (e.g. ``pytest.raises(KeyError)`` after
  a missing-key dict lookup). The highest-count unmatched forms are
  ``ValueError`` (69), ``IndexError`` (19), ``TypeError`` (15),
  ``KeyError`` (13). Among those 69 ``ValueError``s, some wrap a real
  parser/validator where a ``match=`` would catch regressions (e.g.
  shape-geometry preset parsing, ``Emu``/``Pt``/``Inches`` unit parsing,
  hexadecimal color validation). Adding ``match=`` to those specific
  call sites would be a P3 tightening pass but is not urgent.

Overall: **no broad-exception anti-pattern.** The gap is that ~58 % of
raises tests could be more specific, not that any tests are over-broad.

### 3.3 Shared-state polluters

The python-docx sibling audit flagged a C1 polluter-fix that removed
22 ``_restore_part_factory`` workarounds. **This fork still carries the
same workaround in 12 test files**, totalling 49 ``_restore_part_factory``
usages. These are concentrated in the ``tests/test_issue_*.py``
regression suites authored during Waves 7-8:

```
tests/test_issue_62_fill_alpha.py                5
tests/test_issue_275_autoshape_shadows.py        2
tests/test_issue_285_replace_text_preserve_format.py  4
tests/test_issue_357_pie_chart_colors.py         6
tests/test_issue_438_save_as_ppsx.py             6
tests/test_issue_449_content_placeholder_insert_chart.py  5
tests/test_issue_571_catax_number_format.py     2
tests/test_issue_662_data_label_background.py    4
tests/test_issue_825_scatter_point_colors.py     2
tests/test_issue_937_table_cell_layout.py        4
tests/test_issue_1030_chart_title_position.py    6
tests/test_issue_1106_entrance_exit_verify.py    3
```

Additionally, ``tests/test_presentation.py::DescribeIssue694SectionsAPI``
carries an ``autouse`` ``_isolate_part_factory`` fixture with the same
shape (``tests/test_presentation.py:1006-1030``).

**Root cause.** The regression-tests exercise real ``.pptx`` round-trip
plumbing (``Presentation`` → ``save`` → ``Presentation``) which calls
``PartFactory(partname, content_type, package, blob)``. That lookup
reads ``PartFactory.part_type_for`` (module-level dict seeded in
``pptx/__init__.py`` at import time). Any test that **mutates the dict
and does not restore it** poisons every subsequent round-trip test in
the same pytest session.

The fork's defensive pattern is correct *locally*: each issue-regression
test saves / updates / restores. But the *right* fix is centralized —
either:

1. An autouse session/function-scoped ``conftest.py`` fixture that
   snapshots and restores ``PartFactory.part_type_for`` around every
   test. This would let all twelve files drop their ``_restore_part_factory``
   fixture entirely (≈160 lines of boilerplate).
2. Audit which existing test *is* the polluter and fix it at the source.
   Candidates (per git blame on ``tests/opc/test_package.py:638
   DescribePartFactory``, ``tests/test_package.py``, ``tests/parts/
   test_presentation.py::SlidePart_ class_mock``) all appear to use
   ``monkeypatch.setitem`` or module-level ``class_mock``, which do *not*
   mutate ``PartFactory.part_type_for``. It's possible the pollution is
   gone and the twelve guards are now unnecessary defense-in-depth. A
   thirty-minute experiment (delete every ``_restore_part_factory``
   usage, run the full suite in random order, see what fails) would
   answer definitively.

See §6 P1 #1.

No other shared-state polluters (``register_element_cls``, global
registries, ``os.environ`` mutation, ``sys.path`` mutation,
``logging.getLogger().handlers`` mutation) are in use.

### 3.4 Flakiness / random-order sensitivity

``pytest-randomly`` was installed and the suite was run three times with
different seeds — ``--randomly-dont-reset-seed``, ``--randomly-seed=42``,
``--randomly-seed=12345``:

```
4791 passed in 19.90s         (random seed A)
4791 passed in 20.48s         (random seed 42)
4791 passed in 20.47s         (random seed 12345)
```

**Result: no order dependencies observed**, despite the ``PartFactory``
polluter concern in §3.3. The twelve ``_restore_part_factory`` guards
evidently do their job. (This also corroborates the hypothesis that the
guards are over-defensive — no other tests seem to be leaking.)

Other flake sources:

- **Wall-clock dependency.** ``tests/parts/test_coreprops.py:150`` uses
  ``datetime.now`` to bracket a ``core_props.modified`` update and
  asserts the timestamp falls in the ``[before, after]`` window. Safe.
  No other ``datetime.now`` / ``time.sleep`` / ``perf_counter`` in
  ``tests/``.
- **Filesystem state.** ``tests/test_issue_777_html_ole_embed.py:147``
  uses ``tempfile.NamedTemporaryFile(delete=False)`` with a
  ``try:/finally: os.unlink`` cleanup — correct. No other direct
  ``tempfile`` uses in ``tests/``. Many tests use ``tmp_path`` or
  ``io.BytesIO`` round-trips, which are ideal.
- **Network.** Zero tests touch the network.
- **External tooling.** No LibreOffice / external-binary dependencies
  (which *would* be reasonable for round-trip validation — see §6 P3).

Flake risk: **negligible**. The only real concern is the
``_restore_part_factory`` defensive pattern.

### 3.5 Over-mocked / tautological tests

Spot-check of ~15 test modules found the suite is, on the whole, grounded
in behaviour via the ``cxml`` helper plus real-element fixtures (same
pattern as python-docx). A couple of borderline cases:

- **``tests/test_action.py:217-249``** — the ``screen_tip`` (#425 Wave-8)
  tests pin ``None``-when-absent and the happy-path getter/setter
  against a real ``a:hlinkClick`` element. Reasonable.
- **``tests/test_slide.py::DescribeSlide_TitleSlide``** — several "proxy
  forwards" tests mock ``slide_part_`` and assert the property returns
  the mock's return value. Low-value but not harmful.
- **``tests/parts/test_slide.py::DescribeSlidePart``** — a few tests
  mock four collaborators on ``_Relationships`` and assert specific
  calls. Brittle to internal refactoring; not a pattern to propagate.

No ``try/except: pass`` blocks found in ``tests/`` except two
intentional ones in
``tests/test_issue_777_html_ole_embed.py`` (cleanup) and the twelve
``_restore_part_factory`` ``try:/finally:`` blocks.

### 3.6 Outdated patterns

- **``unittest.TestCase``**: **none**. Only ``unittest.mock`` is
  imported, which is the pytest-idiomatic use.
- **``setUp`` / ``tearDown``**: **none**.
- **Stray test stubs**: **none** observed.
- **Commented-out tests**: **none** observed.

---

## 4. Fork-era test coverage check

This section cross-references ``HISTORY.rst`` landings against the
existence of tests (unit, regression, acceptance). The fork's
``CLAUDE.md §6`` mandates unit + acceptance + docs + HISTORY for every
user-visible change; §7 reiterates this as the "strict feature-work
rule". This check looks for landings that diverge from that rule.

The 51 ``feat:`` and 44 ``fix:`` entries in ``HISTORY.rst`` below the
``Unreleased`` heading were audited. The scoring heuristic:

- **unit** — is there a test in ``tests/**/test_*.py`` exercising the
  new API (searched by the distinctive symbol name, not by issue number)?
- **regression** — is there a ``tests/test_issue_<N>_*.py`` file?
- **behave** — is there a scenario in ``features/*.feature`` or a step
  in ``features/steps/*.py``?

### 4.1 Fork landings with full unit + regression + behave

Only **six** fork-era landings hit all three:
``#877`` (Chart.clone_to, ``features/cht-clone-to.feature``),
``#102`` (animation MVP, ``features/shp-animation.feature``),
``#355`` (embed_font, ``features/prs-font-embedding.feature``),
``#487`` (legacy comments, ``features/cmt-comments.feature``),
``#386`` (chartex, ``features/cht-chartex.feature``),
Foundation F7 (sections, ``features/prs-sections.feature``).

### 4.2 Fork landings with unit tests but no behave scenario

Most fork landings are in this bucket. ``HISTORY.rst`` entries list the
test file explicitly, but ``features/`` was not extended. The §6
"strict feature-work rule" item 3 was skipped (usually silently, without
the "N/A — why" note the rule asks for).

| # | area | unit test | behave? |
|---|---|---|---|
| 319 | ``Slide.is_hidden`` | ``tests/test_slide.py`` | no |
| 971 | ``BaseShape.is_hidden`` | ``tests/shapes/test_base.py`` | no |
| 425 | ``ActionSetting.screen_tip`` | ``tests/test_action.py`` | no |
| 1030 | ``ChartTitle.position`` | ``tests/test_issue_1030_*.py`` + ``tests/chart/test_chart.py`` | no |
| 62 | ``ColorFormat.alpha`` | ``tests/test_issue_62_*.py`` + ``tests/dml/test_color.py`` | no |
| 438 | ``Presentation.save_ppsx`` | ``tests/test_issue_438_*.py`` + ``tests/test_presentation.py`` | **no** |
| 243 | ``Chart.apply_template`` | ``tests/chart/test_chart.py`` | **no** |
| 991 | XML-decl quote normalization | ``tests/test_issue_991_*.py`` | no (N/A — internal) |
| 934 | ``Presentation.merge`` | ``tests/test_presentation.py`` | no |
| 309 | ``SlideShapes.by_name`` | ``tests/shapes/test_shapetree.py`` | no |
| 133 | ``TextFrame.rotation`` | ``tests/text/test_text.py`` | no |
| 547 | ``BaseShape.flip_*`` | ``tests/shapes/test_base.py`` | no |
| 144 | ``_Run.delete`` / ``_Paragraph.delete`` | ``tests/text/test_text.py`` | no |
| 528 | ``_Paragraph.add_math_equation`` | ``tests/text/test_text.py`` | no |
| 234 | ``FillFormat.blip_fill`` | ``tests/dml/test_fill.py`` + ``tests/test_issue_234_*.py`` | no |
| 259 | custom properties | ``tests/test_custom_properties.py`` + ``tests/test_issue_259_*.py`` | no |
| 246 | ``BaseShape.replace_with`` | ``tests/test_issue_246_*.py`` + ``tests/shapes/test_base.py`` | no |
| 472 | ``DateAxis.major_unit`` / ``minor_unit`` | ``tests/chart/test_axis.py`` | no |
| 224 | ``Slide.find_shapes_by_xpath`` | ``tests/test_slide.py`` | no |
| 806 | ``SlideShapes.add_picture_link`` | ``tests/test_issue_806_*.py`` | no |
| 1059 | save as Flat OPC | ``tests/opc/test_flat_opc.py`` | no |
| 834 | ``Picture.replace_image`` | ``tests/test_issue_834_*.py`` | no |
| 784 | ``Movie.replace_media`` | ``tests/test_issue_784_*.py`` | no |
| 560 | data-label colors | ``tests/test_issue_560_*.py`` + ``tests/chart/test_datalabel.py`` | no |
| 836 | ``TextFrame.replace_text`` | ``tests/text/test_text.py`` | no |
| 938 | ``Font.effective_color`` | ``tests/text/test_text.py`` | no |
| 883 | ``Presentation(pptx_format=...)`` | ``tests/test_presentation.py`` | no |
| 895 | add/delete table column | ``tests/test_table.py`` | yes (``tbl-table.feature:71-82``) |
| 1004 | ``Transition.speed`` | ``tests/test_slide.py`` | no |
| 942 | MORPH transition | ``tests/test_slide.py`` | no |
| 861 | animation delays | ``tests/test_animation.py`` | no |
| 832 | ``_RowCollection.add()`` | ``tests/test_table.py`` | yes (``tbl-table.feature:57-62``) |
| 837 | delete row | ``tests/test_table.py`` | yes (``tbl-table.feature:65-68``) |
| 49 | z-order | ``tests/shapes/test_base.py`` | yes (``shp-shared.feature:256``) |
| 544 | error bars | ``tests/chart/test_plot.py`` | no |
| 381 | refresh chart data | ``tests/chart/test_chart.py`` | no |
| 752 | OLE embed | ``tests/test_custom_xml.py`` + others | no |
| 199 | ``SlidePlaceholder.insert_chart`` | ``tests/shapes/test_placeholder.py`` | no |
| 333 | content placeholders | ``tests/shapes/test_placeholder.py`` | no |
| 1070 | read ``.potx`` / ``.ppsx`` | ``tests/test_presentation.py`` | no |
| 338 | combo charts | ``tests/chart/test_chart.py`` | no |
| 734 | sound action | ``tests/test_action.py`` | no |
| 27 | apply table styles | ``tests/test_table.py`` | no |
| 668 | chart lookup | ``tests/chart/test_chart.py`` | no |

The "**no**" entries in bold are *headline* user-facing features
(``save_ppsx``, ``apply_template``, add/delete table column/row,
z-order) where a behave scenario is the most natural
documentation-through-execution of the feature. These should be
prioritised (see §6 P2).

### 4.3 Fork landings with HISTORY entry but no regression test

A few recent landings ship only with mirror-path unit tests and no
``tests/test_issue_<N>_*.py`` regression file:

- #319 (``Slide.is_hidden``) — tested in ``tests/test_slide.py``
- #971 (``BaseShape.is_hidden``) — tested in ``tests/shapes/test_base.py``
- #425 (``ActionSetting.screen_tip``) — tested in ``tests/test_action.py``
- #243 (``Chart.apply_template``) — tested in ``tests/chart/test_chart.py``
- #575 (master/layout shapes) — tested in ``tests/test_slide.py``
- #934 (``Presentation.merge``) — tested in ``tests/test_presentation.py``

Pattern is fine (mirror-path is the canonical place per CLAUDE.md §5).
Just noted for completeness; not a defect.

### 4.4 Fork landings fully absent from behave

Every ``fix:`` entry is an "N/A for behave" case unless the fix
materially changes a user-visible API. The ``verify:`` entries
(Waves 5-8) are by-definition "verify existing tests still pass" and
legitimately skip the behave item. The *feat* landings in §4.2 that
lack behave are the real gap.

---

## 5. Missing conftest fixtures / duplication hotspots

There is **no ``tests/conftest.py``** at the top of the tree. The
``unitutil`` package (``tests/unitutil/cxml.py``, ``mock.py``,
``file.py``) holds the builder helpers (``cxml`` element expressions,
``class_mock`` / ``instance_mock`` / ``method_mock`` / ``property_mock``)
— these are in good shape.

The most-duplicated per-class fixture patterns across ``tests/``:

| fixture shape | approx. duplicate count | suggestion |
|---|---:|---|
| ``def _restore_part_factory(self): ... PartFactory.part_type_for.clear(); .update(saved)`` | 12 files × ~50 lines each ≈ **600 lines** | promote to an autouse ``tests/conftest.py`` fixture or delete entirely once the polluter is identified — see §3.3 |
| ``def slide_part_(self, request): return instance_mock(request, SlidePart)`` | 10+ (e.g. ``test_slide.py:553``, ``test_slide.py:1127``, ``parts/test_slide.py:287``, ``parts/test_slide.py:767``, ``test_action.py:468``, ``parts/test_presentation.py:366``, ``shapes/test_placeholder.py:162``, ``shapes/test_picture.py:671``, ``shapes/test_picture.py:868``, ``shapes/test_shapetree.py:951``) | promote to ``tests/conftest.py`` or a ``tests/parts/conftest.py`` |
| ``def presentation_part_(self, request): ...`` | 3 (``test_slide.py:1127``, ``parts/test_slide.py:767``) | local ``conftest.py`` per subpackage |

Promoting these would remove ≈200 lines of boilerplate on top of the
``_restore_part_factory`` removal.

---

## 6. Recommendations (follow-up issue backlog)

Priorities in python-docx style: **P1** correctness / blocker, **P2**
high-ROI coverage or infra, **P3** hygiene.

### P1 — correctness / infra

1. **[S] Resolve the ``PartFactory.part_type_for`` polluter once and
   for all.** Either find the leaking test and fix it at source (the
   likely suspects — ``tests/opc/test_package.py::DescribePartFactory``,
   the ``class_mock`` sites in ``tests/parts/test_presentation.py``,
   ``tests/parts/test_slide.py``, ``tests/shapes/test_shapetree.py`` —
   all appear to use safe ``monkeypatch.setitem`` / local-module
   ``class_mock`` idioms, which suggests the leak may already be
   fixed), **or** add an autouse session/function-scoped fixture at
   ``tests/conftest.py`` that snapshots/restores
   ``PartFactory.part_type_for`` automatically. Either way, **delete
   the twelve ``_restore_part_factory`` fixtures** (~600 lines) after
   confirming the suite still passes under ``--randomly-seed=<N>`` for
   multiple N. This is the single highest-ROI hygiene fix in the repo
   — same magnitude as the 22-guard removal that closed the equivalent
   python-docx issue.

### P2 — coverage fills (high ROI)

2. **[M] ``Chart.apply_template`` (#243) axis / legend / title
   coverage.** 100 lines of ``src/pptx/chart/chart.py`` are uncovered,
   almost all inside the #243 Wave-8 ``apply_template`` implementation
   (``_insert_axis_child`` tag-miss fall-through, ``_apply_legend``
   on a chart missing ``c:plotArea``, ``_apply_title_if_absent``
   no-op when target already has a title, ``_insert_in_schema_order``
   ``ValueError``). Add four parametrised tests in
   ``tests/chart/test_chart.py`` exercising each branch against a
   hand-built fixture ``chartSpace``. ``apply_template`` is a
   headline Wave-8 landing and deserves dense coverage.

3. **[M] ``Font.effective_color`` (#938) master-``txStyles`` fallback
   ordering.** Lines 721-772 of ``src/pptx/text/text.py`` are 45 missed
   statements across the ``_slide_master`` / ``_theme_colors`` /
   ``_master_txStyle_rgb`` cluster. Fixtures where a slide master has
   *only* ``p:bodyStyle`` (no titleStyle / otherStyle), and vice
   versa, would cover the probe-ordering branches. Three parametrised
   tests.

4. **[S] ``ShadowFormat`` / ``GlowFormat`` ``= None`` setter
   branches** in ``src/pptx/dml/effect.py:222-371``. When ``_child`` is
   absent the setter short-circuits; no test exercises that. Five
   one-line parametrised tests.

5. **[S] ``pptx.enum.base.DocsPageFormatter`` smoke test.** Same
   shape as the python-docx recommendation: one test instantiating
   it against e.g. ``WD_BORDER_STYLE.__members__`` (or a local enum
   stub) and asserting the returned string starts with ``.. _``.
   Lifts ``enum/base.py`` from 72 % to ≥95 %.

6. **[S] ``SlideTags.get`` / ``values`` / ``items`` on slides without
   a ``tagLst`` part.** Lines 1817-1837 of ``src/pptx/slide.py``.
   Three one-line tests.

### P2 — behave gaps (high ROI)

7. **[M] Add behave coverage for the six headline Wave-7/Wave-8
   features** that lack scenarios:
   - ``prs-save-ppsx.feature`` (#438 — ``Presentation.save_ppsx``)
   - ``cht-apply-template.feature`` (#243 — ``Chart.apply_template``)
   - ``shp-replace.feature`` (#246 + #834 + #784 — replace_with /
     replace_image / replace_media)

   Each can be ≤50 lines of Gherkin + ~30 lines of step impl, using
   the existing ``features/steps/helpers.py`` and
   ``features/steps/test_files/`` fixture deck. Mirror the
   ``cmt-comments`` / ``cht-clone-to`` / ``prs-font-embedding`` style.

8. **[S] Add a ``prs-merge.feature`` scenario for
   ``Presentation.merge`` (#934).** Currently tested only at unit
   level; behave scenario is the most natural documentation for this
   capability-level API.

### P3 — hygiene / test quality

9. **[S] Promote ``slide_part_`` / ``presentation_part_`` mocks to a
   shared ``tests/conftest.py``.** Removes ~200 lines of boilerplate.
   (Merged with P1 #1, this is a ~800-line diff.)

10. **[S] Tighten ``pytest.raises(ValueError)`` call sites that have
    *meaningful* error messages.** Specifically: ``Emu``/``Pt``/
    ``Inches``/``Cm``/``Mm`` unit parsing, shape-geometry preset
    parsing (``auto_shape.adj`` / ``freeform``), hexadecimal color
    validation, ``PackURI`` validation. Adding ``match=`` tightens
    regressions where the exception *type* is right but the message
    wording drifts (not hypothetical — the #991 XML-decl quote fix
    had to touch a nearby error path).

11. **[P3] ``tests/test_issue_777_html_ole_embed.py:147`` uses
    ``tempfile.NamedTemporaryFile(delete=False)`` with manual
    ``try:/finally: os.unlink``.** Replace with the pytest
    ``tmp_path`` fixture.

12. **[P3] ``tests/parts/test_coreprops.py:150`` bracket-based
    wall-clock assertion**. Safe as-is, but noting the pattern so it
    isn't propagated.

---

## Appendix A — full coverage output

See ``pyproject.toml`` for test configuration. To reproduce:

```
PYTHONPATH=src pytest tests/ --cov=src/pptx --cov-report=term-missing \
  -W ignore::pyparsing.warnings.PyparsingDeprecationWarning -p no:randomly -q
```

Expected outcome: ``4791 passed in ~21s``, overall **97 %** line
coverage.

To reproduce behave:

```
PYTHONPATH=src behave --tags=-wip
```

Expected: ``69 features passed, 0 failed, 0 skipped / 1208 scenarios
passed / 3810 steps passed`` in ~6 s.

To reproduce the random-order probe:

```
pip install pytest-randomly
PYTHONPATH=src pytest tests/ -W ignore::pyparsing.warnings.PyparsingDeprecationWarning \
  -q --randomly-seed=<N>
```

Expected: identical ``4791 passed`` regardless of seed.
