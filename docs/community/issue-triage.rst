Upstream issue triage
=====================

This fork audits every open upstream issue on ``scanny/python-pptx``. Of the
444 items in the 2026-era audit, 337 represented a substantive gap (missing
feature, partial-support regression, or pure bug) and have been resolved by a
shipped feature or verify-close regression test — see ``FEATURES.md`` and
``HISTORY.rst`` for the comprehensive catalogue and changelog.

This page summarises the remaining **107** items that do *not* represent a
code gap in this fork, with a one-line disposition for each. They are retained
here for completeness and so that a user arriving from an upstream issue link
can immediately see where the fork stands on it without having to reopen the
upstream thread.

Disposition categories
----------------------

- **Usage question** — A *how-do-I-accomplish-X* enquiry resolved by the
  existing API; no missing feature. The answer invariably already lives in the
  corresponding ``docs/api/`` reference or ``docs/user/`` tutorial page.
- **Already shipped** — The feature did land — sometimes in this fork,
  sometimes earlier upstream — but the upstream issue was never closed,
  typically because the reporter disengaged before confirming. Regression
  tests in ``tests/test_non_gap_triage.py`` pin the shipped path so any
  accidental regression surfaces in CI.
- **Out of scope** — The request falls outside python-pptx's charter.
  Rendering (PPTX → PNG/SVG/PDF), binary ``.ppt`` (Office 97–2003) support,
  diff / comparator utilities, live COM automation of a running PowerPoint
  instance, and text-coordinate / glyph-position queries all require
  capabilities (a rendering engine, an OLE file-format emitter, a constraint
  solver, an Office COM binding) that are not what this library does.
- **Corrupt source file** — The reporter's ``.pptx`` was damaged on disk: a
  bad ZIP CRC, truncated XML, or malformed output from a third-party emitter.
  python-pptx itself raises a correct exception; fixing the file is up to the
  producer of it. python-pptx's XXE hardening and zip-bomb guard still apply
  in these paths — see the *Encryption and protection* section in
  ``FEATURES.md``.
- **Duplicate of #N** — The issue duplicates another entry that is tracked
  elsewhere. The cross-reference points at the canonical issue, which is
  either already resolved (see ``FEATURES.md``) or still tracked as a gap.
- **Environment** — A Python-version / lxml-wheel / IDE / PyInstaller /
  Docker / Replit / PyPI / OS-package environment problem. The python-pptx
  library itself is not at fault; the fix lives in the environment.
- **Unclear / abandoned** — A bare code-dump with no question asked, a
  single-sentence feature ping with no follow-up, or a thread where the
  reporter disengaged before a maintainer could triage. Nothing actionable
  remains.
- **Documentation** — A pure documentation or comment-typo nit. No code
  change is required and no behaviour is at risk.

Summary
-------

.. list-table::
   :header-rows: 1
   :widths: 30 10 60

   * - Disposition
     - Count
     - Issue numbers
   * - Usage question
     - 38
     - #172, #387, #461, #465, #478, #489, #507, #511, #530, #563, #599, #603, #631, #661, #678, #690, #693, #697, #707, #779, #790, #807, #813, #814, #826, #827, #847, #855, #856, #860, #867, #875, #880, #881, #967, #1037, #1038, #1079
   * - Already shipped
     - 18
     - #473, #538, #540, #541, #553, #564, #614, #665, #667, #671, #680, #684, #710, #729, #794, #841, #962, #968
   * - Out of scope
     - 10
     - #362, #392, #451, #480, #568, #570, #577, #637, #673, #842
   * - Corrupt source file
     - 3
     - #374, #818, #1062
   * - Duplicate of another issue
     - 6
     - #486, #524, #573, #840, #899, #915
   * - Environment
     - 13
     - #466, #548, #630, #648, #732, #755, #795, #796, #804, #817, #831, #854, #959
   * - Unclear / abandoned
     - 14
     - #432, #456, #519, #555, #699, #850, #876, #1051, #1080, #1081, #1082, #1088, #1089, #1092
   * - Documentation
     - 5
     - #592, #717, #802, #869, #1050
   * - **Total**
     - **107**
     - —

Issues by disposition
---------------------

Usage questions
~~~~~~~~~~~~~~~

Each of these upstream issues is a how-do-I-accomplish-X question. The API
needed to achieve the user's goal was already present; these are typically
StackOverflow-style requests that a maintainer or another commenter answered
inline on the issue thread. No code change is indicated; the canonical answer
lives in the corresponding API reference page under ``docs/api/`` or in the
``docs/user/`` tutorial material.

- `#172 <https://github.com/scanny/python-pptx/issues/172>`_ — Suggestion for internal language-tag-to-script helper; not a user-facing gap.
- `#387 <https://github.com/scanny/python-pptx/issues/387>`_ — How does the library open files? — usage question.
- `#461 <https://github.com/scanny/python-pptx/issues/461>`_ — ResourceWarning about unclosed file; Presentation.save() handles closing.
- `#465 <https://github.com/scanny/python-pptx/issues/465>`_ — Broad "can you do X" product-feature scoping question.
- `#478 <https://github.com/scanny/python-pptx/issues/478>`_ — Performance / lazy-XML architectural discussion.
- `#489 <https://github.com/scanny/python-pptx/issues/489>`_ — Support question about user code; maintainer referred to StackOverflow.
- `#507 <https://github.com/scanny/python-pptx/issues/507>`_ — Multi-level category axis usage question; no concrete gap articulated.
- `#511 <https://github.com/scanny/python-pptx/issues/511>`_ — Slide-reorder workaround question; tracked separately.
- `#530 <https://github.com/scanny/python-pptx/issues/530>`_ — Cell centering; alignment/anchor APIs exist.
- `#563 <https://github.com/scanny/python-pptx/issues/563>`_ — How to load a pptx and change text — Stack-Overflow-style.
- `#599 <https://github.com/scanny/python-pptx/issues/599>`_ — python-pptx is a library, not a CLI tool.
- `#603 <https://github.com/scanny/python-pptx/issues/603>`_ — Create-slides-from-template usage question.
- `#631 <https://github.com/scanny/python-pptx/issues/631>`_ — Image/text extraction; existing API covers it.
- `#661 <https://github.com/scanny/python-pptx/issues/661>`_ — Line-chart label/title rotation — no clear gap.
- `#678 <https://github.com/scanny/python-pptx/issues/678>`_ — User set a:lstStyle default pPr, not paragraph alignment; workaround posted.
- `#690 <https://github.com/scanny/python-pptx/issues/690>`_ — User-code ordering issue; distinct from tracked Asian-font gap.
- `#693 <https://github.com/scanny/python-pptx/issues/693>`_ — Architectural question about lxml element dispatch.
- `#697 <https://github.com/scanny/python-pptx/issues/697>`_ — User forgot .pptx file extension; closed by user.
- `#707 <https://github.com/scanny/python-pptx/issues/707>`_ — User read a:sysClr/@val directly, needs lastClr fallback.
- `#779 <https://github.com/scanny/python-pptx/issues/779>`_ — Raw part.blob write request; covered by oxml manipulation.
- `#790 <https://github.com/scanny/python-pptx/issues/790>`_ — Text-coordinates question; no reproducer.
- `#807 <https://github.com/scanny/python-pptx/issues/807>`_ — User not using placeholder-bearing layout.
- `#813 <https://github.com/scanny/python-pptx/issues/813>`_ — Question about OS dependencies.
- `#814 <https://github.com/scanny/python-pptx/issues/814>`_ — User passed MSO_SHAPE_TYPE to add_shape instead of MSO_SHAPE.
- `#826 <https://github.com/scanny/python-pptx/issues/826>`_ — User shadowed pres variable; Slides.slides AttributeError.
- `#827 <https://github.com/scanny/python-pptx/issues/827>`_ — Chart-adding example / StackOverflow-style.
- `#847 <https://github.com/scanny/python-pptx/issues/847>`_ — Emu-type validation on add_movie parameters; minor UX polish.
- `#855 <https://github.com/scanny/python-pptx/issues/855>`_ — pandas DataFrame to table styling usage question.
- `#856 <https://github.com/scanny/python-pptx/issues/856>`_ — Setter for element.xml; design question.
- `#860 <https://github.com/scanny/python-pptx/issues/860>`_ — Multiple charts on one slide: call add_chart twice.
- `#867 <https://github.com/scanny/python-pptx/issues/867>`_ — Wrong import MSO_SHAPE_TYPE vs MSO_SHAPE; resolved in comment.
- `#875 <https://github.com/scanny/python-pptx/issues/875>`_ — Returned prs object via HttpResponse instead of saved bytes.
- `#880 <https://github.com/scanny/python-pptx/issues/880>`_ — Bullet/sort type detection via paragraph.level; no gap.
- `#881 <https://github.com/scanny/python-pptx/issues/881>`_ — Loading Presentation from URL via BytesIO is user responsibility.
- `#967 <https://github.com/scanny/python-pptx/issues/967>`_ — add_ole_object with proprietary .h3d progId; overlaps #981.
- `#1037 <https://github.com/scanny/python-pptx/issues/1037>`_ — Update table data preserving format — template approach suggested.
- `#1038 <https://github.com/scanny/python-pptx/issues/1038>`_ — Invalid en-CN locale; user error.
- `#1079 <https://github.com/scanny/python-pptx/issues/1079>`_ — Slide reorder request mixed into a usage/workaround thread.

Already shipped
~~~~~~~~~~~~~~~

The code path these issues request is present and working. The upstream
issues were never closed, often because the original reporter did not
confirm the fix. The regression-test module ``tests/test_non_gap_triage.py``
pins the behaviour via a minimal public-API reproduction for each item
listed below; see that file for copy-paste-ready snippets.

- `#473 <https://github.com/scanny/python-pptx/issues/473>`_ — chart.value_axis.visible toggle is supported.
- `#538 <https://github.com/scanny/python-pptx/issues/538>`_ — Slide notes (notes_slide) and review comments (slide.comments) both ship.
- `#540 <https://github.com/scanny/python-pptx/issues/540>`_ — Per-point data labels already supported.
- `#541 <https://github.com/scanny/python-pptx/issues/541>`_ — Picture.image.blob exposes raw bytes for PIL consumption.
- `#553 <https://github.com/scanny/python-pptx/issues/553>`_ — Picture.image.blob extracts embedded images.
- `#564 <https://github.com/scanny/python-pptx/issues/564>`_ — Font.highlight_color works in table cells (a:highlight on rPr).
- `#614 <https://github.com/scanny/python-pptx/issues/614>`_ — shape.fill.fore_color.rgb exposes shape/background colour.
- `#665 <https://github.com/scanny/python-pptx/issues/665>`_ — None values in line-chart series already produce gaps (xmlwriter.py:1604-1606).
- `#667 <https://github.com/scanny/python-pptx/issues/667>`_ — number_format strings are passed through to Excel which handles locale.
- `#671 <https://github.com/scanny/python-pptx/issues/671>`_ — Slide.name read/write works (src/pptx/slide.py:57-68).
- `#680 <https://github.com/scanny/python-pptx/issues/680>`_ — Chart.replace_data replaces chart categories/values.
- `#684 <https://github.com/scanny/python-pptx/issues/684>`_ — Replace-text via run.text preserves formatting.
- `#710 <https://github.com/scanny/python-pptx/issues/710>`_ — paragraph.line_spacing=Pt(18) now dispatches to spcPts (oxml/text.py:527-535).
- `#729 <https://github.com/scanny/python-pptx/issues/729>`_ — Length arithmetic with Emu() works as explained by scanny.
- `#794 <https://github.com/scanny/python-pptx/issues/794>`_ — bubble_scale on plot (chart/plot.py:179).
- `#841 <https://github.com/scanny/python-pptx/issues/841>`_ — shape.shadow exists (dml/effect.py); discoverability issue.
- `#962 <https://github.com/scanny/python-pptx/issues/962>`_ — run.hyperlink.address is supported on text runs (including runs in a textbox over a chart); the upstream reporter called a non-existent add_hyperlink_relationship method.
- `#968 <https://github.com/scanny/python-pptx/issues/968>`_ — dispBlanksAs=gap hard-coded for line charts (xmlwriter.py:486).

Out of scope
~~~~~~~~~~~~

These requests fall outside python-pptx's charter. Rendering, binary
``.ppt`` support, diff utilities, presentation comparators, and saving slides
as images all require capabilities (a rendering engine, OLE file-format
emitter, layout solver, Office COM binding) that are not what this library
does. Users wanting these capabilities should look at LibreOffice's headless
mode, Aspose.Slides, or a dedicated rendering project.

- `#362 <https://github.com/scanny/python-pptx/issues/362>`_ — Legacy binary .ppt format is out of scope (not OOXML).
- `#392 <https://github.com/scanny/python-pptx/issues/392>`_ — Asking for editor/plugin recommendations, not a python-pptx feature.
- `#451 <https://github.com/scanny/python-pptx/issues/451>`_ — Request to build a presentation diff tool inside python-pptx.
- `#480 <https://github.com/scanny/python-pptx/issues/480>`_ — Dynamic row-height recomputation requires font metrics; PowerPoint itself does not do this.
- `#568 <https://github.com/scanny/python-pptx/issues/568>`_ — PowerPoint has no native table title; not a gap.
- `#570 <https://github.com/scanny/python-pptx/issues/570>`_ — Full slide-copy across presentations is intentionally out of scope.
- `#577 <https://github.com/scanny/python-pptx/issues/577>`_ — Rendering to SVG/EPS is out of scope.
- `#637 <https://github.com/scanny/python-pptx/issues/637>`_ — PowerPoint diff tool is out of scope.
- `#673 <https://github.com/scanny/python-pptx/issues/673>`_ — Saving chart as picture requires a renderer; out of scope.
- `#842 <https://github.com/scanny/python-pptx/issues/842>`_ — Character coordinate locating requires a rendering engine.

Corrupt source file
~~~~~~~~~~~~~~~~~~~

The source ``.pptx`` the reporter opened was damaged on disk (bad ZIP CRC,
truncated XML, malformed emission from a third-party generator).
python-pptx itself raises a correct exception; fixing the file is the
responsibility of the tool that produced it. A practical workaround is to
open the file in PowerPoint or LibreOffice, save a fresh copy, and reopen.

- `#374 <https://github.com/scanny/python-pptx/issues/374>`_ — XMLSyntaxError from malformed third-party .pptx; reporter refused reopen.
- `#818 <https://github.com/scanny/python-pptx/issues/818>`_ — BadZipFile from corrupt .pptx on disk; legacy .ppt needs LibreOffice.
- `#1062 <https://github.com/scanny/python-pptx/issues/1062>`_ — Generated-file-corrupt report with no MCVE; PowerPoint re-save fixes it.

Duplicates of other issues
~~~~~~~~~~~~~~~~~~~~~~~~~~

Each of these issues duplicates another entry that is tracked elsewhere.
The cross-reference points at the canonical issue (or a shipped feature if
the canonical issue has since been closed).

- `#486 <https://github.com/scanny/python-pptx/issues/486>`_ — Duplicate of #403 (and older #132) slide-portability.
- `#524 <https://github.com/scanny/python-pptx/issues/524>`_ — Duplicate of #71 (cell border colour).
- `#573 <https://github.com/scanny/python-pptx/issues/573>`_ — Duplicate of #71 (cell borders); same root cause as #524.
- `#840 <https://github.com/scanny/python-pptx/issues/840>`_ — Overlaps with #883 (style-hierarchy font resolution).
- `#899 <https://github.com/scanny/python-pptx/issues/899>`_ — Near-duplicate of #925 (grouped-shape position/size).
- `#915 <https://github.com/scanny/python-pptx/issues/915>`_ — SVG content-type; same gap as #885.

Environment / install / packaging
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

These reports describe Python / lxml wheel / IDE / PyInstaller / Docker
/ Replit / PyPI / OS-package environment problems. The python-pptx library
itself is not at fault; the fix is in the environment.

- `#466 <https://github.com/scanny/python-pptx/issues/466>`_ — PHP/COM question; wrong project.
- `#548 <https://github.com/scanny/python-pptx/issues/548>`_ — Wheel packaging request; modern releases ship wheels.
- `#630 <https://github.com/scanny/python-pptx/issues/630>`_ — Install/import error (module vs class confusion).
- `#648 <https://github.com/scanny/python-pptx/issues/648>`_ — IDE autocomplete config issue.
- `#732 <https://github.com/scanny/python-pptx/issues/732>`_ — PyPI stable-version meta request.
- `#755 <https://github.com/scanny/python-pptx/issues/755>`_ — PyCharm autocomplete on Presentation() factory function.
- `#795 <https://github.com/scanny/python-pptx/issues/795>`_ — Replit-specific lxml import error.
- `#796 <https://github.com/scanny/python-pptx/issues/796>`_ — Docker font-directory OS packaging issue (workaround via font_file= in fit_text).
- `#804 <https://github.com/scanny/python-pptx/issues/804>`_ — PyInstaller packaging, self-resolved.
- `#817 <https://github.com/scanny/python-pptx/issues/817>`_ — IntelliSense/IDE question.
- `#831 <https://github.com/scanny/python-pptx/issues/831>`_ — PyInstaller template data-file missing at runtime.
- `#854 <https://github.com/scanny/python-pptx/issues/854>`_ — Python 3.11 lxml wheel issue; no longer relevant.
- `#959 <https://github.com/scanny/python-pptx/issues/959>`_ — PyInstaller/py2exe packaging issue.

Unclear / abandoned
~~~~~~~~~~~~~~~~~~~

Bare code-dumps with no question asked, single-sentence feature pings
with no follow-up, or threads where the reporter disengaged before a
maintainer could triage. Nothing actionable remains — these are retained
in the audit for accounting completeness only.

- `#432 <https://github.com/scanny/python-pptx/issues/432>`_ — Note about backwards-compat issue in a third-party project; no action requested.
- `#456 <https://github.com/scanny/python-pptx/issues/456>`_ — User could not share code; no reproducer, no MCVE.
- `#519 <https://github.com/scanny/python-pptx/issues/519>`_ — Issue body is a bare image with no description.
- `#555 <https://github.com/scanny/python-pptx/issues/555>`_ — One-line UngroupShapes request with no engagement.
- `#699 <https://github.com/scanny/python-pptx/issues/699>`_ — Vague "table not shown" with no repro.
- `#850 <https://github.com/scanny/python-pptx/issues/850>`_ — Chinese rant about a blog-post script, no reproducer.
- `#876 <https://github.com/scanny/python-pptx/issues/876>`_ — Incomplete code snippet, not reproducible.
- `#1051 <https://github.com/scanny/python-pptx/issues/1051>`_ — Keynote speaker-notes interop, no repro.
- `#1080 <https://github.com/scanny/python-pptx/issues/1080>`_ — Arabic code dump with no question.
- `#1081 <https://github.com/scanny/python-pptx/issues/1081>`_ — Duplicate code-dump of #1080.
- `#1082 <https://github.com/scanny/python-pptx/issues/1082>`_ — Another duplicate code-dump.
- `#1088 <https://github.com/scanny/python-pptx/issues/1088>`_ — Unrelated zipfile snippet with no question.
- `#1089 <https://github.com/scanny/python-pptx/issues/1089>`_ — PDF to PPTX conversion code, no question.
- `#1092 <https://github.com/scanny/python-pptx/issues/1092>`_ — Portuguese narrative + code, no articulated issue.

Documentation only
~~~~~~~~~~~~~~~~~~

Pure documentation or comment-typo nit. No code change; the fork's
documentation corpus under ``docs/`` covers the surface these reporters
were asking about.

- `#592 <https://github.com/scanny/python-pptx/issues/592>`_ — Trivial docs/comment typo in lab/parse_xsd/parse_xsd.py.
- `#717 <https://github.com/scanny/python-pptx/issues/717>`_ — Pure docs request.
- `#802 <https://github.com/scanny/python-pptx/issues/802>`_ — Documentation of testing targets outdated.
- `#869 <https://github.com/scanny/python-pptx/issues/869>`_ — Installation docs nit.
- `#1050 <https://github.com/scanny/python-pptx/issues/1050>`_ — Open file-like in binary mode — at most a docs nit.

Where to go from here
---------------------

- If you arrived here from an upstream issue and the feature you want is
  listed under *Already shipped*, see ``FEATURES.md`` for the canonical
  snippet and ``tests/test_non_gap_triage.py`` for a minimal working example.
- If the issue is classified as *Usage question*, the corresponding
  ``docs/user/`` tutorial or ``docs/api/`` reference page answers it.
- If the issue is *Out of scope*, consider LibreOffice's headless mode,
  Aspose.Slides, or a rendering-engine project layered on top of the raw
  OOXML that python-pptx gives you access to.
- If the issue is *Duplicate of #N*, follow the cross-reference and check
  the target issue's disposition in ``FEATURES.md`` or ``HISTORY.rst``.
- If the issue is *Corrupt source file*, open the file in PowerPoint or
  LibreOffice, save a fresh copy, and retry with the fresh copy.
- If the issue is *Environment*, check the install instructions in
  ``README.md`` and the ``docs/user/install.rst`` page.
- For anything else, please file an issue against this fork at
  https://github.com/loadfix/python-pptx/issues.
