# pyright: reportPrivateUsage=false

"""Regression test for issue #1063 — "get font properties (color, size, name)".

Issue #1063 (https://github.com/scanny/python-pptx/issues/1063) reports the
classic "how do I read what PowerPoint actually renders?" workflow: the
caller walks an existing deck, and for each run wants to know the colour,
size, and typeface PowerPoint would apply — regardless of whether the
value was authored on the run, the enclosing paragraph, the shape's
``a:lstStyle``, or inherited unchanged from the slide-master's
``p:txStyles``. The legacy :attr:`.Font.color` / :attr:`.Font.size` /
:attr:`.Font.name` APIs only inspect the run's own ``a:rPr`` and return
|None| (or raise, in the colour case) for any run that inherits its
value from an ancestor, which is the default for placeholder-authored
decks.

Issue #1063 overlaps heavily with two earlier reports:

  * #378 — "get effective font settings (size/bold/italic/name)". Wave 12
    (``feat/issue-378-effective-font-size``) shipped
    :attr:`.Font.effective_size`, :attr:`.Font.effective_bold`,
    :attr:`.Font.effective_italic`, and :attr:`.Font.effective_name`.
  * #765 — "read the actual font properties after loading a deck". Same
    underlying ask, with the emphasis on a round-trip read of an
    already-saved ``.pptx``.

And with the colour-only precursor:

  * #938 — Wave 6 shipped :attr:`.Font.effective_color` (the read-side fix
    for #1013's "color is empty when set via placeholder").

Together the ``Font.effective_*`` chain closes #1063: a caller can walk
any deck and read ``run.font.effective_color`` /
``run.font.effective_size`` / ``run.font.effective_name`` /
``run.font.effective_bold`` / ``run.font.effective_italic`` to get the
resolved value (or |None| when nothing in the chain declares one) without
ever poking at raw ``a:rPr`` XML.

#1063 is therefore verified-and-closed by the combination of #378, #765,
and #938. This suite pins the #1063 reporter's specific workflow — walk a
deck, read the effective font props of every run — against the scenarios
the report emphasises:

  * walking a newly authored deck and reading every run's effective
    colour/size/name — never raising, never returning nonsense;
  * reading effective props on runs inside table cells (a common source
    of #1063-style confusion, since table-cell runs also inherit from the
    shape's ``a:lstStyle``);
  * reading effective props on placeholders (title, body) where the run
    carries no explicit rPr and every value is inherited from the master;
  * reading effective props on a run whose colour is a theme/scheme
    colour (``a:schemeClr``) — the resolved RGB must fall out of the
    theme walk;
  * a save + reopen round-trip, because #765's framing is "read after
    loading" — the effective values must match before and after.

Any future regression that reintroduces a |None| on a run whose value is
plainly visible in PowerPoint, or that raises an exception inside the
inheritance walk, will reproduce the #1063 symptom and be caught here.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Emu, Inches, Length, Pt


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of ``PartFactory.part_type_for``.

    Mirrors the guard used by ``tests/test_issue_1013_placeholder_color_verify.py``
    — ``Presentation(stream)`` reopen dispatches through
    ``PartFactory.part_type_for``, which earlier test modules overwrite
    with mocks and do not restore.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.image import ImagePart
    from pptx.parts.media import MediaPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import (
        NotesMasterPart,
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    saved = dict(PartFactory.part_type_for)
    expected = {
        CT.PML_PRESENTATION_MAIN: PresentationPart,
        CT.PML_PRES_MACRO_MAIN: PresentationPart,
        CT.PML_TEMPLATE_MAIN: PresentationPart,
        CT.PML_SLIDESHOW_MAIN: PresentationPart,
        CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


def _iter_runs(prs):
    """Yield every ``_Run`` in every text frame in ``prs``.

    Walks top-level shapes, table cells, and groups — i.e. everywhere
    the #1063 reporter's "walk a deck" workflow would reach.
    """
    from pptx.shapes.graphfrm import GraphicFrame

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        yield run
            elif isinstance(shape, GraphicFrame) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                yield run


class DescribeIssue1063EffectiveStyle:
    """#1063 effective-font-props read-path verify-and-close via #378/#938/#765.

    Pins that :attr:`.Font.effective_color`, :attr:`.Font.effective_size`,
    :attr:`.Font.effective_name`, :attr:`.Font.effective_bold`, and
    :attr:`.Font.effective_italic` together close the #1063 workflow — the
    read path the reporter wanted when the legacy ``font.color`` /
    ``font.size`` / ``font.name`` accessors came back empty on inherited
    values.
    """

    # -- API surface ------------------------------------------------------

    def it_exposes_the_full_effective_chain_on_Font(self):
        """Guard: the public read-side API that resolves #1063 is present.

        A refactor that renamed or dropped any of the five
        ``effective_*`` properties would silently reintroduce #1063.
        """
        from pptx.text.text import Font

        for attr in (
            "effective_color",
            "effective_size",
            "effective_name",
            "effective_bold",
            "effective_italic",
        ):
            assert hasattr(Font, attr), f"Font lacks {attr}"

    # -- walk-every-run happy path (the #1063 workflow) ------------------

    def it_walks_every_run_without_raising_and_returns_valid_shapes(self):
        """The headline #1063 scenario — walk a deck, read effective
        props on every run. Must never raise; every returned value must
        be a valid shape (|None|, |RGBColor|, |Length|, |str|, |bool|).

        Authors a title, body, table, and textbox to cover the shape
        kinds the #1063 reporter's deck would contain, then iterates
        every run and inspects the five effective properties.
        """
        prs = Presentation()

        # -- slide 1: title + content --
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.placeholders[0].text_frame.text = "Title slide"
        s1.placeholders[1].text_frame.text = "Subtitle text"

        # -- slide 2: title + body --
        s2 = prs.slides.add_slide(prs.slide_layouts[1])
        s2.placeholders[0].text_frame.text = "Body slide"
        s2.placeholders[1].text_frame.text = "Body line 1"

        # -- slide 3: blank + textbox + table --
        s3 = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s3.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tb.text_frame.text = "Textbox text"
        tbl = s3.shapes.add_table(
            rows=2, cols=2, left=Inches(1), top=Inches(3), width=Inches(5), height=Inches(2)
        ).table
        for r in range(2):
            for c in range(2):
                tbl.cell(r, c).text = f"r{r}c{c}"

        # -- walk every run and read every effective property --
        run_count = 0
        for run in _iter_runs(prs):
            run_count += 1
            color = run.font.effective_color
            size = run.font.effective_size
            name = run.font.effective_name
            bold = run.font.effective_bold
            italic = run.font.effective_italic

            # -- types match the API contract, including None --
            assert color is None or isinstance(color, RGBColor)
            assert size is None or isinstance(size, Length)
            assert name is None or isinstance(name, str)
            assert bold is None or isinstance(bold, bool)
            assert italic is None or isinstance(italic, bool)

            # -- validity: no zero-length name, no negative size --
            if name is not None:
                assert name != ""
            if size is not None:
                assert size > 0

        # -- sanity: we actually iterated something; a silent "no runs" --
        # -- would be a false-pass on the loop invariants --
        assert run_count >= 6  # 2 titles + 2 bodies + 1 textbox + 4 cells

    # -- placeholders (title / body) -------------------------------------

    def it_reads_effective_props_on_title_placeholder_runs(self):
        """The title-placeholder branch of the #1063 walk.

        A fresh title placeholder has no rPr on its run — every
        effective property must fall through to an ancestor (or
        legitimately resolve to |None|) rather than raising.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        title = slide.placeholders[0]
        title.text_frame.text = "The Title"
        run = title.text_frame.paragraphs[0].runs[0]

        # -- inherited colour from master titleStyle (default = black) --
        assert run.font.effective_color == RGBColor(0, 0, 0)
        # -- master titleStyle declares a size (44pt in default theme) --
        size = run.font.effective_size
        assert size is not None
        assert size > Pt(10)
        # -- master titleStyle declares a theme-font name --
        assert run.font.effective_name is not None

    def it_reads_effective_props_on_body_placeholder_runs(self):
        """The body-placeholder branch of the #1063 walk.

        Body placeholders inherit from the master's ``p:bodyStyle``
        rather than ``p:titleStyle``; the walk must consult that
        fall-back too.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Body text"
        run = body.text_frame.paragraphs[0].runs[0]

        assert run.font.effective_color == RGBColor(0, 0, 0)
        size = run.font.effective_size
        assert size is not None
        assert size > Pt(10)
        assert run.font.effective_name is not None

    # -- table cells -----------------------------------------------------

    def it_reads_effective_props_on_table_cell_runs(self):
        """Table-cell runs are a common #1063 blind spot.

        A cell's run also inherits from the graphic-frame shape's
        ``a:lstStyle`` when the table-style defines one, plus the
        master's ``p:otherStyle`` fallback. The walk must not fall
        over on any of those hops.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        graphic_frame = slide.shapes.add_table(
            rows=2, cols=2, left=Inches(1), top=Inches(1), width=Inches(4), height=Inches(2)
        )
        table = graphic_frame.table
        for r in range(2):
            for c in range(2):
                table.cell(r, c).text = f"r{r}c{c}"

        seen = 0
        for row in table.rows:
            for cell in row.cells:
                run = cell.text_frame.paragraphs[0].runs[0]
                seen += 1
                # -- colour resolves (master otherStyle default = black) --
                color = run.font.effective_color
                assert color is None or isinstance(color, RGBColor)
                # -- size resolves to a positive length (or None) --
                size = run.font.effective_size
                assert size is None or (isinstance(size, Length) and size > 0)
                # -- name resolves to a theme string or real typeface --
                name = run.font.effective_name
                assert name is None or (isinstance(name, str) and name != "")

        assert seen == 4

    def it_honours_explicit_cell_run_values_over_inherited(self):
        """Sanity: a cell run with explicit rPr values wins.

        Mirrors the #1013 "explicit run colour beats paragraph/master"
        check, but on the table branch of the tree.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        graphic_frame = slide.shapes.add_table(
            rows=1, cols=1, left=Inches(1), top=Inches(1), width=Inches(3), height=Inches(1)
        )
        cell = graphic_frame.table.cell(0, 0)
        cell.text = "Hello"
        run = cell.text_frame.paragraphs[0].runs[0]
        run.font.color.rgb = RGBColor(0x11, 0x22, 0x33)
        run.font.size = Pt(22)
        run.font.name = "Consolas"
        run.font.bold = True
        run.font.italic = True

        assert run.font.effective_color == RGBColor(0x11, 0x22, 0x33)
        assert run.font.effective_size == Pt(22)
        assert run.font.effective_name == "Consolas"
        assert run.font.effective_bold is True
        assert run.font.effective_italic is True

    # -- theme-colour resolution -----------------------------------------

    def it_resolves_effective_color_on_a_theme_colored_run(self):
        """A run coloured via ``MSO_THEME_COLOR.ACCENT_*`` must resolve.

        The walk's scheme-colour lookup against the slide master's
        theme is the hot path for theme-coloured decks — a regression
        that dropped the theme hop would make ``effective_color`` come
        back |None| here even though the text obviously renders in the
        theme's accent RGB.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tb.text_frame.text = "Accent2 text"
        run = tb.text_frame.paragraphs[0].runs[0]
        run.font.color.theme_color = MSO_THEME_COLOR.ACCENT_2

        rgb = run.font.effective_color
        assert isinstance(rgb, RGBColor)
        # -- default theme's accent2 == (192, 80, 77) --
        assert rgb == RGBColor(0xC0, 0x50, 0x4D)

    def it_resolves_effective_color_on_a_theme_colored_placeholder_run(self):
        """Theme-colour resolution on an inherited placeholder run.

        Combines the placeholder walk (run → master ``p:bodyStyle``) with
        the theme hop — paragraph-level ``a:schemeClr`` on the body
        placeholder's paragraph defRPr must reach the theme and return a
        concrete RGB rather than |None|.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Accent3 inherited"
        para = body.text_frame.paragraphs[0]
        para.font.color.theme_color = MSO_THEME_COLOR.ACCENT_3

        run = para.runs[0]
        # -- legacy .color still says "inherited" on the run --
        assert run.font.color.type is None
        # -- effective_color walks up to the paragraph defRPr and through
        # -- the theme to accent3 == (155, 187, 89) --
        assert run.font.effective_color == RGBColor(0x9B, 0xBB, 0x59)

    # -- round-trip through save + reopen (#765's framing) ---------------

    def it_round_trips_effective_props_through_save_and_reload(self, _restore_part_factory):
        """#765's "read after loading" shape of the same workflow.

        Authors a body placeholder with paragraph-level size/name/color,
        saves to an in-memory ``.pptx``, reopens, and confirms every
        effective property matches the pre-save value — the #1063
        reporter cares about decks on disk, not just the ones they
        authored in the same process.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Round-trip text"
        para = body.text_frame.paragraphs[0]
        para.font.size = Pt(28)
        para.font.name = "Calibri"
        para.font.color.rgb = RGBColor(0x33, 0x66, 0x99)
        para.font.bold = True

        # -- capture the pre-save snapshot --
        run_pre = para.runs[0]
        pre_color = run_pre.font.effective_color
        pre_size = run_pre.font.effective_size
        pre_name = run_pre.font.effective_name
        pre_bold = run_pre.font.effective_bold
        pre_italic = run_pre.font.effective_italic

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        body2 = prs2.slides[0].placeholders[1]
        run_post = body2.text_frame.paragraphs[0].runs[0]

        assert run_post.text == "Round-trip text"
        assert run_post.font.effective_color == pre_color == RGBColor(0x33, 0x66, 0x99)
        assert run_post.font.effective_size == pre_size == Pt(28)
        assert run_post.font.effective_name == pre_name == "Calibri"
        assert run_post.font.effective_bold == pre_bold is True
        # -- italic was never set; inherited walk returns the same value --
        assert run_post.font.effective_italic == pre_italic

    def it_round_trips_table_cell_effective_props_through_save_and_reload(
        self, _restore_part_factory
    ):
        """Table cells survive save+reopen too.

        A separate scenario because the #1063 reporter's real deck would
        contain tables — and the table graphic-frame is a different part
        path than the shape-tree. A regression in the cell-level walk
        could slip past the placeholder-only round-trip above.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        graphic_frame = slide.shapes.add_table(
            rows=1, cols=1, left=Inches(1), top=Inches(1), width=Inches(4), height=Inches(1)
        )
        cell = graphic_frame.table.cell(0, 0)
        cell.text = "Cell text"
        run_pre = cell.text_frame.paragraphs[0].runs[0]
        run_pre.font.size = Pt(18)
        run_pre.font.name = "Cambria"
        run_pre.font.color.rgb = RGBColor(0x80, 0x20, 0x40)
        run_pre.font.italic = True

        pre_color = run_pre.font.effective_color
        pre_size = run_pre.font.effective_size
        pre_name = run_pre.font.effective_name
        pre_italic = run_pre.font.effective_italic

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        gf2 = prs2.slides[0].shapes[0]
        run_post = gf2.table.cell(0, 0).text_frame.paragraphs[0].runs[0]

        assert run_post.text == "Cell text"
        assert run_post.font.effective_color == pre_color == RGBColor(0x80, 0x20, 0x40)
        assert run_post.font.effective_size == pre_size == Pt(18)
        assert run_post.font.effective_name == pre_name == "Cambria"
        assert run_post.font.effective_italic == pre_italic is True

    def it_round_trips_theme_colored_run_effective_color(self, _restore_part_factory):
        """Scheme-colour resolution survives save+reopen.

        The theme-walk half of the fix is the most failure-prone part of
        the inheritance chain; reopening needs to re-hydrate the
        master→theme link correctly or the reader will see |None| on a
        run PowerPoint still renders in the accent colour.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        tb.text_frame.text = "Theme-colored"
        run_pre = tb.text_frame.paragraphs[0].runs[0]
        run_pre.font.color.theme_color = MSO_THEME_COLOR.ACCENT_4

        pre_color = run_pre.font.effective_color

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        tb2 = prs2.slides[0].shapes[0]
        run_post = tb2.text_frame.paragraphs[0].runs[0]

        # -- default theme's accent4 == (128, 100, 162) --
        assert pre_color == RGBColor(0x80, 0x64, 0xA2)
        assert run_post.font.effective_color == pre_color

    # -- walk-the-reloaded-deck invariant (the #1063 closing shape) ------

    def it_walks_the_reloaded_deck_and_every_run_reports_valid_effective_props(
        self, _restore_part_factory
    ):
        """Compose the two branches: author a multi-shape deck, save,
        reopen, walk every run, and confirm the effective properties
        still satisfy the API shape invariants.

        This is the literal shape of the #1063 reporter's workflow — a
        fork caller who has an existing .pptx and wants to tabulate the
        font settings per run.
        """
        prs = Presentation()

        # -- title/content --
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.placeholders[0].text_frame.text = "Deck title"
        s1.placeholders[1].text_frame.text = "Deck subtitle"

        # -- body slide --
        s2 = prs.slides.add_slide(prs.slide_layouts[1])
        s2.placeholders[0].text_frame.text = "Body slide"
        body = s2.placeholders[1]
        body.text_frame.text = "A line with settings"
        para = body.text_frame.paragraphs[0]
        para.font.size = Pt(22)
        para.font.color.theme_color = MSO_THEME_COLOR.ACCENT_5

        # -- blank w/ table + textbox --
        s3 = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s3.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tb.text_frame.text = "Textbox"
        gf = s3.shapes.add_table(
            rows=2, cols=2, left=Inches(1), top=Inches(3), width=Inches(4), height=Inches(2)
        )
        for r in range(2):
            for c in range(2):
                gf.table.cell(r, c).text = f"c{r}{c}"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run_count = 0
        for run in _iter_runs(prs2):
            run_count += 1
            color = run.font.effective_color
            size = run.font.effective_size
            name = run.font.effective_name
            bold = run.font.effective_bold
            italic = run.font.effective_italic

            assert color is None or isinstance(color, RGBColor)
            assert size is None or isinstance(size, Length)
            assert name is None or (isinstance(name, str) and name != "")
            assert bold is None or isinstance(bold, bool)
            assert italic is None or isinstance(italic, bool)
            if size is not None:
                assert size > Emu(0)

        # -- 2 slide-1 phs + 2 slide-2 phs + 1 textbox + 4 table cells --
        assert run_count >= 9
