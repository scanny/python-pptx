# pyright: reportPrivateUsage=false

"""Regression test for issue #696 — move a slide from one .pptx to another.

Issue #696 (https://github.com/scanny/python-pptx/issues/696) asks:

    "I've a project and one of the requirements is to move a slide from one
    presentation to another, but i can't to find a solution and i think to
    doesn't exists one, i need to do this with this code

        prs = Presentation('ppt2.pptx')

        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text'):
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if 'Nuestra experiencia en su industria' in run.text:
                                url = 'ppt2.pptx'
                                prsIndustria = Presentation(url)
                                slideAn = prsIndustria.slides[0]
                                slide = slideAn
    ..."

The reporter's scenario is a classic text-keyed "graft a slide from another
.pptx in place of / alongside the one I found" workflow. Two shipped APIs
cover this between them:

  * Wave 1/2 #1036 landed :meth:`Slides.add_slide_from_external`, which
    appends a full-fidelity clone of a foreign slide to the receiver. Wave 7
    #934 promoted the underlying cloning pipeline so image, media, chart
    (distinct embedded workbook), OLE-object, and external-hyperlink
    relationships are all materialised in the target package.
  * Wave 7 #934 also shipped :meth:`Presentation.merge` for appending every
    slide of another presentation in one call.

Combined with the pre-existing :meth:`Slides.move_slide` reorder and
:meth:`Slides.delete` remove, the #696 workflow is now a three-step recipe:

    target = Presentation("ppt1.pptx")
    other = Presentation("ppt2.pptx")
    cloned = target.slides.add_slide_from_external(other.slides[0],
                                                    target.slide_layouts[5])
    target.slides.move_slide(cloned, new_position)   # optional reorder
    target.slides.delete(slide_to_replace)            # optional prune

This regression suite pins the end-to-end behaviour so #696 can be closed as
*resolved by #1036 + #934*:

  * A slide carrying the reporter's "Nuestra experiencia en su industria"
    text can be located in a source deck by iterating ``slides/shapes/
    text_frame/paragraphs/runs`` and cloned into a target deck, preserving
    the text content.
  * The cloned slide can be moved to any position in the target sequence
    (the "put the imported slide right where the old one was" piece of the
    workflow) and the slide being replaced can be deleted.
  * A slide carrying a picture round-trips through the external-copy + save
    + reopen cycle.
  * The source presentation is not mutated by the copy, so the reporter's
    implied "I keep ppt2.pptx as a library I pull slides from" workflow
    holds.
  * A self-copy (cloning a slide from this presentation into itself) works —
    :meth:`Slides.add_slide_from_external` explicitly allows ``source_slide``
    to come from any presentation, including ``self``.
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING, cast

import pytest

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationClass
    from pptx.slide import Slide
    from pptx.text.text import TextFrame


@pytest.fixture
def _restore_part_factory():
    """Guard against ``PartFactory.part_type_for`` pollution from other tests.

    Mirrors the identical guard used by ``tests/test_issue_175_*.py``,
    ``tests/test_issue_403_*.py``, and siblings; earlier test modules
    overwrite slide-part registrations with mocks and do not restore them,
    which breaks ``Presentation(stream)`` reopen in save/reopen round-trip
    assertions.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.embeddedpackage import EmbeddedXlsxPart
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
        CT.SML_SHEET: EmbeddedXlsxPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


def _absjoin(*args: str) -> str:
    from .unitutil.file import absjoin, test_file_dir

    return absjoin(test_file_dir, *args)


# -- canonical reporter text: the industry-experience slide marker used in
# -- the #696 issue body. kept verbatim (including the Spanish accent) so the
# -- test pins the exact substring the reporter was scanning for. --
REPORTER_MARKER = "Nuestra experiencia en su industria"


def _add_text_slide(prs: PresentationClass, text: str, layout_index: int = 5) -> Slide:
    """Append a slide to `prs` with a textbox carrying `text`, return it."""
    slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    p = tb.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(24)
    return slide


def _slide_contains(slide: Slide, substring: str) -> bool:
    """Mirror the reporter's shapes/text_frame/paragraphs/runs traversal."""
    for shape in slide.shapes:
        # -- only shape subclasses that carry a text frame (TextFrame,
        # -- autoshape, textbox, placeholder) have ``has_text_frame`` == True;
        # -- gate the text_frame attribute access behind that check. --
        if not getattr(shape, "has_text_frame", False):
            continue
        text_frame = cast(
            "TextFrame",
            shape.text_frame,  # pyright: ignore[reportAttributeAccessIssue]
        )
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if substring in run.text:
                    return True
    return False


class DescribeIssue696MoveSlideCrossPptx(object):
    """#696 — close as resolved by #1036 + #934.

    Pins the reporter's "move a slide from one .pptx to another" workflow
    end-to-end, including text-keyed slide location, optional reposition and
    delete-to-replace, and save/reopen round-trip.
    """

    # -- (1) the headline #696 scenario: text-keyed graft from ppt2 into ppt1 --

    def it_clones_a_text_keyed_slide_from_another_presentation(self, _restore_part_factory):
        """The exact reporter-style workflow, mutatis mutandis.

        Build a "library" deck (``ppt2.pptx`` in the issue body) whose
        slide[0] carries the reporter's industry-experience marker text.
        Build a "target" deck (``ppt1.pptx``) with a couple of placeholder
        slides. Scan the target for the marker (it is *not* there — the
        reporter's ppt2 has the slide they want), then clone slide[0] of the
        library into the target. The cloned slide must appear in the target
        slide sequence with its text content intact.
        """
        library = Presentation()
        _add_text_slide(library, REPORTER_MARKER)

        target = Presentation()
        _add_text_slide(target, "Intro slide")
        _add_text_slide(target, "Outro slide")
        target_pre_count = len(target.slides)

        # -- reporter's traversal: scan the target first --
        assert not any(_slide_contains(sl, REPORTER_MARKER) for sl in target.slides)

        # -- graft the library's industry-experience slide into the target --
        source_slide = library.slides[0]
        cloned = target.slides.add_slide_from_external(source_slide, target.slide_layouts[5])

        assert cloned in list(target.slides)
        assert len(target.slides) == target_pre_count + 1
        # -- the marker now appears in the target, on the cloned slide --
        assert _slide_contains(cloned, REPORTER_MARKER)
        assert any(_slide_contains(sl, REPORTER_MARKER) for sl in target.slides)

    # -- (2) reposition + delete: graft-and-replace an existing slide -------

    def it_replaces_a_target_slide_via_add_external_plus_move_and_delete(
        self, _restore_part_factory
    ):
        """Full "move a slide from A into B in place of an existing B slide".

        The reporter's pseudo-code does ``slide = slideAn`` in a loop — the
        intent is "put the imported slide where this one was". The shipping
        recipe composes :meth:`Slides.add_slide_from_external` +
        :meth:`Slides.move_slide` + :meth:`Slides.delete` to achieve exactly
        that end-state: the imported slide occupies the replaced slide's
        former position, and the old slide is gone.
        """
        library = Presentation()
        _add_text_slide(library, REPORTER_MARKER)

        target = Presentation()
        _add_text_slide(target, "Intro")
        replaced = _add_text_slide(target, "PLACEHOLDER to swap out")
        _add_text_slide(target, "Outro")
        replaced_idx = target.slides.index(replaced)

        # -- graft the library slide, reposition to the replaced slot,
        # -- then delete the placeholder. --
        cloned = target.slides.add_slide_from_external(library.slides[0], target.slide_layouts[5])
        target.slides.move_slide(cloned, replaced_idx)
        target.slides.delete(replaced)

        # -- final ordering: [Intro, <imported marker slide>, Outro] --
        assert len(target.slides) == 3
        assert _slide_contains(target.slides[replaced_idx], REPORTER_MARKER)
        assert _slide_contains(target.slides[0], "Intro")
        assert _slide_contains(target.slides[2], "Outro")
        # -- the placeholder text is gone from the deck entirely --
        assert not any(_slide_contains(sl, "PLACEHOLDER to swap out") for sl in target.slides)

    # -- (3) save + reopen round-trip with a picture-carrying external slide --

    def it_round_trips_a_picture_bearing_external_copy(self, _restore_part_factory):
        """The #696 reporter's decks in practice will carry pictures.

        Pin that a slide with a picture shape clones across presentations
        and the picture survives the save + reopen cycle — exercising the
        Foundation F1 image-part materialisation that #934 + #1036 rely on.
        """
        library = Presentation()
        lib_slide = library.slides.add_slide(library.slide_layouts[5])
        lib_slide.shapes.add_picture(
            _absjoin("python-icon.jpeg"), Inches(1), Inches(1), height=Inches(1)
        )

        target = Presentation()
        target.slides.add_slide_from_external(lib_slide, target.slide_layouts[5])

        buf = io.BytesIO()
        target.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)

        assert len(reopened.slides) == 1
        assert any(sh.shape_type == MSO_SHAPE_TYPE.PICTURE for sh in reopened.slides[0].shapes)

    # -- (4) source deck not mutated — reporter's "pull from library" usage --

    def it_does_not_mutate_the_source_presentation(self, _restore_part_factory):
        """The #696 reporter opens ``ppt2.pptx`` as the library deck and
        pulls slides out of it. The library must survive intact so the same
        deck can be re-used as a source for further grafts.
        """
        library = Presentation()
        _add_text_slide(library, REPORTER_MARKER)
        _add_text_slide(library, "Other library slide")
        source_count_before = len(library.slides)
        source_shape_counts_before = [len(sl.shapes) for sl in library.slides]

        target = Presentation()
        target.slides.add_slide_from_external(library.slides[0], target.slide_layouts[5])

        assert len(library.slides) == source_count_before
        assert [len(sl.shapes) for sl in library.slides] == source_shape_counts_before
        # -- the library slide's text marker is still present in the library --
        assert _slide_contains(library.slides[0], REPORTER_MARKER)

    # -- (5) the alternate one-shot path: merge the whole library ------------

    def it_also_supports_merge_for_the_whole_library_at_once(self, _restore_part_factory):
        """#934's :meth:`Presentation.merge` is the shortcut when the
        reporter wants every slide of ``ppt2.pptx`` appended to ``ppt1.pptx``
        in one call, rather than per-slide text-keyed grafting.
        """
        library = Presentation()
        _add_text_slide(library, REPORTER_MARKER)
        _add_text_slide(library, "Another library slide")

        target = Presentation()
        _add_text_slide(target, "Intro")
        target_pre = len(target.slides)

        appended = target.merge(library)

        assert len(appended) == len(library.slides)
        assert len(target.slides) == target_pre + len(library.slides)
        # -- library's marker slide landed in the merged target --
        assert any(_slide_contains(sl, REPORTER_MARKER) for sl in target.slides)

    # -- (6) self-copy is explicitly permitted by the docstring --------------

    def it_permits_cloning_a_slide_within_the_same_presentation(self, _restore_part_factory):
        """``add_slide_from_external`` allows ``source_slide`` to come from
        any Presentation including self — a property the reporter's loop
        implicitly relies on when ``prs`` and ``prsIndustria`` point at the
        same file path.
        """
        prs = Presentation()
        original = _add_text_slide(prs, REPORTER_MARKER)
        pre_count = len(prs.slides)

        cloned = prs.slides.add_slide_from_external(original, prs.slide_layouts[5])

        assert cloned is not original
        assert len(prs.slides) == pre_count + 1
        assert _slide_contains(cloned, REPORTER_MARKER)
        # -- the original is untouched; the marker still appears in it --
        assert _slide_contains(original, REPORTER_MARKER)
