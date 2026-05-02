# pyright: reportPrivateUsage=false

"""Regression test for issue #941 — "Unable to open generated presentation".

Issue #941 (https://github.com/scanny/python-pptx/issues/941) reports that
the quickstart snippet from the user guide — a ``Presentation()``, the
layout-0 ``add_slide`` + ``title``/``placeholders[1]`` pattern copied
verbatim from ``docs/user/quickstart.rst`` — produced a ``.pptx`` that
Microsoft PowerPoint, PowerPoint-online, and Google Slides all refused to
open. The reporter's follow-up comment narrowed the trigger to "any
textbox being present" on top of the quickstart snippet. No traceback,
file, or environment detail was provided.

A second contributor (MartinPacker) ran the exact same snippet on macOS
and PowerPoint opened the output cleanly, strongly suggesting an
environment-local failure — a corrupted/partial install of the wheel, a
Windows CRLF mangling of the uploaded ``.pptx``, a zip re-compressor in
the user's transfer chain, or antivirus quarantining the file. No bug in
``pptx`` has ever been identified that matches the symptom.

#941 is therefore verified-and-closed as non-reproducible. This suite
pins the two code paths the reporter specifically called out — the
verbatim quickstart snippet, and the same snippet with a ``add_textbox``
— as producing well-formed ``.pptx`` packages that survive save + reopen
through ``Presentation``. Any future regression that *does* break the
quickstart snippet (a corrupted ``templates/default.pptx``, a content-
types registry that drops a required part, a relationship graph with a
dangling target, a zip layout PowerPoint's reader rejects) would
reproduce the #941 symptom and be caught here.

The tests deliberately exercise the same sequence of public-API calls
the reporter typed, not a synthetic equivalent, so the regression guard
matches the user's mental model of "the quickstart example".
"""

from __future__ import annotations

import io
import zipfile

import pytest

from pptx import Presentation
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of ``PartFactory.part_type_for``.

    Some earlier test modules mock ``PartFactory.part_type_for`` entries
    without restoring them. ``Presentation(stream)`` reopen dispatches
    through that registry, so any pollution would make the round-trip
    tests here fail for reasons unrelated to #941. Mirrors the guard used
    by ``tests/test_issue_1013_placeholder_color_verify.py`` and friends.
    """
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


# ---- package-shape helpers ---------------------------------------------

# -- parts every .pptx PowerPoint will accept must carry (the OPC baseline) --
_REQUIRED_PARTS = frozenset(
    {
        "[Content_Types].xml",
        "_rels/.rels",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
        "ppt/slides/slide1.xml",
        "ppt/slides/_rels/slide1.xml.rels",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/theme/theme1.xml",
    }
)


def _assert_is_well_formed_pptx(data: bytes, *, expect_slide1: bool = True) -> zipfile.ZipFile:
    """Fail with a diagnostic if ``data`` is not a readable ``.pptx`` package.

    Cheap well-formedness checks matching the ones PowerPoint's loader
    applies before it even parses slide XML: the container must be a
    valid zip, every entry's CRC must match, and the required OPC parts
    must all be present. The #941 symptom (MS PowerPoint's generic
    "not valid" refusal) surfaces here if any of those baseline
    invariants are broken.

    ``expect_slide1=False`` drops ``ppt/slides/slide1.xml`` and its rels
    from the required set — a fresh ``Presentation()`` with no slides
    still has to emit a valid package (PowerPoint accepts slide-less
    decks) and we want to pin that too.
    """
    assert len(data) > 0, "saved pptx is empty"
    buf = io.BytesIO(data)
    # -- must be a recognisable zip --
    assert zipfile.is_zipfile(buf), "saved pptx is not a valid zip container"
    buf.seek(0)
    z = zipfile.ZipFile(buf)
    # -- every stored member's CRC must verify --
    bad = z.testzip()
    assert bad is None, f"zip member corrupt: {bad!r}"
    # -- OPC baseline: every part PowerPoint insists on --
    names = set(z.namelist())
    required = set(_REQUIRED_PARTS)
    if not expect_slide1:
        required.discard("ppt/slides/slide1.xml")
        required.discard("ppt/slides/_rels/slide1.xml.rels")
    missing = required - names
    assert not missing, f"missing required OPC parts: {sorted(missing)}"
    return z


class DescribeIssue941QuickstartOpensCleanly:
    """#941 unable-to-open verify-and-close suite.

    Pins that the verbatim quickstart snippet the reporter copied from
    the manual produces a well-formed ``.pptx`` package, both on its
    own and with the ``add_textbox`` variant the reporter flagged as
    "instantly corrupt". Verified non-reproducible; these tests guard
    against any future regression that would reintroduce the symptom.
    """

    # -- the exact snippet the #941 reporter pasted ---------------------

    def it_saves_the_quickstart_snippet_as_a_well_formed_pptx(self):
        """The verbatim quickstart example — the #941 reproducer.

        Same call sequence the manual publishes and the reporter copied:
        ``Presentation()`` → layout 0 ``add_slide`` → ``title.text`` +
        ``placeholders[1].text`` → ``save``. The saved bytes must be a
        valid zip whose members all CRC-verify and whose OPC baseline
        (``[Content_Types].xml``, ``_rels/.rels``, ``presentation.xml``,
        slide1 + layout1 + master1 + theme1) is intact.
        """
        # -- the #941 reporter's verbatim snippet --
        prs = Presentation()
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        assert title is not None  # -- layout 0 always has a title placeholder --
        subtitle = slide.placeholders[1]

        title.text = "Hello, World!"
        subtitle.text = "python-pptx was here!"

        buf = io.BytesIO()
        prs.save(buf)
        _assert_is_well_formed_pptx(buf.getvalue())

    def it_round_trips_the_quickstart_snippet_through_save_and_reopen(self, _restore_part_factory):
        """Save + reopen sanity on the #941 reproducer.

        Going through ``Presentation(stream)`` on the saved bytes
        exercises the same XML-parse path PowerPoint and LibreOffice
        hit. The reopened deck must expose the slide the reporter
        authored, with the title/subtitle text intact.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Hello, World!"
        slide.placeholders[1].text = "python-pptx was here!"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        assert len(prs2.slides) == 1
        reopened = prs2.slides[0]
        assert reopened.shapes.title is not None
        assert reopened.shapes.title.text == "Hello, World!"
        assert reopened.placeholders[1].text == "python-pptx was here!"

    # -- the "add a textbox" follow-up scenario -------------------------

    def it_saves_a_well_formed_pptx_when_a_textbox_is_added(self):
        """The #941 follow-up: quickstart + ``add_textbox``.

        The reporter's second comment claimed "when a textbox is present
        the file will instantly go corrupt". Adding an ``add_textbox``
        to the quickstart slide must still produce a well-formed package
        — that's the exact claim this regression test pins.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Hello, World!"
        slide.placeholders[1].text = "python-pptx was here!"

        # -- the #941 follow-up trigger --
        textbox = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(3), Inches(1))
        textbox.text_frame.text = "a textbox"

        buf = io.BytesIO()
        prs.save(buf)
        _assert_is_well_formed_pptx(buf.getvalue())

    def it_round_trips_the_textbox_variant_through_save_and_reopen(self, _restore_part_factory):
        """Save + reopen sanity on the textbox variant.

        The quickstart-with-textbox deck must survive
        ``Presentation(stream)`` reopen; the textbox shape and its text
        must be present on the reloaded slide.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Hello, World!"
        slide.placeholders[1].text = "python-pptx was here!"
        textbox = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(3), Inches(1))
        textbox.text_frame.text = "a textbox"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        reopened = prs2.slides[0]
        # -- 3 shapes on the slide: title placeholder, body placeholder, textbox --
        assert len(reopened.shapes) == 3
        # -- textbox text survived --
        texts = {sh.text_frame.text for sh in reopened.shapes if sh.has_text_frame}
        assert "a textbox" in texts
        assert "Hello, World!" in texts
        assert "python-pptx was here!" in texts

    # -- minimal-deck sanity checks -------------------------------------

    def it_saves_an_empty_presentation_as_a_well_formed_pptx(self):
        """Bare ``Presentation()`` + ``save`` — the OPC baseline case.

        A regression at the zip/content-type/rels layer would show up
        here first; a ``Presentation()`` with no slides still has to
        emit a valid package (PowerPoint accepts slide-less decks).
        """
        prs = Presentation()
        buf = io.BytesIO()
        prs.save(buf)
        data = buf.getvalue()
        _assert_is_well_formed_pptx(data, expect_slide1=False)

    def it_saves_multiple_slides_with_mixed_shape_kinds(self):
        """Broader shape coverage: multi-slide deck with mixed shape kinds.

        Catches regressions that would break specifically on the second
        slide, on layouts other than layout 0, or on the combination of
        placeholder content + add_textbox on the same slide.
        """
        prs = Presentation()
        # -- slide 1: the reporter's snippet --
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Hello, World!"
        s1.placeholders[1].text = "python-pptx was here!"
        # -- slide 2: bullet layout + textbox + body text --
        s2 = prs.slides.add_slide(prs.slide_layouts[1])
        s2.shapes.title.text = "Slide 2"
        s2.placeholders[1].text_frame.text = "Body content"
        s2.shapes.add_textbox(Inches(1), Inches(5), Inches(3), Inches(0.5)).text_frame.text = (
            "Floating textbox"
        )
        # -- slide 3: blank layout + just a textbox --
        s3 = prs.slides.add_slide(prs.slide_layouts[6])
        s3.shapes.add_textbox(Inches(2), Inches(3), Inches(5), Inches(1)).text_frame.text = (
            "Lone textbox"
        )

        buf = io.BytesIO()
        prs.save(buf)
        data = buf.getvalue()
        z = _assert_is_well_formed_pptx(data)
        names = set(z.namelist())
        # -- three slide parts and three matching rels --
        for i in range(1, 4):
            assert f"ppt/slides/slide{i}.xml" in names
            assert f"ppt/slides/_rels/slide{i}.xml.rels" in names
