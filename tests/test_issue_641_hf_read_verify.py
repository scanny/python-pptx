# pyright: reportPrivateUsage=false

"""Regression test for issue #641 — "Read header/footer data from pptx".

Issue #641 (https://github.com/scanny/python-pptx/issues/641) asked for a
supported way to introspect a ``.pptx``'s header/footer configuration: both
the master-/layout-level visibility toggles encoded by the ``p:hf`` element
(slide number, header, footer, date on/off) and the actual text content of
the header/footer/date/slide-number placeholders on individual slides.

This fork resolves #641 without introducing a dedicated header/footer API
on slides — the reads compose from two independent APIs already shipped:

* Wave 1 #201 (``feat/issue-201-slide-header-footer``) landed
  :class:`pptx.slide._HeaderFooter` returned by
  :attr:`.SlideMaster.header_footer` and :attr:`.SlideLayout.header_footer`.
  The proxy's ``slide_number_visible`` / ``header_visible`` /
  ``footer_visible`` / ``date_visible`` getters read the ``sldNum`` /
  ``hdr`` / ``ftr`` / ``dt`` attributes off the ``p:hf`` child of the
  master or layout (defaulting to ``True`` when ``p:hf`` is absent, per
  the XSD default).

* The existing placeholder-access API — ``slide.placeholders[idx]`` and
  ``placeholder.text_frame.text`` — is all that's needed to read the
  *text content* of a footer/date/slide-number placeholder. Footer-type
  placeholders use the well-known ``idx`` values 10 (date), 11 (footer),
  and 12 (slide number) emitted by PowerPoint and by this library's
  built-in layouts.

#641 is therefore a "verify-and-close" item: no new code is needed, only
a regression suite that pins the two reads from end-user code exactly as
a reporter would write them. This module exercises both halves against a
round-tripped package so that a future refactor of either layer will fail
here before it breaks the documented recipe.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.slide import _HeaderFooter


@pytest.fixture
def pptx_with_header_footer_data():
    """Build a ``.pptx`` buffer exercising every #641 read path.

    The package is authored so that reads can discriminate each flag and
    each placeholder text channel:

    * Master ``p:hf``: ``date_visible = False`` (others inherit default
      ``True``).
    * Layout 0 ``p:hf``: ``slide_number_visible = False`` (others inherit
      default ``True``).
    * Slide 0: carries real ``p:sp`` footer / date / slide-number
      placeholders (idx 10 / 11 / 12) with distinct text bodies, in the
      same shape PowerPoint authors when the user checks "Footer" /
      "Date and time" / "Slide number" in Insert -> Header & Footer.

    Returns a ``BytesIO`` positioned at the start so each test opens a
    pristine ``Presentation`` from it.
    """
    prs = Presentation()

    # -- master-level p:hf: hide the date on every descendant layout/slide ---
    prs.slide_masters[0].header_footer.date_visible = False

    # -- layout-level p:hf: hide the slide number on the first layout only --
    prs.slide_layouts[0].header_footer.slide_number_visible = False

    # -- build a slide with real footer / date / slide-number placeholders --
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    spTree = slide.shapes._spTree
    max_id = max(
        (int(v) for v in spTree.xpath("//p:cNvPr/@id") if v.isdigit()),
        default=0,
    )
    # -- idx 10 = date, idx 11 = footer, idx 12 = slide number (PowerPoint
    # -- convention mirrored by the library's built-in layouts).
    spTree.add_placeholder(
        max_id + 1, "Date Placeholder 2", PP_PLACEHOLDER.DATE, "horz", "half", 10
    )
    spTree.add_placeholder(
        max_id + 2, "Footer Placeholder 3", PP_PLACEHOLDER.FOOTER, "horz", "quarter", 11
    )
    spTree.add_placeholder(
        max_id + 3,
        "Slide Number Placeholder 4",
        PP_PLACEHOLDER.SLIDE_NUMBER,
        "horz",
        "quarter",
        12,
    )

    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 10:
            ph.text_frame.text = "January 2026"
        elif idx == 11:
            ph.text_frame.text = "Confidential - Company Internal"
        elif idx == 12:
            ph.text_frame.text = "1"

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


class DescribeIssue641HeaderFooterRead(object):
    """#641 header/footer *read* paths — master/layout flags + slide text."""

    # -- master-level p:hf flag reads -------------------------------------

    def it_reads_all_master_hf_flags(self, pptx_with_header_footer_data):
        # -- pin the full four-flag surface of the master p:hf proxy, as the
        # -- reporter would inspect it after opening the file.
        prs = Presentation(pptx_with_header_footer_data)

        hf = prs.slide_masters[0].header_footer
        assert isinstance(hf, _HeaderFooter)
        # -- we explicitly set date_visible=False; the other three defaulted
        # -- to True on the master and must read back as True.
        assert hf.date_visible is False
        assert hf.slide_number_visible is True
        assert hf.header_visible is True
        assert hf.footer_visible is True

    # -- layout-level p:hf flag reads -------------------------------------

    def it_reads_all_layout_hf_flags(self, pptx_with_header_footer_data):
        # -- layout 0 had slide_number_visible flipped False explicitly; the
        # -- remaining flags inherit the "no p:hf => all True" default.
        prs = Presentation(pptx_with_header_footer_data)

        hf = prs.slide_layouts[0].header_footer
        assert isinstance(hf, _HeaderFooter)
        assert hf.slide_number_visible is False
        assert hf.header_visible is True
        assert hf.footer_visible is True
        assert hf.date_visible is True

    def it_returns_all_True_defaults_when_layout_has_no_hf(self, pptx_with_header_footer_data):
        # -- every other layout in the pack has no explicit p:hf; verify the
        # -- "absent => visible" default applies uniformly.
        prs = Presentation(pptx_with_header_footer_data)

        for i, layout in enumerate(prs.slide_layouts):
            if i == 0:
                continue  # layout 0 carries the explicit p:hf set in the fixture
            hf = layout.header_footer
            assert hf.slide_number_visible is True
            assert hf.header_visible is True
            assert hf.footer_visible is True
            assert hf.date_visible is True

    # -- slide-level placeholder text reads -------------------------------

    def it_reads_footer_placeholder_text_by_idx(self, pptx_with_header_footer_data):
        # -- the reporter's primary ask: given a saved pptx containing a
        # -- footer, pull the footer string. ``slide.placeholders[11]``
        # -- (idx=11 is the PowerPoint convention for footer) exposes the
        # -- placeholder, and ``.text_frame.text`` yields the string.
        prs = Presentation(pptx_with_header_footer_data)

        slide = prs.slides[0]
        footer_ph = slide.placeholders[11]

        assert footer_ph.placeholder_format.type == PP_PLACEHOLDER.FOOTER
        assert footer_ph.placeholder_format.idx == 11
        assert footer_ph.text_frame.text == "Confidential - Company Internal"

    def it_reads_date_placeholder_text_by_idx(self, pptx_with_header_footer_data):
        prs = Presentation(pptx_with_header_footer_data)

        date_ph = prs.slides[0].placeholders[10]

        assert date_ph.placeholder_format.type == PP_PLACEHOLDER.DATE
        assert date_ph.placeholder_format.idx == 10
        assert date_ph.text_frame.text == "January 2026"

    def it_reads_slide_number_placeholder_text_by_idx(self, pptx_with_header_footer_data):
        prs = Presentation(pptx_with_header_footer_data)

        slide_num_ph = prs.slides[0].placeholders[12]

        assert slide_num_ph.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER
        assert slide_num_ph.placeholder_format.idx == 12
        assert slide_num_ph.text_frame.text == "1"

    # -- the reporter's full recipe: flags AND text together --------------

    def it_supports_reading_flags_and_text_in_one_pass(self, pptx_with_header_footer_data):
        # -- pin the reporter's end-to-end read: given a pptx, harvest the
        # -- master/layout visibility profile *and* each slide's footer text
        # -- in one pass over the package. This is the minimal shape a
        # -- header/footer audit tool would take.
        prs = Presentation(pptx_with_header_footer_data)

        report = {
            "master": {
                "slide_number_visible": prs.slide_masters[0].header_footer.slide_number_visible,
                "header_visible": prs.slide_masters[0].header_footer.header_visible,
                "footer_visible": prs.slide_masters[0].header_footer.footer_visible,
                "date_visible": prs.slide_masters[0].header_footer.date_visible,
            },
            "layout_0": {
                "slide_number_visible": prs.slide_layouts[0].header_footer.slide_number_visible,
            },
            "slides": [],
        }
        for slide in prs.slides:
            ph_text = {}
            for ph in slide.placeholders:
                idx = ph.placeholder_format.idx
                if idx in (10, 11, 12):
                    ph_text[idx] = ph.text_frame.text
            report["slides"].append(ph_text)

        assert report == {
            "master": {
                "slide_number_visible": True,
                "header_visible": True,
                "footer_visible": True,
                "date_visible": False,
            },
            "layout_0": {"slide_number_visible": False},
            "slides": [
                {
                    10: "January 2026",
                    11: "Confidential - Company Internal",
                    12: "1",
                }
            ],
        }

    # -- idx-based iteration across all three hf placeholder kinds --------

    def it_iterates_hf_placeholder_text_by_type(self, pptx_with_header_footer_data):
        # -- belt-and-braces: iterate the slide's placeholders and discover
        # -- the hf ones by ``placeholder_format.type`` rather than by idx.
        # -- This is the shape a type-driven audit would take, and must also
        # -- reach every footer/date/slide-number placeholder on the slide.
        prs = Presentation(pptx_with_header_footer_data)

        hf_types = {
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.SLIDE_NUMBER,
        }
        collected: dict[PP_PLACEHOLDER, str] = {}
        for ph in prs.slides[0].placeholders:
            ph_type = ph.placeholder_format.type
            if ph_type in hf_types:
                collected[ph_type] = ph.text_frame.text

        assert collected == {
            PP_PLACEHOLDER.FOOTER: "Confidential - Company Internal",
            PP_PLACEHOLDER.DATE: "January 2026",
            PP_PLACEHOLDER.SLIDE_NUMBER: "1",
        }

    # -- public API surface ------------------------------------------------

    def it_exposes_header_footer_on_SlideMaster_and_SlideLayout(self):
        # -- pin the attribute surface the reporter actually needs exists.
        from pptx.slide import SlideLayout, SlideMaster

        assert hasattr(SlideMaster, "header_footer")
        assert hasattr(SlideLayout, "header_footer")
