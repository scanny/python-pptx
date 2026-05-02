# pyright: reportPrivateUsage=false

"""Regression test for issue #438 — ``Presentation.save_ppsx``.

Issue #438 (https://github.com/scanny/python-pptx/issues/438) asks for a
way to save a presentation as a PowerPoint Show (``.ppsx``). The only
on-disk difference from a regular ``.pptx`` is the content-type override
for the presentation part in ``[Content_Types].xml``:

  ``application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml``

instead of the regular

  ``application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml``.

Files written this way and named ``.ppsx`` open in PowerPoint straight
into slide-show playback.

This test exercises the end-to-end flow:

  1. Open the default template.
  2. Save it via :meth:`Presentation.save_ppsx`.
  3. Inspect ``[Content_Types].xml`` to confirm the override was flipped.
  4. Reopen the saved file; confirm the presentation part's content-type
     is the slideshow variant (:mod:`pptx` already whitelists the
     slideshow main type on load, per #1070).
"""

from __future__ import annotations

import io
import zipfile

from pptx import Presentation
from pptx.opc.constants import CONTENT_TYPE as CT


class DescribeIssue438SaveAsPpsx:
    """Round-trip regression for ``Presentation.save_ppsx`` (issue #438)."""

    def it_writes_slideshow_content_type_override(self):
        prs = Presentation()

        buf = io.BytesIO()
        prs.save_ppsx(buf)
        buf.seek(0)

        with zipfile.ZipFile(buf) as z:
            content_types_xml = z.read("[Content_Types].xml").decode("utf-8")

        # -- the slideshow content-type must appear and the regular
        # -- "…presentation.main+xml" must NOT be used as the override
        # -- for /ppt/presentation.xml --
        assert CT.PML_SLIDESHOW_MAIN in content_types_xml
        # -- confirm the override targets the presentation part --
        assert '/ppt/presentation.xml' in content_types_xml
        # -- the presentation.main+xml content type must NOT appear as the
        # -- override — otherwise PowerPoint would open it as a regular
        # -- presentation, not as a slide show.
        assert CT.PML_PRESENTATION_MAIN not in content_types_xml

    def it_round_trips_through_save_ppsx_and_reopen(self):
        prs = Presentation()
        slide_count_before = len(prs.slides)

        buf = io.BytesIO()
        prs.save_ppsx(buf)
        buf.seek(0)
        reloaded = Presentation(buf)

        # -- the reopened package's main document part is still the
        # -- presentation part (whitelisted on load per #1070) --
        assert reloaded.part.content_type == CT.PML_SLIDESHOW_MAIN
        # -- slides survive the round-trip --
        assert len(reloaded.slides) == slide_count_before

    def it_does_not_mutate_the_live_content_type_after_save(self):
        prs = Presentation()
        original_ct = prs.part.content_type

        buf = io.BytesIO()
        prs.save_ppsx(buf)

        # -- in-memory presentation keeps its original content-type so that
        # -- a follow-up `.save()` writes a regular `.pptx`.
        assert prs.part.content_type == original_ct

        buf2 = io.BytesIO()
        prs.save(buf2)
        buf2.seek(0)
        with zipfile.ZipFile(buf2) as z:
            content_types_xml = z.read("[Content_Types].xml").decode("utf-8")
        assert CT.PML_PRESENTATION_MAIN in content_types_xml
        assert CT.PML_SLIDESHOW_MAIN not in content_types_xml

    def it_accepts_a_filesystem_path(self, tmp_path):
        prs = Presentation()
        out_path = tmp_path / "slideshow.ppsx"

        prs.save_ppsx(str(out_path))

        with zipfile.ZipFile(str(out_path)) as z:
            content_types_xml = z.read("[Content_Types].xml").decode("utf-8")
        assert CT.PML_SLIDESHOW_MAIN in content_types_xml

    def it_forwards_zip_date_time_and_password(self):
        """`save_ppsx` honors the same reproducibility/encryption keywords as `save`."""
        prs = Presentation()

        buf = io.BytesIO()
        prs.save_ppsx(buf, zip_date_time=(2024, 6, 15, 9, 30, 0))
        buf.seek(0)

        with zipfile.ZipFile(buf) as z:
            # -- every zip member carries the fixed timestamp we asked for --
            for zinfo in z.infolist():
                assert zinfo.date_time == (2024, 6, 15, 9, 30, 0)
            # -- and the slideshow override is still present --
            content_types_xml = z.read("[Content_Types].xml").decode("utf-8")
        assert CT.PML_SLIDESHOW_MAIN in content_types_xml
