# pyright: reportPrivateUsage=false

"""FU-6 — ``_BaseShapes.clear()`` + save GC composition regression test.

Wave 15 #96 added :meth:`SlideShapes.clear` /
:meth:`Slide.clear_shapes` for bulk removal of a slide's shapes while
preserving its placeholders. Wave 12 #956 verified that
:meth:`OpcPackage.save` walks the rels graph from the package root via
:meth:`~pptx.opc.package.OpcPackage.iter_parts` and silently drops parts
no longer reachable — giving the library save-time garbage collection
for "free" when a shape's ``.delete()`` override also drops the shape's
slide-part relationships.

FU-6 pins that the *two* features compose correctly: calling
``slide.clear_shapes()`` on a deck of picture-bearing slides and re-
saving should produce a meaningfully smaller zip because each removed
:class:`Picture` drops its image rel from the slide part, making the
image part unreachable on the next save.

A companion test closes the chart-GC loop surfaced while writing the
original FU-6 regression: :class:`GraphicFrame` now overrides
:meth:`BaseShape.delete` (FU-7) and drops every slide-part rel it
carries — classic ``c:chart`` / extended ``cx:chart``, embedded or
linked ``p:oleObj``, OLE-icon ``a:blip`` embed, the four
``dgm:relIds`` SmartArt rIds, and the ``am3d:model3D`` embed —
letting the save-time GC reap the no-longer-reachable chart /
embedded xlsx / OLE / diagram / 3D parts. The test in
:meth:`DescribeFU6ClearShapesGC.it_gcs_chart_parts_on_clear_shapes`
pins that the chart part and its embedded xlsx are both gone from
the saved zip after ``clear_shapes()``.

Cross-references:

* Wave 15 #96 — ``SlideShapes.clear()`` / ``Slide.clear_shapes()``
* Wave 12 #956 — save-time ``iter_parts()`` reachability GC
* FU-7 — ``GraphicFrame.delete()`` drops chart / chartex / OLE /
  SmartArt / 3D-model rels
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.parts.image import ImagePart
from pptx.util import Inches

_TEST_FILES = Path(__file__).parent / "test_files"

_PICTURE_NAMES = ("python-powered.png", "monty-truth.png", "python-icon.jpeg")


def _sha1_of_file(path: Path) -> str:
    """Return the hex SHA1 digest of the bytes at ``path``."""
    return hashlib.sha1(path.read_bytes()).hexdigest()


def _image_shas(prs):
    """Return the set of image-part SHA1s reachable from ``prs``."""
    return {p.sha1 for p in prs.part.package.iter_parts() if isinstance(p, ImagePart)}


def _zip_entries(buf):
    """Return the list of zip-member names in the saved package ``buf``."""
    buf.seek(0)
    with zipfile.ZipFile(buf) as zf:
        return zf.namelist()


def _build_picture_deck():
    """Return a Presentation with three picture-bearing slides (blank layout).

    Uses the blank layout so the only images reachable from the deck at
    save time are the ones we add explicitly (the default template's
    title layout carries a theme image and the deck itself ships a
    ``/docProps/thumbnail.jpeg`` — using the blank layout keeps the
    SHA-set assertion tight).
    """
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for img in _PICTURE_NAMES:
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(str(_TEST_FILES / img), Inches(1), Inches(1), Inches(3), Inches(3))
    return prs


class DescribeFU6ClearShapesGC:
    """FU-6: ``clear_shapes()`` + ``save()`` composes with iter_parts() GC."""

    # -- primary pin: pictures are GC'd and the zip shrinks ---------------

    def it_shrinks_the_saved_zip_after_clearing_picture_shapes(self):
        # -- build a picture-bearing deck and save it --
        prs = _build_picture_deck()
        before_buf = io.BytesIO()
        prs.save(before_buf)
        before_size = len(before_buf.getvalue())

        # -- clear every slide (keep placeholders) and save again --
        for slide in prs.slides:
            slide.clear_shapes()
        after_buf = io.BytesIO()
        prs.save(after_buf)
        after_size = len(after_buf.getvalue())

        # -- the three embedded PNG/JPEG parts are ~60 KB combined;
        # -- require a conservative 1 KB shrink so the pin is robust to
        # -- harmless serialization-overhead changes while still failing
        # -- loudly if image GC regresses. --
        shrink = before_size - after_size
        assert shrink >= 1024, (
            f"expected save-after-clear to shrink zip by at least 1 KB; "
            f"got before={before_size} after={after_size} delta={shrink}"
        )

    def it_removes_orphaned_image_parts_from_the_saved_zip(self):
        # -- Pin the reporter-facing invariant: after clear_shapes() +
        # -- save(), none of the deleted pictures' ppt/media/* entries
        # -- remain in the zip and none of our three picture SHAs
        # -- survive into the reopened package's ImagePart set.
        # --
        # -- The bare default template ships a ``/docProps/thumbnail.jpeg``
        # -- so the post-clear ImagePart set is not empty — we assert
        # -- specifically on the three added picture SHAs. --
        added_shas = {_sha1_of_file(_TEST_FILES / name) for name in _PICTURE_NAMES}
        prs = _build_picture_deck()
        # -- all three added pictures are reachable pre-clear --
        assert added_shas.issubset(_image_shas(prs))

        for slide in prs.slides:
            slide.clear_shapes()

        buf = io.BytesIO()
        prs.save(buf)

        # -- no ppt/media/* entries in the saved zip --
        media_entries = [n for n in _zip_entries(buf) if n.startswith("ppt/media/")]
        assert media_entries == []

        # -- and none of our three SHAs survive into the reopened package --
        buf.seek(0)
        prs2 = Presentation(buf)
        assert added_shas.isdisjoint(_image_shas(prs2))

    def it_preserves_placeholders_so_layout_inheritance_is_kept(self):
        # -- Belt-and-braces: clear_shapes() default keeps placeholders,
        # -- so a reopened deck still renders the title layout's
        # -- placeholder shapes (the whole point of the default).
        prs = Presentation()
        title_layout = prs.slide_layouts[0]
        for img in _PICTURE_NAMES:
            slide = prs.slides.add_slide(title_layout)
            slide.shapes.add_picture(
                str(_TEST_FILES / img), Inches(1), Inches(1), Inches(3), Inches(3)
            )

        for slide in prs.slides:
            slide.clear_shapes()

        buf = io.BytesIO()
        prs.save(buf)

        buf.seek(0)
        prs2 = Presentation(buf)
        for slide in prs2.slides:
            # -- each slide retains exactly its layout's placeholder set --
            assert len(list(slide.placeholders)) == 2
            # -- and no non-placeholder shapes are left behind --
            non_ph = [s for s in slide.shapes if not s.is_placeholder]
            assert non_ph == []

    def it_does_not_shrink_when_every_slide_has_no_deletable_shapes(self):
        # -- Inverse invariant: if clear_shapes() has nothing to remove
        # -- (placeholders-only slides, preserve_placeholders=True), the
        # -- zip shouldn't meaningfully change in size. This guards
        # -- against accidentally attributing any shrink to something
        # -- other than orphan-part GC.
        prs = Presentation()
        title_layout = prs.slide_layouts[0]
        for _ in range(3):
            prs.slides.add_slide(title_layout)

        before_buf = io.BytesIO()
        prs.save(before_buf)
        before_size = len(before_buf.getvalue())

        for slide in prs.slides:
            slide.clear_shapes()  # default preserve_placeholders=True

        after_buf = io.BytesIO()
        prs.save(after_buf)
        after_size = len(after_buf.getvalue())

        # -- no non-placeholder shapes to remove -> no meaningful shrink --
        assert abs(before_size - after_size) < 512

    # -- charts: FU-7 closes the GraphicFrame.delete() gap --------------

    def it_gcs_chart_parts_on_clear_shapes(self):
        # -- FU-7: :class:`GraphicFrame` now overrides
        # -- :meth:`BaseShape.delete` to drop every slide-part rel the
        # -- frame holds (classic chart, extended chartex, OLE object,
        # -- OLE icon image, SmartArt four-rIds, and 3D-model embed).
        # -- Clearing a chart-bearing slide therefore drops the slide ->
        # -- chart rel, and :meth:`OpcPackage.save` -> iter_parts()
        # -- reachability walk prunes the chart part + its embedded
        # -- xlsx dependency on the next save.
        prs = Presentation()
        layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(layout)
        chart_data = CategoryChartData()
        chart_data.categories = ("a", "b", "c")
        chart_data.add_series("series", (1, 2, 3))
        slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(4),
            chart_data,
        )

        slide.clear_shapes()

        buf = io.BytesIO()
        prs.save(buf)

        # -- the chart part and its embedded xlsx are both gone from the
        # -- zip: clear_shapes() dropped the slide -> chart rel, making
        # -- both parts unreachable from iter_parts() at save time. --
        entries = _zip_entries(buf)
        assert not any("ppt/charts/chart" in n for n in entries), (
            "expected chart part to be GC'd after clear_shapes(); "
            f"got entries: {[n for n in entries if 'chart' in n]}"
        )
        assert not any("ppt/embeddings/" in n for n in entries), (
            "expected embedded xlsx to be GC'd after clear_shapes(); "
            f"got entries: {[n for n in entries if 'embed' in n]}"
        )
        # -- and the slide's rels file no longer points at the chart --
        buf.seek(0)
        with zipfile.ZipFile(buf) as zf:
            slide_rels = zf.read("ppt/slides/_rels/slide1.xml.rels").decode()
        assert "charts/chart" not in slide_rels

    # -- mixed deck: partial GC --------------------------------------------

    def it_gcs_pictures_and_charts_from_a_mixed_deck(self):
        # -- FU-7 closes the chart-GC gap: in a mixed deck (pictures +
        # -- chart), both the picture parts AND the chart / embedded
        # -- xlsx parts are now GC'd after clear_shapes() + save. --
        prs = Presentation()
        blank = prs.slide_layouts[6]
        pic_slide = prs.slides.add_slide(blank)
        pic_slide.shapes.add_picture(str(_TEST_FILES / "python-powered.png"), Inches(1), Inches(1))
        chart_slide = prs.slides.add_slide(blank)
        chart_data = CategoryChartData()
        chart_data.categories = ("a", "b")
        chart_data.add_series("series", (1, 2))
        chart_slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(4),
            chart_data,
        )

        for slide in prs.slides:
            slide.clear_shapes(preserve_placeholders=False)

        buf = io.BytesIO()
        prs.save(buf)

        entries = _zip_entries(buf)
        # -- picture part GC'd -> no ppt/media/* entries --
        assert [n for n in entries if n.startswith("ppt/media/")] == []
        # -- chart part + xlsx GC'd -> no ppt/charts/* or ppt/embeddings/* entries --
        assert not any("ppt/charts/chart" in n for n in entries)
        assert not any("ppt/embeddings/" in n for n in entries)
