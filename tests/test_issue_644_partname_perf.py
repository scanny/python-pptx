"""Regression test for issue #644 — O(N**2) partname allocation.

Issue #644 (https://github.com/scanny/python-pptx/issues/644) reports
that creating a large presentation (many hundreds of slides / images /
notes) is extremely slow because every new part triggers a full walk of
the existing part graph to pick the next available partname. That makes
each allocation O(N) in the number of existing parts and overall
deck-construction O(N**2).

The fix caches the set of partnames already handed out per template on
the ``OpcPackage`` so subsequent allocations for the same template are
O(1). This test guards the fix along two axes:

* **Functional** — a freshly-constructed package can hand out thousands
  of partnames for a mix of templates and each one is unique and follows
  the expected numbering; partnames allocated for one template do not
  collide with partnames allocated for another template that happens to
  share a prefix.
* **Performance-smoke** — a timing check that authoring a presentation
  with many slides completes well under an upper-bound threshold on any
  reasonable machine. This runs only when the environment variable
  ``RUN_PERF_TESTS=1`` is set so it doesn't slow CI, but is available
  locally to catch any regression that reintroduces the O(N**2) scan.
"""

from __future__ import annotations

import io
import os
import time

import pytest

from pptx import Presentation
from pptx.opc.package import OpcPackage
from pptx.opc.packuri import PackURI


class DescribeIssue644PartnameAllocatorCorrectness:
    """Functional correctness of the cached partname allocator."""

    def it_returns_unique_partnames_across_many_calls_for_one_tmpl(self):
        """The cached allocator hands out a contiguous, unique sequence."""
        package = OpcPackage(None)
        tmpl = "/ppt/slides/slide%d.xml"

        partnames = [package.next_partname(tmpl) for _ in range(500)]

        # -- every allocation is unique (the core invariant the packaging
        # -- spec requires: no two parts may share a partname) --
        assert len(set(partnames)) == 500
        # -- and forms the expected contiguous 1..N sequence --
        assert partnames == [PackURI(tmpl % n) for n in range(1, 501)]

    def it_keeps_partname_streams_independent_across_tmpls(self):
        """Distinct templates never collide, even with shared prefixes."""
        package = OpcPackage(None)
        slide_tmpl = "/ppt/slides/slide%d.xml"
        notes_tmpl = "/ppt/notesSlides/notesSlide%d.xml"
        image_tmpl = "/ppt/media/image%d.png"

        # -- interleave allocations across templates to exercise the
        # -- per-template cache --
        allocated: list[PackURI] = []
        for _ in range(200):
            allocated.append(package.next_partname(slide_tmpl))
            allocated.append(package.next_partname(notes_tmpl))
            allocated.append(package.next_partname(image_tmpl))

        # -- no duplicates across the combined stream --
        assert len(set(allocated)) == 600
        # -- and each template's own subsequence is the expected 1..200 --
        slide_ns = [allocated[i] for i in range(0, 600, 3)]
        notes_ns = [allocated[i] for i in range(1, 600, 3)]
        image_ns = [allocated[i] for i in range(2, 600, 3)]
        assert slide_ns == [PackURI(slide_tmpl % n) for n in range(1, 201)]
        assert notes_ns == [PackURI(notes_tmpl % n) for n in range(1, 201)]
        assert image_ns == [PackURI(image_tmpl % n) for n in range(1, 201)]

    def it_preserves_partname_uniqueness_end_to_end_when_adding_many_slides(self):
        """A real `Presentation` round-trip yields unique slide partnames."""
        prs = Presentation()
        layout = prs.slide_layouts[6]  # -- blank layout --

        for _ in range(50):
            prs.slides.add_slide(layout)

        # -- collect every slide-part partname in the saved package; each
        # -- one must be unique or PowerPoint will refuse to open the
        # -- file (this is the user-visible behavior #644 protects) --
        slide_partnames = [
            slide.part.partname for slide in prs.slides
        ]
        assert len(slide_partnames) == 50
        assert len(set(slide_partnames)) == 50

        # -- and the whole presentation still saves and re-opens cleanly --
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        reopened = Presentation(out)
        assert len(reopened.slides) == 50


@pytest.mark.skipif(
    os.environ.get("RUN_PERF_TESTS") != "1",
    reason="perf-smoke test; set RUN_PERF_TESTS=1 to run locally",
)
class DescribeIssue644PartnameAllocatorPerformance:
    """Performance-smoke guard against the O(N**2) regression."""

    def it_adds_many_slides_in_linear_time(self):
        """Adding 200 slides must complete well under two seconds.

        Pre-fix this took 5–10+ seconds on commodity hardware because
        each slide triggered a linear scan of all prior parts. Post-fix
        the operation is linear in the slide count. A 2-second ceiling
        leaves generous headroom for slow CI boxes while still catching
        any accidental reintroduction of the quadratic loop.
        """
        prs = Presentation()
        layout = prs.slide_layouts[6]

        t0 = time.perf_counter()
        for _ in range(200):
            prs.slides.add_slide(layout)
        elapsed = time.perf_counter() - t0

        assert elapsed < 2.0, f"200-slide build took {elapsed:.2f}s (expected <2s)"
