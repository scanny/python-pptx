# pyright: reportPrivateUsage=false, reportUnknownVariableType=false
# pyright: reportUnknownMemberType=false, reportAttributeAccessIssue=false

"""Regression test for issue #824 — add a picture to a slide master.

Issue #824 (https://github.com/scanny/python-pptx/issues/824) asks whether a
picture can be appended to a :class:`~pptx.slide.SlideMaster` so that every
slide inheriting from that master displays the shared image — typically a
client logo or watermark that should appear on every slide in the
presentation. This is the master-level sibling of issue #1044 (layout-level
textbox authoring) and was delivered by the fork-era resolution of issue
#575, which promoted ``LayoutShapes`` and ``MasterShapes`` to subclass
``_BaseGroupShapes`` and therefore expose the full shape-authoring surface —
``add_shape``, ``add_picture``, ``add_textbox``, ``add_connector``,
``add_group_shape`` and ``build_freeform``.

Cross-reference: issue #1044 (layout-textbox verify) pinned the same
class-hierarchy contract from the layout angle; this module pins it from
the master-picture angle. If a future refactor accidentally un-promotes
``MasterShapes`` from ``_BaseGroupShapes``, both suites will fail loudly.
"""

from __future__ import annotations

import io
import zipfile

from pptx import Presentation
from pptx.shapes.picture import Picture
from pptx.shapes.shapetree import (
    MasterShapes,
    _BaseGroupShapes,
)
from pptx.util import Inches

IMAGE_PATH = "tests/test_files/python-powered.png"


class DescribeMasterShapesAddPicture(object):
    """Master-picture authoring contract for issue #824."""

    def it_is_promoted_via_the_BaseGroupShapes_inheritance(self):
        # -- sentinel: confirms MasterShapes still inherits the full --
        # -- authoring surface from _BaseGroupShapes; if a refactor --
        # -- regresses this, add_picture silently disappears. --
        assert issubclass(MasterShapes, _BaseGroupShapes)
        assert hasattr(MasterShapes, "add_picture")
        assert hasattr(MasterShapes, "add_textbox")
        assert hasattr(MasterShapes, "add_shape")

    def it_appends_a_Picture_to_the_masters_shape_tree(self):
        prs = Presentation()
        master = prs.slide_master
        original_count = len(master.shapes)

        pic = master.shapes.add_picture(IMAGE_PATH, Inches(1), Inches(1), Inches(2), Inches(2))

        assert isinstance(pic, Picture)
        assert len(master.shapes) == original_count + 1
        assert master.shapes[-1].shape_id == pic.shape_id

    def it_survives_save_and_reopen(self):
        prs = Presentation()
        master = prs.slide_master
        master.shapes.add_picture(IMAGE_PATH, Inches(1), Inches(1), Inches(2), Inches(2))

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        reopened = Presentation(buf)
        pics = [s for s in reopened.slide_master.shapes if isinstance(s, Picture)]
        assert len(pics) == 1

    def it_writes_the_picture_into_the_master_part_not_a_slide(self):
        """The picture is owned by the master XML, not any slide."""
        prs = Presentation()
        master = prs.slide_master
        master.shapes.add_picture(IMAGE_PATH, Inches(1), Inches(1), Inches(2), Inches(2))

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        with zipfile.ZipFile(buf) as zf:
            master_xml = zf.read("ppt/slideMasters/slideMaster1.xml").decode("utf-8")
            # -- `p:pic` identifies a picture shape in the master spTree --
            assert "p:pic" in master_xml or "<pic>" in master_xml
            # -- there should be an image part registered with the master --
            master_rels = zf.read("ppt/slideMasters/_rels/slideMaster1.xml.rels").decode("utf-8")
            assert "image" in master_rels

    def it_is_visible_on_every_slide_that_inherits_the_master(self):
        """Pictures on the master are rendered onto every inheriting slide.

        Inheritance is rendering-time: the slide's ``spTree`` stays empty,
        but PowerPoint composes the master's shapes onto every slide that
        uses a layout derived from that master. This test asserts the
        invariant from the API side — the master owns exactly one picture,
        and the slides inherit-by-reference (their own shape collections
        are unchanged).
        """
        prs = Presentation()
        master = prs.slide_master
        master.shapes.add_picture(IMAGE_PATH, Inches(1), Inches(1), Inches(2), Inches(2))

        # -- build several slides from different layouts, all of which --
        # -- descend from the same master. --
        slides = [
            prs.slides.add_slide(prs.slide_layouts[0]),
            prs.slides.add_slide(prs.slide_layouts[1]),
            prs.slides.add_slide(prs.slide_layouts[5]),
        ]

        # -- the master's picture count is unchanged --
        master_pics = [s for s in master.shapes if isinstance(s, Picture)]
        assert len(master_pics) == 1

        # -- inheriting slides don't duplicate the picture into their own --
        # -- spTree; rendering-time inheritance handles display. --
        for slide in slides:
            slide_pics = [s for s in slide.shapes if isinstance(s, Picture)]
            assert len(slide_pics) == 0

        # -- round-trip: after save/reopen the master still owns the one --
        # -- picture, and every layout still points back at that master. --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        reopened_master_pics = [s for s in reopened.slide_master.shapes if isinstance(s, Picture)]
        assert len(reopened_master_pics) == 1
        # -- every reopened slide descends from the same master --
        for slide in reopened.slides:
            assert slide.slide_layout.slide_master is reopened.slide_master

    def it_accepts_a_png_file_path_for_master_branding(self):
        """Happy-path smoke test mirroring the reporter's use case."""
        prs = Presentation()
        pic = prs.slide_master.shapes.add_picture(
            IMAGE_PATH, Inches(0), Inches(0), Inches(1), Inches(1)
        )
        image = pic.image
        assert image is not None
        assert image.content_type == "image/png"
