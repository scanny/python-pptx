# pyright: reportPrivateUsage=false, reportUnknownVariableType=false
# pyright: reportUnknownMemberType=false, reportAttributeAccessIssue=false

"""Regression test for issue #1044 — add a textbox to a layout.

Issue #1044 (https://github.com/scanny/python-pptx/issues/1044) asks whether a
textbox can be appended to a :class:`~pptx.slide.SlideLayout` so that every
slide inheriting from that layout displays the shared text (the layout-level
equivalent of ``slide.shapes.add_textbox(...)``). The capability was
delivered by the fork-era resolution of issue #575, which promoted
``LayoutShapes`` and ``MasterShapes`` to subclass ``_BaseGroupShapes`` and
therefore expose the full shape-authoring surface — ``add_shape``,
``add_picture``, ``add_textbox``, ``add_connector``, ``add_group_shape`` and
``build_freeform``.

This module pins the layout-textbox contract from the reporter's
perspective — a standalone regression that will fail loudly if a future
refactor accidentally un-promotes the authoring methods:

* ``layout.shapes.add_textbox(...)`` returns a ``Shape`` with a live
  ``TextFrame`` whose text is round-tripped through save + reopen.
* The textbox is appended to the layout's ``p:spTree``, not a slide —
  inspection of the unzipped package confirms it lives in
  ``ppt/slideLayouts/slideLayoutN.xml``.
* Slides built from that layout inherit the layout's shape count (the
  textbox stays on the layout; PowerPoint renders it onto every inheriting
  slide via layout-to-slide inheritance).
"""

from __future__ import annotations

import io
import zipfile
from typing import TYPE_CHECKING, Callable

import pytest

from pptx import Presentation
from pptx.shapes.autoshape import Shape
from pptx.shapes.shapetree import (
    LayoutShapes,
    MasterShapes,
    _BaseGroupShapes,
    _BaseShapes,
)
from pptx.util import Inches

if TYPE_CHECKING:
    from pptx.presentation import Presentation as _PresentationType


class DescribeLayoutShapesAddTextbox(object):
    """Layout-textbox authoring contract for issue #1044."""

    def it_is_promoted_via_the_BaseGroupShapes_inheritance(self):
        # -- a cheap sentinel: if a refactor regresses the class hierarchy --
        # -- the authoring methods silently disappear from LayoutShapes. --
        assert issubclass(LayoutShapes, _BaseGroupShapes)
        assert issubclass(MasterShapes, _BaseGroupShapes)
        assert hasattr(LayoutShapes, "add_textbox")
        assert hasattr(MasterShapes, "add_textbox")

    def it_appends_a_textbox_Shape_to_the_layouts_shape_tree(self):
        prs = Presentation()
        layout = prs.slide_layouts[0]
        original_count = len(layout.shapes)

        shape = layout.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))

        assert isinstance(shape, Shape)
        assert shape.has_text_frame
        assert len(layout.shapes) == original_count + 1
        assert layout.shapes[-1] is not None
        # -- same element reference whether accessed by index or by return --
        assert layout.shapes[-1].shape_id == shape.shape_id

    def it_accepts_text_on_the_returned_shapes_text_frame(self):
        prs = Presentation()
        layout = prs.slide_layouts[0]
        shape = layout.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        shape.text_frame.text = "CONFIDENTIAL — Draft"
        assert shape.text_frame.text == "CONFIDENTIAL — Draft"

    def it_survives_save_and_reopen(self):
        prs = Presentation()
        layout = prs.slide_layouts[0]
        layout_index = 0
        shape = layout.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        shape.text_frame.text = "Shared footer text"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        reopened = Presentation(buf)
        reopened_layout = reopened.slide_layouts[layout_index]
        textboxes = [
            s
            for s in reopened_layout.shapes
            if s.has_text_frame and s.text_frame.text == "Shared footer text"
        ]
        assert len(textboxes) == 1

    def it_writes_the_textbox_into_the_layout_part_not_a_slide(self):
        prs = Presentation()
        layout = prs.slide_layouts[0]
        layout.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text_frame.text = (
            "Shared footer text"
        )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        with zipfile.ZipFile(buf) as zf:
            layout_xml = zf.read("ppt/slideLayouts/slideLayout1.xml").decode("utf-8")
        assert "Shared footer text" in layout_xml

    def it_does_not_add_the_shape_to_inheriting_slides_shape_collection(self):
        """Inheritance is rendering-time; the slide's spTree stays untouched."""
        prs = Presentation()
        layout = prs.slide_layouts[0]
        layout.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text_frame.text = (
            "Shared footer text"
        )

        pre_slide_count = len(layout.shapes)
        slide = prs.slides.add_slide(layout)
        # -- the layout's shape-count is unchanged by creating a slide --
        assert len(layout.shapes) == pre_slide_count
        # -- the slide's spTree does not directly own the layout textbox --
        slide_texts = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
        assert "Shared footer text" not in slide_texts


class DescribeMasterShapesAddTextbox(object):
    """Master-textbox authoring contract — sibling of the layout case."""

    def it_can_add_a_textbox_to_a_slide_master(self):
        prs = Presentation()
        master = prs.slide_master
        original_count = len(master.shapes)

        shape = master.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        shape.text_frame.text = "All-slides banner"

        assert isinstance(shape, Shape)
        assert shape.has_text_frame
        assert len(master.shapes) == original_count + 1

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        master_texts = [s.text_frame.text for s in reopened.slide_master.shapes if s.has_text_frame]
        assert "All-slides banner" in master_texts


# -- parametrized: every shape-collection type that should expose add_textbox --


def _master_shapes(prs: _PresentationType) -> _BaseShapes:
    return prs.slide_master.shapes


def _layout_shapes(prs: _PresentationType) -> _BaseShapes:
    return prs.slide_layouts[0].shapes


def _slide_shapes(prs: _PresentationType) -> _BaseShapes:
    return prs.slides.add_slide(prs.slide_layouts[0]).shapes


@pytest.mark.parametrize(
    "collection_accessor",
    [_master_shapes, _layout_shapes, _slide_shapes],
    ids=["master", "layout", "slide"],
)
def it_exposes_add_textbox_on_every_slide_type_shape_tree(
    collection_accessor: Callable[[_PresentationType], _BaseShapes],
):
    prs = Presentation()
    shapes = collection_accessor(prs)
    original_count = len(shapes)
    # -- `add_textbox` is inherited from `_BaseGroupShapes`; the runtime
    # -- collection is always a `_BaseGroupShapes` subclass, but the
    # -- accessor is typed as the weaker `_BaseShapes` to keep callers
    # -- from depending on internal class nesting.
    assert isinstance(shapes, _BaseGroupShapes)
    shape = shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    assert isinstance(shape, Shape)
    assert len(shapes) == original_count + 1
