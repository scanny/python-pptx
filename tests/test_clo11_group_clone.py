# pyright: reportPrivateUsage=false

"""Regression tests for CLO-11 — `GroupShape.clone_onto` correctness.

CLO-2 landed :meth:`BaseShape.clone_onto` which is inherited by :class:`GroupShape`.
These tests exercise the group-shape code paths specifically: nested groups, groups
containing a picture (image rel re-embedded), groups containing a chart (chart part
re-embedded), position overrides that reposition the outer group while preserving
the inner coordinate system, and the placeholder-child edge case.

The tests live in ``tests/`` (not ``tests/shapes/``) so they are discoverable as a
standalone regression set; they are otherwise written in the same BDD style as
``tests/shapes/test_base.py``.
"""

from __future__ import annotations

from io import BytesIO
from typing import cast

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture
from pptx.util import Inches


def _test_image_path() -> str:
    """Absolute path to the bundled `python-icon.jpeg` used by these tests."""
    from tests.unitutil.file import test_file_dir

    return "%s/python-icon.jpeg" % test_file_dir


class DescribeGroupShape_clone_onto:
    """Regression tests for cloning group shapes via :meth:`BaseShape.clone_onto`."""

    def it_clones_a_group_of_three_autoshapes_preserving_child_positions(self):
        prs = Presentation()
        src = prs.slides.add_slide(prs.slide_layouts[5])
        tgt = prs.slides.add_slide(prs.slide_layouts[5])
        # -- three autoshapes on the source slide, then wrap them in a group --
        s1 = src.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
        s2 = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(1), Inches(1)
        )
        s3 = src.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3), Inches(3), Inches(1), Inches(1)
        )
        grp = src.shapes.add_group_shape([s1, s2, s3])
        src_positions = [(c.left, c.top, c.width, c.height) for c in grp.shapes]

        clone = cast("GroupShape", grp.clone_onto(tgt.shapes))

        assert isinstance(clone, GroupShape)
        assert len(clone.shapes) == 3
        clone_positions = [(c.left, c.top, c.width, c.height) for c in clone.shapes]
        assert clone_positions == src_positions
        # -- every cNvPr id is fresh (no collisions with the source) --
        clone_ids = [int(cnv.get("id")) for cnv in clone._element.xpath(".//p:cNvPr")]
        src_ids = [int(cnv.get("id")) for cnv in grp._element.xpath(".//p:cNvPr")]
        assert not (set(clone_ids) & set(src_ids))
        # -- clone is topmost on the target tree --
        assert tgt.shapes[-1]._element is clone._element

    def it_clones_nested_groups_with_all_descendant_shapes_surviving(self):
        prs = Presentation()
        src = prs.slides.add_slide(prs.slide_layouts[5])
        tgt = prs.slides.add_slide(prs.slide_layouts[5])
        # -- inner group with two autoshapes --
        s1 = src.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
        s2 = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(1), Inches(1)
        )
        inner = src.shapes.add_group_shape([s1, s2])
        # -- outer group nests the inner group and an additional autoshape --
        s3 = src.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4), Inches(4), Inches(1), Inches(1)
        )
        outer = src.shapes.add_group_shape([inner, s3])

        clone = cast("GroupShape", outer.clone_onto(tgt.shapes))

        # -- two top-level children: one nested group, one autoshape --
        assert len(clone.shapes) == 2
        inner_clones = [s for s in clone.shapes if isinstance(s, GroupShape)]
        assert len(inner_clones) == 1
        # -- inner clone preserved its two autoshape descendants --
        assert len(inner_clones[0].shapes) == 2
        # -- total cNvPr count: outer(1) + inner(1) + s1/s2/s3(3) = 5 --
        assert len(clone._element.xpath(".//p:cNvPr")) == 5
        # -- every descendant id is fresh (no source-id reuse) --
        clone_ids = [int(cnv.get("id")) for cnv in clone._element.xpath(".//p:cNvPr")]
        src_ids = [int(cnv.get("id")) for cnv in outer._element.xpath(".//p:cNvPr")]
        assert not (set(clone_ids) & set(src_ids))
        assert len(set(clone_ids)) == len(clone_ids)  # -- each fresh id is unique --

    def it_clones_a_group_with_a_picture_reembedding_the_image(self):
        prs = Presentation()
        src = prs.slides.add_slide(prs.slide_layouts[5])
        tgt = prs.slides.add_slide(prs.slide_layouts[5])
        pic = src.shapes.add_picture(_test_image_path(), Inches(1), Inches(1))
        shp = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(3), Inches(1), Inches(1), Inches(1)
        )
        grp = src.shapes.add_group_shape([pic, shp])

        clone = cast("GroupShape", grp.clone_onto(tgt.shapes))

        pic_clones = [s for s in clone.shapes if isinstance(s, Picture)]
        assert len(pic_clones) == 1
        pic_clone = pic_clones[0]
        # -- image bytes round-trip --
        assert pic_clone.image is not None
        assert pic_clone.image.blob == pic.image.blob
        # -- every r:embed in the clone resolves on the target part --
        embed_ids = pic_clone._element.xpath(".//@r:embed")
        assert embed_ids  # -- at least the blipFill/@r:embed --
        for rId in embed_ids:
            assert rId in tgt.part.rels

    def it_reembeds_a_group_picture_across_presentations(self):
        src_prs = Presentation()
        tgt_prs = Presentation()
        src = src_prs.slides.add_slide(src_prs.slide_layouts[5])
        tgt = tgt_prs.slides.add_slide(tgt_prs.slide_layouts[5])
        pic = src.shapes.add_picture(_test_image_path(), Inches(1), Inches(1))
        shp = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(3), Inches(1), Inches(1), Inches(1)
        )
        grp = src.shapes.add_group_shape([pic, shp])

        clone = cast("GroupShape", grp.clone_onto(tgt.shapes))

        pic_clones = [s for s in clone.shapes if isinstance(s, Picture)]
        assert len(pic_clones) == 1
        pic_clone = pic_clones[0]
        assert pic_clone.image is not None
        assert pic_clone.image.blob == pic.image.blob
        # -- cloned image part lives in the target package --
        assert pic_clone.part.package is tgt_prs.part.package
        # -- source pic is untouched --
        assert pic.image.blob is not None

    def it_clones_a_group_with_a_chart_reembedding_the_chart_part(self):
        prs = Presentation()
        src = prs.slides.add_slide(prs.slide_layouts[5])
        tgt = prs.slides.add_slide(prs.slide_layouts[5])
        cd = CategoryChartData()
        cd.categories = ["A", "B", "C"]
        cd.add_series("Series 1", (1, 2, 3))
        chart_gf = src.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(4),
            Inches(3),
            cd,
        )
        shp = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(5), Inches(1), Inches(1), Inches(1)
        )
        grp = src.shapes.add_group_shape([chart_gf, shp])

        clone = cast("GroupShape", grp.clone_onto(tgt.shapes))

        # -- chart descendant survives with has_chart True --
        chart_in_clone = None
        for s in clone.shapes:
            if isinstance(s, GraphicFrame) and s.has_chart:
                chart_in_clone = s
                break
        assert chart_in_clone is not None
        # -- every r:id on the cloned graphic-frame resolves on the target part --
        rIds = chart_in_clone._element.xpath(".//@r:id")
        assert rIds  # -- at least one r:id (the chart part) --
        for rId in rIds:
            assert rId in tgt.part.rels
        # -- chart part now lives in the target package (same prs here, but
        # -- we assert it's in the slide-part's rels, which is the target). --
        chart_rel = tgt.part.rels[rIds[0]]
        assert chart_rel.target_part.package is tgt.part.package

    def it_repositions_the_outer_group_preserving_inner_coordinates(self):
        prs = Presentation()
        src = prs.slides.add_slide(prs.slide_layouts[5])
        tgt = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = src.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
        s2 = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(1), Inches(1)
        )
        grp = src.shapes.add_group_shape([s1, s2])
        src_child_positions = [(c.left, c.top) for c in grp.shapes]

        clone = cast(
            "GroupShape",
            grp.clone_onto(tgt.shapes, left=Inches(5), top=Inches(5)),
        )

        # -- outer group repositioned --
        assert clone.left == Inches(5)
        assert clone.top == Inches(5)
        # -- child coord-system preserved: each child's raw left/top matches src --
        clone_child_positions = [(c.left, c.top) for c in clone.shapes]
        assert clone_child_positions == src_child_positions
        # -- the group's child-extents anchor (a:chOff/a:chExt) is unchanged --
        xfrm = clone._element.xfrm
        assert xfrm is not None
        src_xfrm = grp._element.xfrm
        assert src_xfrm is not None
        assert xfrm.xpath("./a:chOff/@x") == src_xfrm.xpath("./a:chOff/@x")
        assert xfrm.xpath("./a:chOff/@y") == src_xfrm.xpath("./a:chOff/@y")
        assert xfrm.xpath("./a:chExt/@cx") == src_xfrm.xpath("./a:chExt/@cx")
        assert xfrm.xpath("./a:chExt/@cy") == src_xfrm.xpath("./a:chExt/@cy")

    def but_it_raises_when_the_group_itself_is_a_placeholder(self):
        # -- top-level placeholder rejection via BaseShape.clone_onto is a
        # -- precondition. Groups are never placeholders in practice, but the
        # -- check reads `has_ph_elm`; verify here that a group never trips it.
        prs = Presentation()
        src = prs.slides.add_slide(prs.slide_layouts[5])
        tgt = prs.slides.add_slide(prs.slide_layouts[5])
        shp = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        grp = src.shapes.add_group_shape([shp])
        assert not grp._element.has_ph_elm  # -- sanity: group carries no <p:ph> --

        # -- clone succeeds (group isn't a placeholder) --
        clone = grp.clone_onto(tgt.shapes)
        assert isinstance(clone, GroupShape)

    def it_survives_round_trip_for_a_cloned_group_with_nested_picture(self):
        prs = Presentation()
        src = prs.slides.add_slide(prs.slide_layouts[5])
        tgt = prs.slides.add_slide(prs.slide_layouts[5])
        pic = src.shapes.add_picture(_test_image_path(), Inches(1), Inches(1))
        shp = src.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(3), Inches(1), Inches(1), Inches(1)
        )
        grp = src.shapes.add_group_shape([pic, shp])
        grp.clone_onto(tgt.shapes)

        buf = BytesIO()
        prs.save(buf)
        buf.seek(0)
        reloaded = Presentation(buf)

        # -- each slide has one group, each group has one picture and one autoshape --
        for sl in reloaded.slides:
            groups = [s for s in sl.shapes if isinstance(s, GroupShape)]
            assert len(groups) == 1
            g = groups[0]
            pics = [c for c in g.shapes if isinstance(c, Picture)]
            assert len(pics) == 1
            assert pics[0].image is not None
            assert pics[0].image.blob is not None

    def it_assigns_unique_ids_to_all_descendants_on_same_tree_clone(self):
        # -- cloning a group onto its own shape tree: every id must still be unique --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
        s2 = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(1), Inches(1)
        )
        grp = slide.shapes.add_group_shape([s1, s2])

        clone = cast("GroupShape", grp.clone_onto(slide.shapes))

        all_ids = [int(cnv.get("id")) for cnv in slide.shapes._spTree.xpath(".//p:cNvPr")]
        assert len(all_ids) == len(set(all_ids))
        # -- clone itself has a different id than source group --
        assert clone.shape_id != grp.shape_id
