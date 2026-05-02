"""Test suite for pptx.shapes.group module."""

from __future__ import annotations

import pytest

from pptx import Presentation
from pptx.dml.effect import ShadowFormat
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
from pptx.shapes.group import GroupShape
from pptx.shapes.shapetree import GroupShapes
from pptx.util import Emu, Inches

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, initializer_mock, instance_mock


def BaseShape_for_elm(shape_elm, parent):
    """Return a bare :class:`BaseShape` proxy over `shape_elm`.

    Used in tests to read composited ``effective_*`` values on a raw shape
    element without needing the full concrete subclass hierarchy.
    """
    return BaseShape(shape_elm, parent)


class DescribeGroupShape(object):
    def it_raises_on_access_click_action(self, click_action_fixture):
        group = click_action_fixture
        with pytest.raises(TypeError):
            group.click_action

    def it_provides_access_to_its_shadow(self, ShadowFormat_, shadow_):
        grpSp = element("p:grpSp/p:grpSpPr")
        grpSpPr = grpSp.xpath("//p:grpSpPr")[0]
        ShadowFormat_.return_value = shadow_
        group_shape = GroupShape(grpSp, None)

        shadow = group_shape.shadow

        ShadowFormat_.assert_called_once_with(grpSpPr)
        assert shadow is shadow_

    def it_knows_its_shape_type(self, shape_type_fixture):
        group = shape_type_fixture
        assert group.shape_type == MSO_SHAPE_TYPE.GROUP

    def it_provides_access_to_its_sub_shapes(self, shapes_fixture):
        group, GroupShapes_init_, grpSp = shapes_fixture

        shapes = group.shapes

        GroupShapes_init_.assert_called_once_with(shapes, grpSp, group)
        assert isinstance(shapes, GroupShapes)

    # -- #1085: GroupShape.duplicate() -------------------------------------

    def it_can_duplicate_a_top_level_group(self):
        # -- top-level group with two simple children. Duplicate should land --
        # -- at the same slide-relative rectangle as the original.           --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        s2 = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(3), Inches(2), Inches(1), Inches(1)
        )
        grp = slide.shapes.add_group_shape([s1, s2])
        orig_count = len(slide.shapes)
        eff_left = grp.effective_left
        eff_top = grp.effective_top
        eff_width = grp.effective_width
        eff_height = grp.effective_height

        dup = grp.duplicate()

        assert isinstance(dup, GroupShape)
        assert len(slide.shapes) == orig_count + 1
        assert dup.shape_id != grp.shape_id
        assert dup.name != grp.name
        # -- cloned group occupies the same slide-rectangle as the source --
        assert dup.effective_left == eff_left
        assert dup.effective_top == eff_top
        assert dup.effective_width == eff_width
        assert dup.effective_height == eff_height
        # -- duplicate contains the same number of child shapes --
        assert len(list(dup.shapes)) == len(list(grp.shapes))
        # -- every shape id on the slide is unique --
        ids = slide.shapes._spTree.xpath(".//p:cNvPr/@id")
        assert len(ids) == len(set(ids))

    def it_places_a_clone_of_a_nested_group_at_effective_slide_coords(self):
        # -- #925 2:1 group fixture: outer group maps a 10000-wide child coord  --
        # -- space to 5000 EMU on the slide (sx=0.5), sy=1.0. An inner p:grpSp  --
        # -- nested inside has raw off/ext in the outer's child coord space; a  --
        # -- naive clone that just copied the raw off/ext would place the       --
        # -- duplicate at the wrong slide-relative rectangle. Using             --
        # -- effective_* on the source yields the correct slide coords.         --
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls, qn
        from pptx.oxml.shapes.groupshape import CT_GroupShape

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        spTree = slide.shapes._spTree

        # -- hand-craft the nested 2:1 group: outer wraps an inner group that --
        # -- itself contains a single rectangle.                               --
        outer_xml = (
            f"<p:grpSp {nsdecls('p', 'a', 'r')}>"
            "  <p:nvGrpSpPr>"
            '    <p:cNvPr id="100" name="Outer"/>'
            "    <p:cNvGrpSpPr/>"
            "    <p:nvPr/>"
            "  </p:nvGrpSpPr>"
            "  <p:grpSpPr>"
            "    <a:xfrm>"
            '      <a:off x="1000" y="2000"/>'
            '      <a:ext cx="5000" cy="8000"/>'
            '      <a:chOff x="0" y="0"/>'
            '      <a:chExt cx="10000" cy="8000"/>'
            "    </a:xfrm>"
            "  </p:grpSpPr>"
            "  <p:grpSp>"
            "    <p:nvGrpSpPr>"
            '      <p:cNvPr id="101" name="Inner"/>'
            "      <p:cNvGrpSpPr/>"
            "      <p:nvPr/>"
            "    </p:nvGrpSpPr>"
            "    <p:grpSpPr>"
            "      <a:xfrm>"
            '        <a:off x="2000" y="4000"/>'
            '        <a:ext cx="4000" cy="2000"/>'
            '        <a:chOff x="0" y="0"/>'
            '        <a:chExt cx="4000" cy="2000"/>'
            "      </a:xfrm>"
            "    </p:grpSpPr>"
            "    <p:sp>"
            "      <p:nvSpPr>"
            '        <p:cNvPr id="102" name="Rect"/>'
            "        <p:cNvSpPr/>"
            "        <p:nvPr/>"
            "      </p:nvSpPr>"
            "      <p:spPr>"
            "        <a:xfrm>"
            '          <a:off x="0" y="0"/>'
            '          <a:ext cx="4000" cy="2000"/>'
            "        </a:xfrm>"
            "      </p:spPr>"
            "    </p:sp>"
            "  </p:grpSp>"
            "</p:grpSp>"
        )
        outer_grpSp = parse_xml(outer_xml)
        spTree.insert_element_before(outer_grpSp, "p:extLst")

        # -- the *inner* group is what we duplicate; it's nested inside outer  --
        inner_grpSp = outer_grpSp.xpath("./p:grpSp")[0]
        assert isinstance(inner_grpSp, CT_GroupShape)
        inner_group = GroupShape(inner_grpSp, slide.shapes)

        # -- sanity-check the #925 composited geometry:                        --
        # --   outer sx = 5000/10000 = 0.5, sy = 8000/8000 = 1.0               --
        # --   inner raw off/ext (2000, 4000, 4000, 2000) in outer child space --
        # --   -> effective (1000 + 2000*0.5, 2000 + 4000*1.0, 4000*0.5, 2000) --
        # --                = (2000, 6000, 2000, 2000) in slide coords         --
        assert inner_group.effective_left == 2000
        assert inner_group.effective_top == 6000
        assert inner_group.effective_width == 2000
        assert inner_group.effective_height == 2000

        dup = inner_group.duplicate()

        # -- clone is placed at the effective slide rectangle, NOT at the raw --
        # -- (2000, 4000, 4000, 2000) taken from the outer-child coord space  --
        assert isinstance(dup, GroupShape)
        assert dup.effective_left == Emu(2000)
        assert dup.effective_top == Emu(6000)
        assert dup.effective_width == Emu(2000)
        assert dup.effective_height == Emu(2000)
        # -- the clone is now a top-level sibling (parented by spTree), not   --
        # -- a descendant of the outer group.                                 --
        assert dup._element.getparent().tag == qn("p:spTree")
        # -- every shape id in the tree is still unique after duplication    --
        ids = spTree.xpath(".//p:cNvPr/@id")
        assert len(ids) == len(set(ids))

    def it_clones_part_relationships_when_duplicating_a_group_with_a_picture(self):
        # -- group contains a picture; rel-cloner must establish a new rel on --
        # -- the slide part and rewrite the blip's r:embed on the clone.     --
        from tests.unitutil.file import test_file_dir

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        pic = slide.shapes.add_picture(
            "%s/python-icon.jpeg" % test_file_dir, Inches(1), Inches(1)
        )
        grp = slide.shapes.add_group_shape([pic])
        orig_rel_count = len(slide.part.rels)

        dup = grp.duplicate()

        # -- every rId in the clone's subtree is a valid rel on the slide    --
        # -- part (rel-cloner should reuse the existing image rel when the   --
        # -- source and target are the same slide part).                     --
        rIds = dup._element.xpath(".//@r:embed | .//@r:id | .//@r:link")
        for rId in rIds:
            # -- will raise KeyError if the rId is not a valid relationship --
            assert slide.part.rels[rId] is not None
        # -- same-slide clone: no new image parts materialised (rel-cloner   --
        # -- reuses the existing target part when src and tgt are in the     --
        # -- same package).                                                  --
        assert len(slide.part.rels) == orig_rel_count

    # -- #730: GroupShape.ungroup() ---------------------------------------

    def it_can_ungroup_a_top_level_group(self):
        # -- A group with two simple children placed at known slide          --
        # -- coordinates. After ungroup the freed shapes should live at the  --
        # -- same slide-rectangle they rendered at before.                   --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2))
        s2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4), Inches(3), Inches(1), Inches(1))
        grp = slide.shapes.add_group_shape([s1, s2])
        grp_cNvPr_id = grp.shape_id
        before_effective = [
            (c.effective_left, c.effective_top, c.effective_width, c.effective_height)
            for c in grp.shapes
        ]
        before_top_count = len(slide.shapes)

        freed = grp.ungroup()

        # -- two shapes were freed --
        assert len(freed) == 2
        # -- grpSp element was removed (no longer referenced from spTree) --
        spTree = slide.shapes._spTree
        assert grp._element.getparent() is None
        assert grp_cNvPr_id not in [int(i) for i in spTree.xpath(".//p:cNvPr/@id")]
        # -- top-level shape count is unchanged: the grpSp was replaced by   --
        # -- its two children as top-level siblings.                         --
        assert len(slide.shapes) == before_top_count + 1  # two children + group removed = +1 net
        # -- each freed shape now lives directly under the spTree --
        for shape in freed:
            assert shape._element.getparent() is spTree
        # -- effective slide-rectangle for each freed shape matches what it  --
        # -- rendered at before the ungroup.                                 --
        after_effective = [
            (s.effective_left, s.effective_top, s.effective_width, s.effective_height)
            for s in freed
        ]
        assert after_effective == before_effective

    def it_returns_freed_shapes_in_original_zorder(self):
        # -- preserving children's relative z-order is a correctness invariant --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))
        s2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), Inches(2), Inches(1), Inches(1))
        s3 = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(3), Inches(3), Inches(1), Inches(1))
        grp = slide.shapes.add_group_shape([s1, s2, s3])

        freed = grp.ungroup()

        # -- returned list matches child document order inside the group; --
        # -- in a newly-created group the insertion order was s1, s2, s3. --
        assert [s._element for s in freed] == [s1._element, s2._element, s3._element]

    def it_ungroups_a_nested_group_and_lands_children_at_slide_coords(self):
        # -- The #925 2:1 nested-group fixture: the inner group's child lives    --
        # -- in a coord space the outer maps 2:1 to the slide. After ungrouping  --
        # -- the inner group its child must render at the same slide rectangle.  --
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls
        from pptx.oxml.shapes.groupshape import CT_GroupShape

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        spTree = slide.shapes._spTree
        outer_xml = (
            f"<p:grpSp {nsdecls('p', 'a', 'r')}>"
            "  <p:nvGrpSpPr>"
            '    <p:cNvPr id="200" name="Outer"/>'
            "    <p:cNvGrpSpPr/>"
            "    <p:nvPr/>"
            "  </p:nvGrpSpPr>"
            "  <p:grpSpPr>"
            "    <a:xfrm>"
            '      <a:off x="1000" y="2000"/>'
            '      <a:ext cx="5000" cy="8000"/>'
            '      <a:chOff x="0" y="0"/>'
            '      <a:chExt cx="10000" cy="8000"/>'
            "    </a:xfrm>"
            "  </p:grpSpPr>"
            "  <p:grpSp>"
            "    <p:nvGrpSpPr>"
            '      <p:cNvPr id="201" name="Inner"/>'
            "      <p:cNvGrpSpPr/>"
            "      <p:nvPr/>"
            "    </p:nvGrpSpPr>"
            "    <p:grpSpPr>"
            "      <a:xfrm>"
            '        <a:off x="2000" y="4000"/>'
            '        <a:ext cx="4000" cy="2000"/>'
            '        <a:chOff x="0" y="0"/>'
            '        <a:chExt cx="4000" cy="2000"/>'
            "      </a:xfrm>"
            "    </p:grpSpPr>"
            "    <p:sp>"
            "      <p:nvSpPr>"
            '        <p:cNvPr id="202" name="Rect"/>'
            "        <p:cNvSpPr/>"
            "        <p:nvPr/>"
            "      </p:nvSpPr>"
            "      <p:spPr>"
            "        <a:xfrm>"
            '          <a:off x="0" y="0"/>'
            '          <a:ext cx="4000" cy="2000"/>'
            "        </a:xfrm>"
            '        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            "      </p:spPr>"
            "      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>"
            "    </p:sp>"
            "  </p:grpSp>"
            "</p:grpSp>"
        )
        outer_grpSp = parse_xml(outer_xml)
        spTree.insert_element_before(outer_grpSp, "p:extLst")

        inner_grpSp = outer_grpSp.xpath("./p:grpSp")[0]
        assert isinstance(inner_grpSp, CT_GroupShape)
        inner_group = GroupShape(inner_grpSp, slide.shapes)

        # -- inner group child's effective rectangle: raw (0, 0, 4000, 2000)  --
        # -- inside inner (i_off 2000/4000, i_sx 1.0, i_sy 1.0) -> inside     --
        # -- outer (outer.off 1000/2000, outer.sx 0.5, outer.sy 1.0) yields   --
        # -- slide (1000 + 2000*0.5, 2000 + 4000, 4000*0.5, 2000) =           --
        # -- (2000, 6000, 2000, 2000).                                        --
        child_elm = inner_grpSp.xpath("./p:sp")[0]
        child_proxy = BaseShape_for_elm(child_elm, inner_group)
        assert child_proxy.effective_left == 2000
        assert child_proxy.effective_top == 6000
        assert child_proxy.effective_width == 2000
        assert child_proxy.effective_height == 2000

        freed = inner_group.ungroup()

        assert len(freed) == 1
        shape = freed[0]
        # -- freed shape is now a top-level sibling of spTree                 --
        assert shape._element.getparent() is spTree
        # -- and its raw xfrm has been rewritten to its former slide-rect    --
        # -- because now nothing above it composites a transform.            --
        assert shape.effective_left == 2000
        assert shape.effective_top == 6000
        assert shape.effective_width == 2000
        assert shape.effective_height == 2000
        # -- inner p:grpSp has been removed from the tree                    --
        assert inner_grpSp.getparent() is None
        # -- outer p:grpSp still exists but is now empty of shape children   --
        assert outer_grpSp.getparent() is spTree
        assert list(outer_grpSp.iter_shape_elms()) == []

    def it_preserves_child_groups_internal_layout(self):
        # -- Ungrouping an outer group with a nested child-group should hoist --
        # -- the child group to the slide with its internal chOff/chExt       --
        # -- unchanged and its off/ext rewritten to its effective slide rect, --
        # -- so its leaf descendants still render at their original slide     --
        # -- positions (no cumulative distortion).                            --
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        spTree = slide.shapes._spTree
        outer_xml = (
            f"<p:grpSp {nsdecls('p', 'a', 'r')}>"
            "  <p:nvGrpSpPr>"
            '    <p:cNvPr id="300" name="Outer"/>'
            "    <p:cNvGrpSpPr/>"
            "    <p:nvPr/>"
            "  </p:nvGrpSpPr>"
            "  <p:grpSpPr>"
            "    <a:xfrm>"
            '      <a:off x="1000" y="2000"/>'
            '      <a:ext cx="5000" cy="8000"/>'
            '      <a:chOff x="0" y="0"/>'
            '      <a:chExt cx="10000" cy="8000"/>'
            "    </a:xfrm>"
            "  </p:grpSpPr>"
            "  <p:grpSp>"
            "    <p:nvGrpSpPr>"
            '      <p:cNvPr id="301" name="Inner"/>'
            "      <p:cNvGrpSpPr/>"
            "      <p:nvPr/>"
            "    </p:nvGrpSpPr>"
            "    <p:grpSpPr>"
            "      <a:xfrm>"
            '        <a:off x="2000" y="4000"/>'
            '        <a:ext cx="4000" cy="2000"/>'
            '        <a:chOff x="0" y="0"/>'
            '        <a:chExt cx="4000" cy="2000"/>'
            "      </a:xfrm>"
            "    </p:grpSpPr>"
            "    <p:sp>"
            "      <p:nvSpPr>"
            '        <p:cNvPr id="302" name="Leaf"/>'
            "        <p:cNvSpPr/>"
            "        <p:nvPr/>"
            "      </p:nvSpPr>"
            "      <p:spPr>"
            "        <a:xfrm>"
            '          <a:off x="0" y="0"/>'
            '          <a:ext cx="4000" cy="2000"/>'
            "        </a:xfrm>"
            '        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            "      </p:spPr>"
            "      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>"
            "    </p:sp>"
            "  </p:grpSp>"
            "</p:grpSp>"
        )
        outer_grpSp = parse_xml(outer_xml)
        spTree.insert_element_before(outer_grpSp, "p:extLst")
        outer_group = GroupShape(outer_grpSp, slide.shapes)
        leaf_elm = outer_grpSp.xpath(".//p:sp")[0]
        leaf_before = BaseShape_for_elm(leaf_elm, outer_group)
        leaf_eff_before = (
            leaf_before.effective_left,
            leaf_before.effective_top,
            leaf_before.effective_width,
            leaf_before.effective_height,
        )

        freed = outer_group.ungroup()

        # -- one child was freed: the inner group (with the leaf inside)    --
        assert len(freed) == 1
        inner_group_proxy = freed[0]
        assert isinstance(inner_group_proxy, GroupShape)
        inner_grpSp = inner_group_proxy._element
        # -- inner group's chOff/chExt unchanged; off/ext rewritten to its  --
        # -- effective slide rectangle (2000, 6000, 2000, 2000).            --
        assert inner_grpSp.xpath("./p:grpSpPr/a:xfrm/a:chOff/@x") == ["0"]
        assert inner_grpSp.xpath("./p:grpSpPr/a:xfrm/a:chExt/@cx") == ["4000"]
        assert int(inner_grpSp.xpath("./p:grpSpPr/a:xfrm/a:off/@x")[0]) == 2000
        assert int(inner_grpSp.xpath("./p:grpSpPr/a:xfrm/a:off/@y")[0]) == 6000
        assert int(inner_grpSp.xpath("./p:grpSpPr/a:xfrm/a:ext/@cx")[0]) == 2000
        assert int(inner_grpSp.xpath("./p:grpSpPr/a:xfrm/a:ext/@cy")[0]) == 2000
        # -- leaf's effective_* unchanged: the inner group now carries the  --
        # -- scale that the outer used to contribute.                       --
        leaf_after = BaseShape_for_elm(leaf_elm, inner_group_proxy)
        leaf_eff_after = (
            leaf_after.effective_left,
            leaf_after.effective_top,
            leaf_after.effective_width,
            leaf_after.effective_height,
        )
        assert leaf_eff_after == leaf_eff_before

    def it_can_ungroup_recursively_to_fully_flatten_nested_groups(self):
        # -- Callers can flatten fully by recursively ungrouping every freed  --
        # -- GroupShape. The returned list makes this straightforward.       --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))
        s2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3), Inches(3), Inches(1), Inches(1))
        inner = slide.shapes.add_group_shape([s1, s2])
        s3 = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(5), Inches(5), Inches(1), Inches(1))
        outer = slide.shapes.add_group_shape([inner, s3])

        # -- recursive ungroup helper --
        def _flatten(grp: GroupShape) -> list:
            out: list = []
            for f in grp.ungroup():
                if isinstance(f, GroupShape):
                    out.extend(_flatten(f))
                else:
                    out.append(f)
            return out

        flat = _flatten(outer)
        assert len(flat) == 3
        # -- no grpSp elements remain anywhere under the spTree --
        assert slide.shapes._spTree.xpath(".//p:grpSp") == []
        # -- every remaining shape sits directly under the spTree --
        for f in flat:
            assert f._element.getparent() is slide.shapes._spTree

    def it_raises_when_ungrouping_a_detached_group(self):
        # -- ungroup() requires the group to be attached to a shape tree --
        from pptx.oxml.shapes.groupshape import CT_GroupShape

        loose_grpSp = CT_GroupShape.new_grpSp(1, "Loose")
        loose = GroupShape(loose_grpSp, None)

        with pytest.raises(ValueError):
            loose.ungroup()

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def click_action_fixture(self):
        return GroupShape(None, None)

    @pytest.fixture
    def shape_type_fixture(self):
        return GroupShape(None, None)

    @pytest.fixture
    def shapes_fixture(self, GroupShapes_init_):
        grpSp = element("p:grpSp")
        group = GroupShape(grpSp, None)
        return group, GroupShapes_init_, grpSp

    # fixture components ---------------------------------------------

    @pytest.fixture
    def GroupShapes_init_(self, request):
        return initializer_mock(request, GroupShapes, autospec=True)

    @pytest.fixture
    def shadow_(self, request):
        return instance_mock(request, ShadowFormat)

    @pytest.fixture
    def ShadowFormat_(self, request):
        return class_mock(request, "pptx.shapes.group.ShadowFormat")
