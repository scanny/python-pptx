"""Test suite for pptx.shapes.group module."""

from __future__ import annotations

import pytest

from pptx import Presentation
from pptx.dml.effect import ShadowFormat
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.shapes.group import GroupShape
from pptx.shapes.shapetree import GroupShapes
from pptx.util import Emu, Inches

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, initializer_mock, instance_mock


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
