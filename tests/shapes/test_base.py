# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.shapes.base` module."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

import pytest

from pptx import Presentation
from pptx.action import ActionSetting
from pptx.dml.effect import ShadowFormat
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE, PP_PLACEHOLDER
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.text import CT_TextBody
from pptx.shapes import Subshape
from pptx.shapes.autoshape import Shape
from pptx.shapes.base import BaseShape, _PlaceholderFormat
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.shapetree import BaseShapeFactory, SlideShapes
from pptx.util import Inches

from ..oxml.unitdata.shape import (
    a_cNvPr,
    a_cxnSp,
    a_graphicFrame,
    a_grpSp,
    a_grpSpPr,
    a_p_xfrm,
    a_pic,
    an_ext,
    an_nvSpPr,
    an_off,
    an_sp,
    an_spPr,
    an_xfrm,
)
from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock, loose_mock

if TYPE_CHECKING:
    from pptx.opc.package import XmlPart
    from pptx.oxml.shapes import ShapeElement
    from pptx.shapes.connector import Connector
    from pptx.types import ProvidesPart


class DescribeBaseShape(object):
    """Unit-test suite for `pptx.shapes.base.BaseShape` objects."""

    def it_provides_access_to_its_click_action(self, click_action_fixture):
        shape, ActionSetting_, cNvPr, click_action_ = click_action_fixture
        click_action = shape.click_action
        ActionSetting_.assert_called_once_with(cNvPr, shape)
        assert click_action is click_action_

    def it_provides_access_to_its_hover_action(self, hover_action_fixture):
        shape, ActionSetting_, cNvPr, hover_action_ = hover_action_fixture
        hover_action = shape.hover_action
        ActionSetting_.assert_called_once_with(cNvPr, shape, hover=True)
        assert hover_action is hover_action_

    def it_returns_the_same_hover_action_on_repeated_access(
        self, ActionSetting_, action_setting_
    ):
        sp = element("p:sp/p:nvSpPr/p:cNvPr")
        shape = BaseShape(sp, None)

        first = shape.hover_action
        second = shape.hover_action

        # lazyproperty: one construction, two reads
        assert ActionSetting_.call_count == 1
        assert first is second is action_setting_

    def it_knows_its_shape_id(self, id_fixture):
        shape, expected_value = id_fixture
        assert shape.shape_id == expected_value

    def it_knows_its_name(self, name_get_fixture):
        shape, name = name_get_fixture
        assert shape.name == name

    def it_can_change_its_name(self, name_set_fixture):
        shape, new_value, expected_xml = name_set_fixture
        shape.name = new_value
        assert shape._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("shape_cxml", "expected_x", "expected_y"),
        [
            ("p:cxnSp/p:spPr", None, None),
            ("p:cxnSp/p:spPr/a:xfrm", None, None),
            ("p:cxnSp/p:spPr/a:xfrm/a:off{x=123,y=456}", 123, 456),
            ("p:graphicFrame/p:xfrm", None, None),
            ("p:graphicFrame/p:xfrm/a:off{x=123,y=456}", 123, 456),
            ("p:grpSp/p:grpSpPr", None, None),
            ("p:grpSp/p:grpSpPr/a:xfrm/a:off{x=123,y=456}", 123, 456),
            ("p:pic/p:spPr", None, None),
            ("p:pic/p:spPr/a:xfrm", None, None),
            ("p:pic/p:spPr/a:xfrm/a:off{x=123,y=456}", 123, 456),
            ("p:sp/p:spPr", None, None),
            ("p:sp/p:spPr/a:xfrm", None, None),
            ("p:sp/p:spPr/a:xfrm/a:off{x=123,y=456}", 123, 456),
        ],
    )
    def it_has_a_position(
        self,
        shape_cxml: str,
        expected_x: int | None,
        expected_y: int | None,
        provides_part: ProvidesPart,
    ):
        shape_elm = cast("ShapeElement", element(shape_cxml))

        shape = BaseShape(shape_elm, provides_part)

        assert shape.left == expected_x
        assert shape.top == expected_y

    @pytest.fixture
    def provides_part(self) -> ProvidesPart:

        class FakeProvidesPart:
            @property
            def part(self) -> XmlPart:
                raise NotImplementedError

        return FakeProvidesPart()

    def it_can_change_its_position(self, position_set_fixture):
        shape, left, top, expected_xml = position_set_fixture
        shape.left = left
        shape.top = top
        assert shape._element.xml == expected_xml

    def it_has_dimensions(self, dimensions_get_fixture):
        shape, expected_width, expected_height = dimensions_get_fixture
        assert shape.width == expected_width
        assert shape.height == expected_height

    def it_can_change_its_dimensions(self, dimensions_set_fixture):
        shape, width, height, expected_xml = dimensions_set_fixture
        shape.width = width
        shape.height = height
        assert shape._element.xml == expected_xml

    # -- #925: effective_* composited-transform geometry --------------------

    def it_reports_effective_geometry_equal_to_raw_for_top_level_shape(self):
        # -- shape directly under p:spTree: no enclosing group, so effective --
        # -- values match the raw xfrm values exactly.                       --
        spTree = cast(
            "ShapeElement",
            element(
                "p:spTree/(p:nvGrpSpPr,p:grpSpPr,"
                "p:sp/p:spPr/a:xfrm/(a:off{x=100,y=200},a:ext{cx=300,cy=400}))"
            ),
        )
        sp = spTree.xpath("p:sp")[0]
        shape = BaseShape(cast("ShapeElement", sp), None)

        assert shape.effective_left == 100
        assert shape.effective_top == 200
        assert shape.effective_width == 300
        assert shape.effective_height == 400

    def it_applies_the_group_transform_to_an_enclosed_shape(self):
        # -- group's child coordinate system is twice the size of its slide  --
        # -- bounding box in x, so a child's position and width are halved.  --
        spTree = cast(
            "ShapeElement",
            element(
                "p:spTree/(p:nvGrpSpPr,p:grpSpPr,"
                "p:grpSp/(p:nvGrpSpPr,p:grpSpPr/a:xfrm/("
                "a:off{x=1000,y=2000},a:ext{cx=5000,cy=8000},"
                "a:chOff{x=0,y=0},a:chExt{cx=10000,cy=8000}),"
                "p:sp/p:spPr/a:xfrm/(a:off{x=2000,y=4000},"
                "a:ext{cx=4000,cy=2000})))"
            ),
        )
        sp = spTree.xpath(".//p:sp")[0]
        shape = BaseShape(cast("ShapeElement", sp), None)

        # -- raw values unchanged (backwards compatibility) --
        assert shape.left == 2000
        assert shape.top == 4000
        assert shape.width == 4000
        assert shape.height == 2000

        # -- effective values apply the group's (chOff/chExt -> off/ext) map:   --
        # --   sx = 5000/10000 = 0.5,  sy = 8000/8000 = 1.0                     --
        # --   ex_left = 1000 + (2000 - 0) * 0.5 = 2000                         --
        # --   ex_top  = 2000 + (4000 - 0) * 1.0 = 6000                         --
        # --   ex_w    = 4000 * 0.5 = 2000                                      --
        # --   ex_h    = 2000 * 1.0 = 2000                                      --
        assert shape.effective_left == 2000
        assert shape.effective_top == 6000
        assert shape.effective_width == 2000
        assert shape.effective_height == 2000

    def it_cascades_transforms_through_nested_groups(self):
        # -- outer group maps 10000 EMU child-space to 5000 EMU on slide (sx=0.5) --
        # -- inner group (itself inside outer) maps 1000 -> 500 in its own coords  --
        # --   => combined scale 0.25 for the innermost shape.                    --
        spTree = cast(
            "ShapeElement",
            element(
                "p:spTree/(p:nvGrpSpPr,p:grpSpPr,"
                "p:grpSp/(p:nvGrpSpPr,p:grpSpPr/a:xfrm/("
                "a:off{x=0,y=0},a:ext{cx=5000,cy=5000},"
                "a:chOff{x=0,y=0},a:chExt{cx=10000,cy=10000}),"
                "p:grpSp/(p:nvGrpSpPr,p:grpSpPr/a:xfrm/("
                "a:off{x=0,y=0},a:ext{cx=1000,cy=1000},"
                "a:chOff{x=0,y=0},a:chExt{cx=2000,cy=2000}),"
                "p:sp/p:spPr/a:xfrm/(a:off{x=400,y=400},"
                "a:ext{cx=800,cy=800}))))"
            ),
        )
        sp = spTree.xpath(".//p:sp")[0]
        shape = BaseShape(cast("ShapeElement", sp), None)

        # -- raw --
        assert shape.left == 400
        assert shape.width == 800

        # -- inner group: sx=sy=0.5 -> (200, 200) size 400x400 in outer space --
        # -- outer group: sx=sy=0.5 -> (100, 100) size 200x200 on slide --
        assert shape.effective_left == 100
        assert shape.effective_top == 100
        assert shape.effective_width == 200
        assert shape.effective_height == 200

    def it_translates_through_nonzero_group_offsets(self):
        # -- chOff is nonzero: child local-x=1000 is at the "origin" of group's  --
        # -- child space, so it should land at off.x on the slide.               --
        spTree = cast(
            "ShapeElement",
            element(
                "p:spTree/(p:nvGrpSpPr,p:grpSpPr,"
                "p:grpSp/(p:nvGrpSpPr,p:grpSpPr/a:xfrm/("
                "a:off{x=2000,y=3000},a:ext{cx=4000,cy=4000},"
                "a:chOff{x=1000,y=1000},a:chExt{cx=4000,cy=4000}),"
                "p:sp/p:spPr/a:xfrm/(a:off{x=1000,y=1000},"
                "a:ext{cx=2000,cy=2000})))"
            ),
        )
        sp = spTree.xpath(".//p:sp")[0]
        shape = BaseShape(cast("ShapeElement", sp), None)

        # -- scale 1.0 in both axes; translation is off - chOff = (1000, 2000) --
        assert shape.effective_left == 2000
        assert shape.effective_top == 3000
        assert shape.effective_width == 2000
        assert shape.effective_height == 2000

    def it_returns_raw_values_when_enclosing_group_has_no_xfrm(self):
        # -- missing a:xfrm on the group is unusual but possible; the walker    --
        # -- treats it as identity and moves on.                               --
        spTree = cast(
            "ShapeElement",
            element(
                "p:spTree/(p:nvGrpSpPr,p:grpSpPr,"
                "p:grpSp/(p:nvGrpSpPr,p:grpSpPr,"
                "p:sp/p:spPr/a:xfrm/(a:off{x=500,y=600},"
                "a:ext{cx=700,cy=800})))"
            ),
        )
        sp = spTree.xpath(".//p:sp")[0]
        shape = BaseShape(cast("ShapeElement", sp), None)

        assert shape.effective_left == 500
        assert shape.effective_top == 600
        assert shape.effective_width == 700
        assert shape.effective_height == 800

    def it_returns_None_for_effective_values_when_raw_values_are_None(self):
        # -- a:xfrm missing entirely on the shape: raw left/top/etc. are None, --
        # -- and the effective_* variants preserve that.                      --
        sp = cast("ShapeElement", element("p:sp/p:spPr"))
        shape = BaseShape(sp, None)

        assert shape.left is None
        assert shape.effective_left is None
        assert shape.effective_top is None
        assert shape.effective_width is None
        assert shape.effective_height is None

    def it_composites_the_group_shape_own_position(self):
        # -- a GroupShape nested inside another group also has its own off/ext --
        # -- that must be transformed to be slide-relative.                    --
        spTree = cast(
            "ShapeElement",
            element(
                "p:spTree/(p:nvGrpSpPr,p:grpSpPr,"
                "p:grpSp/(p:nvGrpSpPr,p:grpSpPr/a:xfrm/("
                "a:off{x=0,y=0},a:ext{cx=2000,cy=2000},"
                "a:chOff{x=0,y=0},a:chExt{cx=4000,cy=4000}),"
                "p:grpSp/(p:nvGrpSpPr,p:grpSpPr/a:xfrm/("
                "a:off{x=1000,y=1000},a:ext{cx=2000,cy=2000},"
                "a:chOff{x=0,y=0},a:chExt{cx=2000,cy=2000}))))"
            ),
        )
        # -- inner group (second p:grpSp, index 1 within outer) --
        inner_grpSp = spTree.xpath(".//p:grpSp/p:grpSp")[0]
        shape = BaseShape(cast("ShapeElement", inner_grpSp), None)

        # -- raw inner off = (1000,1000), ext = (2000,2000) --
        assert shape.left == 1000
        assert shape.width == 2000

        # -- outer sx = 2000/4000 = 0.5 applied to (1000, 2000)               --
        # --   effective left = 0 + (1000 - 0) * 0.5 = 500                    --
        # --   effective width = 2000 * 0.5 = 1000                            --
        assert shape.effective_left == 500
        assert shape.effective_top == 500
        assert shape.effective_width == 1000
        assert shape.effective_height == 1000

    def it_knows_its_rotation_angle(self, rotation_get_fixture):
        shape, expected_value = rotation_get_fixture
        assert shape.rotation == expected_value

    def it_can_change_its_rotation_angle(self, rotation_set_fixture):
        shape, new_value, expected_xml = rotation_set_fixture
        shape.rotation = new_value
        assert shape._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("sp_cxml", "expected_value"),
        [
            ("p:sp/p:spPr", False),
            ("p:sp/p:spPr/a:xfrm", False),
            ("p:sp/p:spPr/a:xfrm{flipH=1}", True),
            ("p:sp/p:spPr/a:xfrm{flipH=0}", False),
            ("p:pic/p:spPr/a:xfrm{flipH=1}", True),
        ],
    )
    def it_knows_whether_it_is_flipped_horizontally(self, sp_cxml, expected_value):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        assert shape.flip_horizontal is expected_value

    @pytest.mark.parametrize(
        ("sp_cxml", "new_value", "expected_cxml"),
        [
            ("p:sp/p:spPr/a:xfrm", True, "p:sp/p:spPr/a:xfrm{flipH=1}"),
            ("p:sp/p:spPr/a:xfrm{flipH=1}", False, "p:sp/p:spPr/a:xfrm"),
            ("p:sp/p:spPr/a:xfrm{flipV=1}", True, "p:sp/p:spPr/a:xfrm{flipV=1,flipH=1}"),
        ],
    )
    def it_can_change_its_horizontal_flip(self, sp_cxml, new_value, expected_cxml):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        shape.flip_horizontal = new_value
        assert shape._element.xml == xml(expected_cxml)

    def it_creates_an_xfrm_when_setting_horizontal_flip_on_bare_spPr(self):
        """Setting flip_horizontal on a shape without an `a:xfrm` creates one."""
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)
        shape.flip_horizontal = True
        assert shape.flip_horizontal is True
        assert shape._element.find(qn("p:spPr")).find(qn("a:xfrm")) is not None

    @pytest.mark.parametrize(
        ("sp_cxml", "expected_value"),
        [
            ("p:sp/p:spPr", False),
            ("p:sp/p:spPr/a:xfrm", False),
            ("p:sp/p:spPr/a:xfrm{flipV=1}", True),
            ("p:sp/p:spPr/a:xfrm{flipV=0}", False),
            ("p:pic/p:spPr/a:xfrm{flipV=1}", True),
        ],
    )
    def it_knows_whether_it_is_flipped_vertically(self, sp_cxml, expected_value):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        assert shape.flip_vertical is expected_value

    @pytest.mark.parametrize(
        ("sp_cxml", "new_value", "expected_cxml"),
        [
            ("p:sp/p:spPr/a:xfrm", True, "p:sp/p:spPr/a:xfrm{flipV=1}"),
            ("p:sp/p:spPr/a:xfrm{flipV=1}", False, "p:sp/p:spPr/a:xfrm"),
            ("p:sp/p:spPr/a:xfrm{flipH=1}", True, "p:sp/p:spPr/a:xfrm{flipH=1,flipV=1}"),
        ],
    )
    def it_can_change_its_vertical_flip(self, sp_cxml, new_value, expected_cxml):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        shape.flip_vertical = new_value
        assert shape._element.xml == xml(expected_cxml)

    def it_creates_an_xfrm_when_setting_vertical_flip_on_bare_spPr(self):
        """Setting flip_vertical on a shape without an `a:xfrm` creates one."""
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)
        shape.flip_vertical = True
        assert shape.flip_vertical is True
        assert shape._element.find(qn("p:spPr")).find(qn("a:xfrm")) is not None

    def it_can_toggle_its_horizontal_flip(self):
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)

        shape.flip_horizontally()
        assert shape.flip_horizontal is True

        shape.flip_horizontally()
        assert shape.flip_horizontal is False

    def it_can_toggle_its_vertical_flip(self):
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)

        shape.flip_vertically()
        assert shape.flip_vertical is True

        shape.flip_vertically()
        assert shape.flip_vertical is False

    @pytest.mark.parametrize(
        ("sp_cxml", "expected_value"),
        [
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}", False),
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,hidden=0}", False),
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,hidden=1}", True),
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,hidden=true}", True),
            ("p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,hidden=1}", True),
            ("p:grpSp/p:nvGrpSpPr/p:cNvPr{id=1,name=A,hidden=1}", True),
            ("p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=1,name=A,hidden=1}", True),
            ("p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=1,name=A,hidden=1}", True),
        ],
    )
    def it_knows_whether_it_is_hidden(self, sp_cxml, expected_value):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        assert shape.is_hidden is expected_value

    @pytest.mark.parametrize(
        ("sp_cxml", "new_value", "expected_cxml"),
        [
            (
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}",
                True,
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,hidden=1}",
            ),
            (
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,hidden=1}",
                False,
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}",
            ),
            (
                "p:pic/p:nvPicPr/p:cNvPr{id=1,name=A}",
                True,
                "p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,hidden=1}",
            ),
        ],
    )
    def it_can_change_its_hidden_state(self, sp_cxml, new_value, expected_cxml):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        shape.is_hidden = new_value
        assert shape._element.xml == xml(expected_cxml)

    # -- #447: theme_style_refs read/write access -------------------------

    def it_returns_None_for_theme_style_refs_on_a_shape_without_p_style(self):
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)
        assert shape.theme_style_refs is None

    @pytest.mark.parametrize(
        ("sp_cxml", "expected"),
        [
            (
                "p:sp/p:style/(a:lnRef{idx=1},a:fillRef{idx=3},a:effectRef{idx=2},"
                "a:fontRef{idx=minor})",
                (1, 3, 2, "minor"),
            ),
            (
                "p:cxnSp/p:style/(a:lnRef{idx=2},a:fillRef{idx=0},a:effectRef{idx=1},"
                "a:fontRef{idx=minor})",
                (2, 0, 1, "minor"),
            ),
            (
                "p:pic/p:style/(a:lnRef{idx=0},a:fillRef{idx=0},a:effectRef{idx=0},"
                "a:fontRef{idx=none})",
                (0, 0, 0, "none"),
            ),
        ],
    )
    def it_reads_theme_style_refs_from_p_style(self, sp_cxml, expected):
        from pptx.shapes.base import ThemeStyleRefs

        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        refs = shape.theme_style_refs
        assert isinstance(refs, ThemeStyleRefs)
        assert (refs.line_ref, refs.fill_ref, refs.effect_ref, refs.font_ref) == expected

    def it_returns_None_for_theme_style_refs_on_a_graphicFrame(self):
        shape = BaseShape(
            cast("ShapeElement", element("p:graphicFrame/p:nvGraphicFramePr")), None
        )
        assert shape.theme_style_refs is None

    def it_returns_None_for_theme_style_refs_on_a_group_shape(self):
        shape = BaseShape(
            cast("ShapeElement", element("p:grpSp/p:nvGrpSpPr")), None
        )
        assert shape.theme_style_refs is None

    def it_can_set_theme_style_refs_on_an_sp(self):
        from pptx.shapes.base import ThemeStyleRefs

        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)

        shape.theme_style_refs = ThemeStyleRefs(1, 2, 3, "major")

        refs = shape.theme_style_refs
        assert refs is not None
        assert refs == (1, 2, 3, "major")

    def it_can_set_theme_style_refs_using_a_plain_tuple(self):
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)

        shape.theme_style_refs = (0, 5, 2, "minor")  # type: ignore[assignment]

        refs = shape.theme_style_refs
        assert refs is not None
        assert refs == (0, 5, 2, "minor")

    def it_replaces_any_existing_p_style_on_assignment(self):
        from pptx.shapes.base import ThemeStyleRefs

        shape = BaseShape(
            cast(
                "ShapeElement",
                element(
                    "p:sp/p:style/(a:lnRef{idx=9},a:fillRef{idx=9},a:effectRef{idx=9},"
                    "a:fontRef{idx=major})"
                ),
            ),
            None,
        )

        shape.theme_style_refs = ThemeStyleRefs(2, 3, 1, "minor")

        refs = shape.theme_style_refs
        assert refs == (2, 3, 1, "minor")
        # -- exactly one p:style child after replacement --
        assert len(shape._element.findall(qn("p:style"))) == 1

    def it_clears_p_style_when_assigned_None(self):
        shape = BaseShape(
            cast(
                "ShapeElement",
                element(
                    "p:sp/p:style/(a:lnRef{idx=1},a:fillRef{idx=3},a:effectRef{idx=2},"
                    "a:fontRef{idx=minor})"
                ),
            ),
            None,
        )

        shape.theme_style_refs = None

        assert shape.theme_style_refs is None
        assert shape._element.find(qn("p:style")) is None

    def it_inserts_p_style_before_txBody_on_an_sp(self):
        from pptx.shapes.base import ThemeStyleRefs

        shape = BaseShape(
            cast("ShapeElement", element("p:sp/(p:spPr,p:txBody)")), None
        )

        shape.theme_style_refs = ThemeStyleRefs(1, 2, 3, "major")

        # -- p:style must appear after p:spPr but before p:txBody --
        children = list(shape._element)
        tags = [c.tag for c in children]
        assert tags.index(qn("p:style")) < tags.index(qn("p:txBody"))
        assert tags.index(qn("p:spPr")) < tags.index(qn("p:style"))

    def it_raises_when_setting_theme_style_refs_on_a_graphicFrame(self):
        shape = BaseShape(
            cast("ShapeElement", element("p:graphicFrame/p:nvGraphicFramePr")), None
        )
        with pytest.raises(ValueError, match="does not support a p:style"):
            shape.theme_style_refs = (1, 2, 3, "minor")  # type: ignore[assignment]

    def it_raises_when_setting_theme_style_refs_on_a_group_shape(self):
        shape = BaseShape(
            cast("ShapeElement", element("p:grpSp/p:nvGrpSpPr")), None
        )
        with pytest.raises(ValueError, match="does not support a p:style"):
            shape.theme_style_refs = (1, 2, 3, "minor")  # type: ignore[assignment]

    @pytest.mark.parametrize(
        "bad_font_ref",
        ["Major", "regular", "", "MAJOR", "bold"],
    )
    def it_rejects_invalid_font_ref_strings(self, bad_font_ref):
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)
        with pytest.raises(ValueError, match="font_ref must be one of"):
            shape.theme_style_refs = (1, 2, 3, bad_font_ref)  # type: ignore[assignment]

    def it_rejects_negative_ref_indices(self):
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)
        with pytest.raises(ValueError, match="must be non-negative"):
            shape.theme_style_refs = (-1, 2, 3, "minor")  # type: ignore[assignment]

    def it_rejects_non_integer_ref_values(self):
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)
        with pytest.raises(TypeError, match="must be non-negative int"):
            shape.theme_style_refs = ("1", 2, 3, "minor")  # type: ignore[assignment]

    # -- #508: alt_text / title accessibility description ------------------

    @pytest.mark.parametrize(
        ("sp_cxml", "expected_value"),
        [
            # -- defaults to empty string when `descr` is absent --
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}", ""),
            # -- round-trips the `descr` value for every shape element --
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,descr=a photo}", "a photo"),
            ("p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,descr=a photo}", "a photo"),
            ("p:grpSp/p:nvGrpSpPr/p:cNvPr{id=1,name=A,descr=a photo}", "a photo"),
            ("p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=1,name=A,descr=a photo}", "a photo"),
            (
                "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=1,name=A,descr=a photo}",
                "a photo",
            ),
        ],
    )
    def it_knows_its_alt_text(self, sp_cxml, expected_value):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        assert shape.alt_text == expected_value

    @pytest.mark.parametrize(
        ("sp_cxml", "new_value", "expected_cxml"),
        [
            # -- add descr to a cNvPr that doesn't have one --
            (
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}",
                "chart of Q3 revenue",
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,descr=chart of Q3 revenue}",
            ),
            # -- overwrite an existing descr value --
            (
                "p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,descr=old text}",
                "new text",
                "p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,descr=new text}",
            ),
            # -- assigning the empty string removes the attribute (matches default) --
            (
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,descr=something}",
                "",
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}",
            ),
        ],
    )
    def it_can_change_its_alt_text(self, sp_cxml, new_value, expected_cxml):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        shape.alt_text = new_value
        assert shape._element.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("sp_cxml", "expected_value"),
        [
            # -- defaults to empty string when `title` is absent --
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}", ""),
            # -- round-trips the `title` value across shape element types --
            ("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,title=Short title}", "Short title"),
            ("p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,title=Short title}", "Short title"),
            (
                "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=1,name=A,title=Short title}",
                "Short title",
            ),
        ],
    )
    def it_knows_its_title(self, sp_cxml, expected_value):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        assert shape.title == expected_value

    @pytest.mark.parametrize(
        ("sp_cxml", "new_value", "expected_cxml"),
        [
            (
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}",
                "Q3 chart",
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,title=Q3 chart}",
            ),
            (
                "p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,title=Old title}",
                "New title",
                "p:pic/p:nvPicPr/p:cNvPr{id=1,name=A,title=New title}",
            ),
            # -- assigning the empty string removes the attribute (matches default) --
            (
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A,title=Something}",
                "",
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}",
            ),
        ],
    )
    def it_can_change_its_title(self, sp_cxml, new_value, expected_cxml):
        shape = BaseShape(cast("ShapeElement", element(sp_cxml)), None)
        shape.title = new_value
        assert shape._element.xml == xml(expected_cxml)

    def it_preserves_alt_text_and_title_separately(self):
        """Setting one accessibility attribute must not disturb the other."""
        shape = BaseShape(cast("ShapeElement", element("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")), None)

        shape.alt_text = "a long description of the picture contents"
        shape.title = "short title"

        assert shape.alt_text == "a long description of the picture contents"
        assert shape.title == "short title"

        # -- clearing one leaves the other intact --
        shape.alt_text = ""
        assert shape.alt_text == ""
        assert shape.title == "short title"

    def it_provides_access_to_its_shadow(self, shadow_fixture):
        shape, ShadowFormat_, spPr, shadow_ = shadow_fixture

        shadow = shape.shadow

        ShadowFormat_.assert_called_once_with(spPr)
        assert shadow is shadow_

    def it_exposes_the_full_ShadowFormat_api_through_shape_shadow(self):
        """Issue #130: Shape.shadow exposes full blur/distance/direction/color API.

        This regression test exercises the real |ShadowFormat| (no mock) so a regression
        that removes, renames, or gates any of the four outer-shadow properties behind the
        skeletal `.inherit` bool is caught immediately.
        """
        from pptx.dml.color import ColorFormat
        from pptx.util import Emu

        shape = BaseShape(cast("ShapeElement", element("p:sp/p:spPr")), None)
        shadow = shape.shadow

        # -- baseline: nothing configured, all four properties are None / inherit=True --
        assert shadow.inherit is True
        assert shadow.blur_radius is None
        assert shadow.distance is None
        assert shadow.direction is None

        # -- each property round-trips independently --
        shadow.blur_radius = Emu(50800)
        shadow.distance = Emu(38100)
        shadow.direction = 270.0

        assert shadow.blur_radius == Emu(50800)
        assert shadow.distance == Emu(38100)
        assert shadow.direction == 270.0
        # -- inherit now reports False, since an explicit a:effectLst exists --
        assert shadow.inherit is False
        # -- first access to `.color` materializes a default color-choice child --
        assert isinstance(shadow.color, ColorFormat)

    def it_knows_the_part_it_belongs_to(self, part_fixture):
        shape, parent_ = part_fixture
        part = shape.part
        assert part is parent_.part

    def it_knows_it_doesnt_have_a_text_frame(self):
        shape = BaseShape(None, None)
        assert shape.has_text_frame is False

    def it_computes_its_text_frame_rect_from_its_extents_and_margins(self):
        # -- textbox at (1", 2") sized 4" x 1" — default insets are
        # -- 0.1" l/r (91440 EMU each), 0.05" t/b (45720 EMU each) on a
        # -- fresh a:bodyPr.
        sp_cxml = (
            "p:sp/(p:spPr/a:xfrm/(a:off{x=914400,y=1828800},a:ext{cx=3657600,"
            "cy=914400}),p:txBody/(a:bodyPr,a:p))"
        )
        shape = Shape(cast("ShapeElement", element(sp_cxml)), None)

        rect = shape.text_frame_rect

        # -- TextFrameRect is a NamedTuple of 4 Length values
        assert rect.left == 914400 + 91440  # 1" + 0.1"
        assert rect.top == 1828800 + 45720  # 2" + 0.05"
        assert rect.width == 3657600 - 2 * 91440  # 4" - 0.2"
        assert rect.height == 914400 - 2 * 45720  # 1" - 0.1"
        # -- tuple unpacking and EMU readback both work
        left, top, width, height = rect
        assert left.emu == 1005840
        assert top.emu == 1874520
        assert width.emu == 3474720
        assert height.emu == 822960

    def it_honors_explicit_margins_when_computing_text_frame_rect(self):
        # -- 2" x 2" shape at origin with explicit 1/2" insets on every side
        sp_cxml = (
            "p:sp/(p:spPr/a:xfrm/(a:off{x=0,y=0},a:ext{cx=1828800,cy=1828800"
            "}),p:txBody/(a:bodyPr{lIns=457200,tIns=457200,rIns=457200,bIns="
            "457200},a:p))"
        )
        shape = Shape(cast("ShapeElement", element(sp_cxml)), None)

        rect = shape.text_frame_rect

        assert rect.left == Inches(0.5)
        assert rect.top == Inches(0.5)
        assert rect.width == Inches(1.0)
        assert rect.height == Inches(1.0)

    def it_raises_on_text_frame_rect_when_no_text_frame(self):
        shape = BaseShape(None, None)
        with pytest.raises(ValueError, match="shape has no text frame"):
            shape.text_frame_rect

    def it_can_delete_itself_from_its_parent(self):
        spTree = cast("ShapeElement", element("p:spTree/(p:sp,p:sp,p:sp)"))
        sps = spTree.xpath("p:sp")
        shape = BaseShape(cast("ShapeElement", sps[1]), None)

        shape.delete()

        remaining = spTree.xpath("p:sp")
        assert len(remaining) == 2
        assert sps[1] not in remaining
        assert sps[0] in remaining
        assert sps[2] in remaining

    def it_can_replace_itself_with_another_shape(self):
        spTree = cast(
            "ShapeElement",
            element(
                "p:spTree/("
                "p:sp/p:spPr/a:xfrm/(a:off{x=100,y=200},a:ext{cx=300,cy=400}),"
                "p:sp/p:spPr/a:xfrm/(a:off{x=11,y=22},a:ext{cx=33,cy=44})"
                ")"
            ),
        )
        old_sp, new_sp = spTree.xpath("p:sp")
        old_shape = BaseShape(cast("ShapeElement", old_sp), None)
        new_shape = BaseShape(cast("ShapeElement", new_sp), None)

        old_shape.replace_with(new_shape)

        # -- new_shape now occupies the old_shape's position/size --
        assert new_shape.left == 100
        assert new_shape.top == 200
        assert new_shape.width == 300
        assert new_shape.height == 400
        # -- old_shape is gone from the tree, new_shape is in its slot (z-order first) --
        remaining = spTree.xpath("p:sp")
        assert len(remaining) == 1
        assert remaining[0] is new_sp
        # -- new_shape moved to old_shape's slot (index 0) --
        assert list(spTree).index(new_sp) == 0

    def it_raises_on_replace_with_self(self):
        spTree = cast("ShapeElement", element("p:spTree/p:sp/p:spPr"))
        sp = spTree.xpath("p:sp")[0]
        shape = BaseShape(cast("ShapeElement", sp), None)

        with pytest.raises(ValueError, match="replace a shape with itself"):
            shape.replace_with(shape)

    def it_raises_on_replace_when_other_is_unparented(self):
        spTree = cast("ShapeElement", element("p:spTree/p:sp/p:spPr"))
        sp = spTree.xpath("p:sp")[0]
        orphan = cast("ShapeElement", element("p:sp/p:spPr"))
        old_shape = BaseShape(cast("ShapeElement", sp), None)
        new_shape = BaseShape(orphan, None)

        with pytest.raises(ValueError, match="other_shape has no parent"):
            old_shape.replace_with(new_shape)

    def it_knows_whether_it_is_a_placeholder(self, is_placeholder_fixture):
        shape, is_placeholder = is_placeholder_fixture
        assert shape.is_placeholder is is_placeholder

    def it_provides_access_to_its_placeholder_format(self, phfmt_fixture):
        shape, _PlaceholderFormat_, placeholder_format_, ph = phfmt_fixture
        placeholder_format = shape.placeholder_format
        _PlaceholderFormat_.assert_called_once_with(ph)
        assert placeholder_format is placeholder_format_

    def it_raises_when_shape_is_not_a_placeholder(self, phfmt_raise_fixture):
        shape = phfmt_raise_fixture
        with pytest.raises(ValueError):
            shape.placeholder_format

    def it_knows_it_doesnt_contain_a_chart(self):
        shape = BaseShape(None, None)
        assert shape.has_chart is False

    def it_knows_it_doesnt_contain_a_table(self):
        shape = BaseShape(None, None)
        assert shape.has_table is False

    def it_reports_its_zorder_index(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B},p:sp/p:nvSpPr/p:cNvPr{id=4,name=C})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shapes = [BaseShape(cast("ShapeElement", sp), None) for sp in sp_elms]

        assert [s.zorder_index for s in shapes] == [0, 1, 2]

    def it_reports_zorder_index_ignoring_non_shape_siblings(self):
        # -- p:extLst is a valid trailing sibling but must not count as a shape --
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:pic/p:nvPicPr/p:cNvPr{id=3,name=B},p:extLst)"
        )
        sp_elm = spTree.find(qn("p:sp"))
        pic_elm = spTree.find(qn("p:pic"))

        assert BaseShape(cast("ShapeElement", sp_elm), None).zorder_index == 0
        assert BaseShape(cast("ShapeElement", pic_elm), None).zorder_index == 1

    def it_raises_on_zorder_index_when_shape_has_no_parent(self):
        sp = cast("ShapeElement", element("p:sp/p:nvSpPr/p:cNvPr{id=2,name=A}"))
        shape = BaseShape(sp, None)
        with pytest.raises(ValueError, match="no parent"):
            shape.zorder_index

    def it_can_bring_a_shape_to_the_front(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B},p:sp/p:nvSpPr/p:cNvPr{id=4,name=C})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_B = BaseShape(cast("ShapeElement", sp_elms[1]), None)

        shape_B.bring_to_front()

        names = [sp.xpath("./p:nvSpPr/p:cNvPr/@name")[0] for sp in spTree.findall(qn("p:sp"))]
        assert names == ["A", "C", "B"]
        assert shape_B.zorder_index == 2

    def it_preserves_trailing_extLst_on_bring_to_front(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B},p:extLst)"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_A = BaseShape(cast("ShapeElement", sp_elms[0]), None)

        shape_A.bring_to_front()

        # -- A is now last shape, but p:extLst remains the final child --
        assert spTree[-1].tag == qn("p:extLst")
        names = [sp.xpath("./p:nvSpPr/p:cNvPr/@name")[0] for sp in spTree.findall(qn("p:sp"))]
        assert names == ["B", "A"]

    def it_is_a_noop_when_bring_to_front_on_frontmost_shape(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_B = BaseShape(cast("ShapeElement", sp_elms[1]), None)
        before_xml = spTree.xml

        shape_B.bring_to_front()

        assert spTree.xml == before_xml

    def it_can_send_a_shape_to_the_back(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B},p:sp/p:nvSpPr/p:cNvPr{id=4,name=C})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_C = BaseShape(cast("ShapeElement", sp_elms[2]), None)

        shape_C.send_to_back()

        names = [sp.xpath("./p:nvSpPr/p:cNvPr/@name")[0] for sp in spTree.findall(qn("p:sp"))]
        assert names == ["C", "A", "B"]
        assert shape_C.zorder_index == 0

    def it_is_a_noop_when_send_to_back_on_backmost_shape(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_A = BaseShape(cast("ShapeElement", sp_elms[0]), None)
        before_xml = spTree.xml

        shape_A.send_to_back()

        assert spTree.xml == before_xml

    def it_can_bring_a_shape_forward_one_step(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B},p:sp/p:nvSpPr/p:cNvPr{id=4,name=C},"
            "p:sp/p:nvSpPr/p:cNvPr{id=5,name=D})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_B = BaseShape(cast("ShapeElement", sp_elms[1]), None)

        shape_B.bring_forward()

        names = [sp.xpath("./p:nvSpPr/p:cNvPr/@name")[0] for sp in spTree.findall(qn("p:sp"))]
        assert names == ["A", "C", "B", "D"]
        assert shape_B.zorder_index == 2

    def it_is_a_noop_when_bring_forward_on_frontmost_shape(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_B = BaseShape(cast("ShapeElement", sp_elms[1]), None)
        before_xml = spTree.xml

        shape_B.bring_forward()

        assert spTree.xml == before_xml

    def it_can_send_a_shape_backward_one_step(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B},p:sp/p:nvSpPr/p:cNvPr{id=4,name=C},"
            "p:sp/p:nvSpPr/p:cNvPr{id=5,name=D})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_C = BaseShape(cast("ShapeElement", sp_elms[2]), None)

        shape_C.send_backward()

        names = [sp.xpath("./p:nvSpPr/p:cNvPr/@name")[0] for sp in spTree.findall(qn("p:sp"))]
        assert names == ["A", "C", "B", "D"]
        assert shape_C.zorder_index == 1

    def it_is_a_noop_when_send_backward_on_backmost_shape(self):
        spTree = element(
            "p:spTree/(p:nvGrpSpPr,p:grpSpPr,p:sp/p:nvSpPr/p:cNvPr{id=2,name=A},"
            "p:sp/p:nvSpPr/p:cNvPr{id=3,name=B})"
        )
        sp_elms = spTree.findall(qn("p:sp"))
        shape_A = BaseShape(cast("ShapeElement", sp_elms[0]), None)
        before_xml = spTree.xml

        shape_A.send_backward()

        assert spTree.xml == before_xml

    def it_supports_zorder_on_shapes_in_a_group(self):
        # -- z-order operates on siblings within the same parent (grpSp or spTree),
        # -- so shapes inside a group reorder only within that group.
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        group = slide.shapes.add_group_shape()
        s1 = group.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        s2 = group.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(2), Inches(2), Inches(1), Inches(1)
        )
        s3 = group.shapes.add_shape(
            MSO_SHAPE.DIAMOND, Inches(3), Inches(3), Inches(1), Inches(1)
        )

        assert [s1.zorder_index, s2.zorder_index, s3.zorder_index] == [0, 1, 2]

        s1.bring_to_front()
        assert [s.shape_id for s in group.shapes] == [s2.shape_id, s3.shape_id, s1.shape_id]
        assert s1.zorder_index == 2

        s1.send_to_back()
        assert [s.shape_id for s in group.shapes] == [s1.shape_id, s2.shape_id, s3.shape_id]
        assert s1.zorder_index == 0

        s1.bring_forward()
        assert [s.shape_id for s in group.shapes] == [s2.shape_id, s1.shape_id, s3.shape_id]

        s1.send_backward()
        assert [s.shape_id for s in group.shapes] == [s1.shape_id, s2.shape_id, s3.shape_id]

    def it_round_trips_zorder_via_slide_shapes_index(self):
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        s2 = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(2), Inches(2), Inches(1), Inches(1)
        )
        s3 = slide.shapes.add_shape(
            MSO_SHAPE.DIAMOND, Inches(3), Inches(3), Inches(1), Inches(1)
        )
        base = slide.shapes.index(s1)
        assert [s1.zorder_index, s2.zorder_index, s3.zorder_index] == [
            base,
            base + 1,
            base + 2,
        ]

        s3.send_to_back()
        assert s3.zorder_index == 0
        assert slide.shapes.index(s3) == 0
    def it_can_duplicate_an_autoshape(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )
        shape.text = "hello"
        orig_count = len(slide.shapes)

        dup = shape.duplicate()

        assert len(slide.shapes) == orig_count + 1
        assert isinstance(dup, Shape)
        assert dup is not shape
        assert dup._element is not shape._element
        assert dup.shape_id != shape.shape_id
        assert dup.shape_id > shape.shape_id
        assert dup.name != shape.name
        assert dup.name.startswith(shape.name)
        assert dup.text_frame.text == "hello"
        assert dup.left == shape.left
        assert dup.top == shape.top
        assert dup.width == shape.width
        assert dup.height == shape.height
        # -- dup is the last shape in z-order --
        assert slide.shapes[-1]._element is dup._element

    def it_can_duplicate_a_textbox(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        tb.text_frame.text = "tbx"

        dup = tb.duplicate()

        assert isinstance(dup, Shape)
        assert dup.shape_id != tb.shape_id
        assert dup.text_frame.text == "tbx"

    def it_can_duplicate_a_connector(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR_TYPE.STRAIGHT, Inches(1), Inches(1), Inches(3), Inches(1)
        )

        dup = cast("Connector", conn.duplicate())

        assert type(dup).__name__ == "Connector"
        assert dup.shape_id != conn.shape_id
        assert dup.begin_x == conn.begin_x
        assert dup.end_x == conn.end_x

    def it_assigns_unique_name_to_duplicate(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )
        base_name = shape.name

        dup1 = shape.duplicate()
        dup2 = shape.duplicate()
        dup3 = dup1.duplicate()

        names = {shape.name, dup1.name, dup2.name, dup3.name}
        assert len(names) == 4
        # -- all duplicate names derive from the original name --
        for d in (dup1, dup2, dup3):
            assert d.name.startswith(base_name)

    def it_raises_when_duplicating_a_picture(self):
        from tests.unitutil.file import test_file_dir

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        pic = slide.shapes.add_picture("%s/python-icon.jpeg" % test_file_dir, Inches(1), Inches(1))
        with pytest.raises(NotImplementedError, match="simple shapes"):
            pic.duplicate()

    def it_raises_when_duplicating_a_graphic_frame(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        gf = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(3), Inches(1))
        with pytest.raises(NotImplementedError, match="simple shapes"):
            gf.duplicate()

    def it_can_duplicate_an_empty_group_shape(self):
        # -- #1085: duplicating a group is now supported; duplicating an empty --
        # -- group returns a fresh GroupShape with a new id and unique name.  --
        from pptx.shapes.group import GroupShape

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        gs = slide.shapes.add_group_shape()
        orig_count = len(slide.shapes)

        dup = gs.duplicate()

        assert isinstance(dup, GroupShape)
        assert len(slide.shapes) == orig_count + 1
        assert dup.shape_id != gs.shape_id
        assert dup.name != gs.name
        # -- dup is the last shape in z-order --
        assert slide.shapes[-1]._element is dup._element

    def it_raises_when_duplicating_a_placeholder(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        title = slide.shapes.title
        assert title is not None
        with pytest.raises(NotImplementedError, match="placeholder"):
            title.duplicate()
    @pytest.mark.parametrize(
        ("shape_cxml", "expected_value"),
        [
            # -- no OMML at all --
            ("p:sp/p:txBody/a:p", False),
            # -- text-frame with regular text but no equation --
            ("p:sp/p:txBody/a:p/a:r/a:t", False),
            # -- inline m:oMath in a paragraph (typical choice-wrapped shape) --
            ("p:sp/p:txBody/a:p/m:oMath", True),
            # -- m:oMath inside m:oMathPara --
            ("p:sp/p:txBody/a:p/m:oMathPara/m:oMath", True),
            # -- m:oMath as a direct descendant, no paragraph wrapper --
            ("p:sp/m:oMath", True),
            # -- non-shape element types also work (graphicFrame, pic, grpSp, cxnSp) --
            ("p:pic/p:nvPicPr", False),
            ("p:graphicFrame/m:oMath", True),
        ],
    )
    def it_knows_whether_it_contains_an_OMML_equation(
        self, shape_cxml: str, expected_value: bool
    ):
        shape_elm = cast("ShapeElement", element(shape_cxml))
        shape = BaseShape(shape_elm, None)
        assert shape.has_math_equation is expected_value

    def it_returns_None_for_math_equation_xml_when_no_equation(self):
        shape_elm = cast("ShapeElement", element("p:sp/p:txBody/a:p"))
        shape = BaseShape(shape_elm, None)
        assert shape.math_equation_xml is None

    def it_returns_the_raw_OMML_XML_for_the_first_equation(self):
        shape_elm = cast("ShapeElement", element("p:sp/p:txBody/a:p/m:oMath"))
        shape = BaseShape(shape_elm, None)

        oMath_xml = shape.math_equation_xml

        assert oMath_xml is not None
        assert oMath_xml.startswith("<m:oMath")
        assert 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' in oMath_xml

    def and_it_returns_only_the_first_equation_when_multiple_are_present(self):
        shape_elm = cast(
            "ShapeElement",
            element("p:sp/p:txBody/(a:p/m:oMath,a:p/m:oMath)"),
        )
        shape = BaseShape(shape_elm, None)

        oMath_xml = shape.math_equation_xml

        assert oMath_xml is not None
        # -- exactly one `<m:oMath` open-tag (the first), not two --
        assert oMath_xml.count("<m:oMath") == 1

    def it_surfaces_an_OMML_equation_on_a_real_slide_regression_892(
        self
    ):
        """Regression test for issue #892 ("Support for parsing Equations").

        The user-visible request in #892 was "given a .pptx, discover and
        access the math equations a slide contains." Issue #126 delivered
        the MVP API that resolves that use case:

        - iterate ``slide.shapes`` (Foundation F3 surfaces equation-bearing
          shapes wrapped in ``mc:AlternateContent`` transparently);
        - use ``shape.has_math_equation`` to detect equations;
        - use ``shape.math_equation_xml`` to read the raw OMML subtree.

        Structured parsing of OMML (typed ``Equation`` proxy, LaTeX / MathML
        conversion, programmatic authoring) remains **explicitly deferred**
        -- see ``docs/dev/analysis/omml-parsing.rst``.
        """
        from os.path import abspath, dirname, join

        pptx_path = abspath(
            join(
                dirname(__file__),
                "..",
                "..",
                "features",
                "steps",
                "test_files",
                "shp-math-equation.pptx",
            )
        )
        prs = Presentation(pptx_path)
        slide = prs.slides[0]

        # -- the #892 "find-equation-in-slide" use case: iterate shapes and
        # -- pull out those that carry an equation --
        equations = [
            shape.math_equation_xml
            for shape in slide.shapes
            if shape.has_math_equation
        ]

        # -- exactly one equation is present in the fixture --
        assert len(equations) == 1
        oMath_xml = equations[0]
        assert oMath_xml is not None
        # -- the caller receives a well-formed OMML fragment they can hand to
        # -- an external converter (pandoc, omml.xsl, etc.) --
        assert oMath_xml.startswith("<m:oMath")
        assert (
            'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
            in oMath_xml
        )
        # -- non-equation shapes on the same slide cleanly report None --
        non_eq = [
            shape for shape in slide.shapes if not shape.has_math_equation
        ]
        assert len(non_eq) >= 1
        for shape in non_eq:
            assert shape.math_equation_xml is None

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            "p:sp/p:nvSpPr/p:cNvPr",
            "p:grpSp/p:nvGrpSpPr/p:cNvPr",
            "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr",
            "p:cxnSp/p:nvCxnSpPr/p:cNvPr",
            "p:pic/p:nvPicPr/p:cNvPr",
        ]
    )
    def click_action_fixture(self, request, ActionSetting_, action_setting_):
        sp_cxml = request.param
        sp = element(sp_cxml)
        cNvPr = sp.xpath("//p:cNvPr")[0]
        shape = BaseShape(sp, None)
        return shape, ActionSetting_, cNvPr, action_setting_

    @pytest.fixture(
        params=[
            "p:sp/p:nvSpPr/p:cNvPr",
            "p:grpSp/p:nvGrpSpPr/p:cNvPr",
            "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr",
            "p:cxnSp/p:nvCxnSpPr/p:cNvPr",
            "p:pic/p:nvPicPr/p:cNvPr",
        ]
    )
    def hover_action_fixture(self, request, ActionSetting_, action_setting_):
        sp_cxml = request.param
        sp = element(sp_cxml)
        cNvPr = sp.xpath("//p:cNvPr")[0]
        shape = BaseShape(sp, None)
        return shape, ActionSetting_, cNvPr, action_setting_

    @pytest.fixture(
        params=[
            ("sp", False),
            ("sp_with_ext", True),
            ("pic", False),
            ("pic_with_ext", True),
            ("graphicFrame", False),
            ("graphicFrame_with_ext", True),
            ("grpSp", False),
            ("grpSp_with_ext", True),
            ("cxnSp", False),
            ("cxnSp_with_ext", True),
        ]
    )
    def dimensions_get_fixture(self, request, width, height):
        shape_elm_fixt_name, expect_values = request.param
        shape_elm = request.getfixturevalue(shape_elm_fixt_name)
        shape = BaseShape(shape_elm, None)
        if not expect_values:
            width = height = None
        return shape, width, height

    @pytest.fixture(
        params=[
            ("sp", "sp_with_ext"),
            ("pic", "pic_with_ext"),
            ("graphicFrame", "graphicFrame_with_ext"),
            ("grpSp", "grpSp_with_ext"),
            ("cxnSp", "cxnSp_with_ext"),
        ]
    )
    def dimensions_set_fixture(self, request, width, height):
        start_elm_fixt_name, expected_elm_fixt_name = request.param
        start_elm = request.getfixturevalue(start_elm_fixt_name)
        shape = BaseShape(start_elm, None)
        expected_xml = request.getfixturevalue(expected_elm_fixt_name).xml
        return shape, width, height, expected_xml

    @pytest.fixture(
        params=[
            ("p:sp/p:nvSpPr/p:cNvPr{id=1}", 1),
            ("p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=2}", 2),
            ("p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=3}", 3),
            ("p:grpSp/p:nvGrpSpPr/p:cNvPr{id=4}", 4),
            ("p:pic/p:nvPicPr/p:cNvPr{id=5}", 5),
        ]
    )
    def id_fixture(self, request):
        xSp_cxml, expected_value = request.param
        shape = BaseShape(element(xSp_cxml), None)
        return shape, expected_value

    @pytest.fixture(params=[True, False])
    def is_placeholder_fixture(self, request, shape_elm_, txBody_):
        is_placeholder = request.param
        shape_elm_.has_ph_elm = is_placeholder
        shape = BaseShape(shape_elm_, None)
        return shape, is_placeholder

    @pytest.fixture
    def name_get_fixture(self, shape_name):
        shape_elm = (
            an_sp()
            .with_nsdecls()
            .with_child(an_nvSpPr().with_child(a_cNvPr().with_name(shape_name)))
        ).element
        shape = BaseShape(shape_elm, None)
        return shape, shape_name

    @pytest.fixture(
        params=[
            (
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=foo}",
                Shape,
                "Shape1",
                "p:sp/p:nvSpPr/p:cNvPr{id=1,name=Shape1}",
            ),
            (
                "p:grpSp/p:nvGrpSpPr/p:cNvPr{id=2,name=bar}",
                BaseShape,
                "Shape2",
                "p:grpSp/p:nvGrpSpPr/p:cNvPr{id=2,name=Shape2}",
            ),
            (
                "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=3,name=baz}",
                GraphicFrame,
                "Shape3",
                "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=3,name=Shape3}",
            ),
            (
                "p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=4,name=boo}",
                BaseShape,
                "Shape4",
                "p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=4,name=Shape4}",
            ),
            (
                "p:pic/p:nvPicPr/p:cNvPr{id=5,name=far}",
                Picture,
                "Shape5",
                "p:pic/p:nvPicPr/p:cNvPr{id=5,name=Shape5}",
            ),
        ]
    )
    def name_set_fixture(self, request):
        xSp_cxml, ShapeCls, new_value, expected_xSp_cxml = request.param
        shape = ShapeCls(element(xSp_cxml), None)
        expected_xml = xml(expected_xSp_cxml)
        return shape, new_value, expected_xml

    @pytest.fixture
    def part_fixture(self, shapes_):
        parent_ = shapes_
        shape = BaseShape(None, parent_)
        return shape, parent_

    @pytest.fixture
    def phfmt_fixture(self, _PlaceholderFormat_, placeholder_format_):
        sp = element("p:sp/p:nvSpPr/p:nvPr/p:ph")
        ph = sp.xpath("//p:ph")[0]
        shape = BaseShape(sp, None)
        return shape, _PlaceholderFormat_, placeholder_format_, ph

    @pytest.fixture
    def phfmt_raise_fixture(self):
        return BaseShape(element("p:sp/p:nvSpPr/p:nvPr"), None)

    @pytest.fixture(
        params=[
            ("sp", "sp_with_off"),
            ("pic", "pic_with_off"),
            ("graphicFrame", "graphicFrame_with_off"),
            ("grpSp", "grpSp_with_off"),
            ("cxnSp", "cxnSp_with_off"),
        ]
    )
    def position_set_fixture(self, request, left, top):
        start_elm_fixt_name, expected_elm_fixt_name = request.param
        start_elm = request.getfixturevalue(start_elm_fixt_name)
        shape = BaseShape(start_elm, None)
        expected_xml = request.getfixturevalue(expected_elm_fixt_name).xml
        return shape, left, top, expected_xml

    @pytest.fixture(
        params=[
            ("p:sp/p:spPr", 0.0),
            ("p:sp/p:spPr/a:xfrm{rot=60000}", 1.0),
            ("p:sp/p:spPr/a:xfrm{rot=2545200}", 42.42),
            ("p:sp/p:spPr/a:xfrm{rot=-60000}", 359.0),
            ("p:grpSp/p:grpSpPr/a:xfrm{rot=2545200}", 42.42),
        ]
    )
    def rotation_get_fixture(self, request):
        xSp_cxml, expected_value = request.param
        shape = BaseShapeFactory(element(xSp_cxml), None)
        return shape, expected_value

    @pytest.fixture(
        params=[
            ("p:sp/p:spPr/a:xfrm", 1.0, "p:sp/p:spPr/a:xfrm{rot=60000}"),
            ("p:sp/p:spPr/a:xfrm{rot=60000}", 0.0, "p:sp/p:spPr/a:xfrm"),
            (
                "p:sp/p:spPr/a:xfrm{rot=60000}",
                -420.0,
                "p:sp/p:spPr/a:xfrm{rot=18000000}",
            ),
            ("p:grpSp/p:grpSpPr/a:xfrm", 1.0, "p:grpSp/p:grpSpPr/a:xfrm{rot=60000}"),
        ]
    )
    def rotation_set_fixture(self, request):
        xSp_cxml, new_value, expected_xSp_cxml = request.param
        shape = BaseShapeFactory(element(xSp_cxml), None)
        expected_xml = xml(expected_xSp_cxml)
        return shape, new_value, expected_xml

    @pytest.fixture(
        params=[
            "p:sp/p:spPr",
            "p:cxnSp/p:spPr",
            "p:pic/p:spPr",
            # ---group and graphic frame shapes override this property---
        ]
    )
    def shadow_fixture(self, request, ShadowFormat_, shadow_):
        sp_cxml = request.param
        sp = element(sp_cxml)
        spPr = sp.xpath("//p:spPr")[0]
        ShadowFormat_.return_value = shadow_

        shape = BaseShape(sp, None)
        return shape, ShadowFormat_, spPr, shadow_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ActionSetting_(self, request, action_setting_):
        return class_mock(request, "pptx.shapes.base.ActionSetting", return_value=action_setting_)

    @pytest.fixture
    def action_setting_(self, request):
        return instance_mock(request, ActionSetting)

    @pytest.fixture
    def cxnSp(self):
        return a_cxnSp().with_nsdecls().with_child(an_spPr()).element

    @pytest.fixture
    def cxnSp_with_ext(self, width, height):
        return (
            a_cxnSp()
            .with_nsdecls()
            .with_child(
                an_spPr().with_child(an_xfrm().with_child(an_ext().with_cx(width).with_cy(height)))
            )
        ).element

    @pytest.fixture
    def cxnSp_with_off(self, left, top):
        return (
            a_cxnSp()
            .with_nsdecls()
            .with_child(
                an_spPr().with_child(an_xfrm().with_child(an_off().with_x(left).with_y(top)))
            )
        ).element

    @pytest.fixture
    def graphicFrame(self):
        # Note that <p:xfrm> element is required on graphicFrame
        return a_graphicFrame().with_nsdecls().with_child(a_p_xfrm()).element

    @pytest.fixture
    def graphicFrame_with_ext(self, width, height):
        return (
            a_graphicFrame()
            .with_nsdecls()
            .with_child(a_p_xfrm().with_child(an_ext().with_cx(width).with_cy(height)))
        ).element

    @pytest.fixture
    def graphicFrame_with_off(self, left, top):
        return (
            a_graphicFrame()
            .with_nsdecls()
            .with_child(a_p_xfrm().with_child(an_off().with_x(left).with_y(top)))
        ).element

    @pytest.fixture
    def grpSp(self):
        return (a_grpSp().with_nsdecls("p", "a").with_child(a_grpSpPr())).element

    @pytest.fixture
    def grpSp_with_ext(self, width, height):
        return (
            a_grpSp()
            .with_nsdecls("p", "a")
            .with_child(
                a_grpSpPr().with_child(
                    an_xfrm().with_child(an_ext().with_cx(width).with_cy(height))
                )
            )
        ).element

    @pytest.fixture
    def grpSp_with_off(self, left, top):
        return (
            a_grpSp()
            .with_nsdecls("p", "a")
            .with_child(
                a_grpSpPr().with_child(an_xfrm().with_child(an_off().with_x(left).with_y(top)))
            )
        ).element

    @pytest.fixture
    def height(self):
        return 654

    @pytest.fixture
    def left(self):
        return 123

    @pytest.fixture
    def pic(self):
        return a_pic().with_nsdecls().with_child(an_spPr()).element

    @pytest.fixture
    def pic_with_off(self, left, top):
        return (
            a_pic()
            .with_nsdecls()
            .with_child(
                an_spPr().with_child(an_xfrm().with_child(an_off().with_x(left).with_y(top)))
            )
        ).element

    @pytest.fixture
    def pic_with_ext(self, width, height):
        return (
            a_pic()
            .with_nsdecls()
            .with_child(
                an_spPr().with_child(an_xfrm().with_child(an_ext().with_cx(width).with_cy(height)))
            )
        ).element

    @pytest.fixture
    def _PlaceholderFormat_(self, request, placeholder_format_):
        return class_mock(
            request,
            "pptx.shapes.base._PlaceholderFormat",
            return_value=placeholder_format_,
        )

    @pytest.fixture
    def placeholder_format_(self, request):
        return instance_mock(request, _PlaceholderFormat)

    @pytest.fixture
    def shadow_(self, request):
        return instance_mock(request, ShadowFormat)

    @pytest.fixture
    def ShadowFormat_(self, request):
        return class_mock(request, "pptx.shapes.base.ShadowFormat")

    @pytest.fixture
    def shape_elm_(self, request, shape_id, shape_name, txBody_):
        return instance_mock(
            request,
            BaseShapeElement,
            shape_id=shape_id,
            shape_name=shape_name,
            txBody=txBody_,
        )

    @pytest.fixture
    def shape_id(self):
        return 42

    @pytest.fixture
    def shape_name(self):
        return "Foobar 41"

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, SlideShapes)

    @pytest.fixture
    def sp(self):
        return an_sp().with_nsdecls().with_child(an_spPr()).element

    @pytest.fixture
    def sp_with_ext(self, width, height):
        return (
            an_sp()
            .with_nsdecls()
            .with_child(
                an_spPr().with_child(an_xfrm().with_child(an_ext().with_cx(width).with_cy(height)))
            )
        ).element

    @pytest.fixture
    def sp_with_off(self, left, top):
        return (
            an_sp()
            .with_nsdecls()
            .with_child(
                an_spPr().with_child(an_xfrm().with_child(an_off().with_x(left).with_y(top)))
            )
        ).element

    @pytest.fixture
    def top(self):
        return 456

    @pytest.fixture
    def txBody_(self, request):
        return instance_mock(request, CT_TextBody)

    @pytest.fixture
    def width(self):
        return 321


class DescribeBaseShape_animation(object):
    """Unit-test suite for `BaseShape.animation` / `.set_animation` (issue #102)."""

    def it_returns_None_when_the_slide_has_no_timing_element(self):
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        assert shape.animation is None

    def it_returns_None_when_slide_has_timing_but_no_effect_for_shape(self):
        from pptx.enum.animation import MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s1 = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        s2 = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(2), Inches(2), Inches(1), Inches(1)
        )
        s1.set_animation(MSO_ANIMATION_TYPE.FADE_IN)
        # -- s2 has no animation bound --
        assert s2.animation is None

    def it_returns_the_effect_proxy_when_bound(self):
        from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        shape.set_animation(
            MSO_ANIMATION_TYPE.FADE_IN, trigger="onClick", delay=250
        )
        effect = shape.animation
        assert effect is not None
        assert effect.type is MSO_ANIMATION_TYPE.FADE_IN
        assert effect.trigger is MSO_ANIMATION_TRIGGER.ON_CLICK
        assert effect.delay == 250
        assert effect.shape_id == shape.shape_id

    def it_accepts_onPrev_trigger_alias(self):
        from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        shape.set_animation(MSO_ANIMATION_TYPE.APPEAR, trigger="onPrev")
        assert shape.animation.trigger is MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS

    def it_removes_animation_when_effect_type_is_None(self):
        from pptx.enum.animation import MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        shape.set_animation(MSO_ANIMATION_TYPE.FADE_IN)
        assert shape.animation is not None

        shape.set_animation(None)
        assert shape.animation is None

    def it_replaces_any_prior_effect_on_the_same_shape(self):
        from pptx.enum.animation import MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        shape.set_animation(MSO_ANIMATION_TYPE.FADE_IN)
        shape.set_animation(MSO_ANIMATION_TYPE.PULSE)
        effect = shape.animation
        assert effect.type is MSO_ANIMATION_TYPE.PULSE
        # -- Only one p:spTgt for this shape --
        sld = slide._element
        matches = sld.xpath(
            ".//p:spTgt[@spid='%d']" % shape.shape_id
        )
        assert len(matches) == 1

    def it_raises_when_effect_type_is_not_an_enum_member(self):
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        with pytest.raises(TypeError, match="MSO_ANIMATION_TYPE"):
            shape.set_animation("fade_in")  # type: ignore[arg-type]

    def it_raises_for_unrecognized_trigger_string(self):
        from pptx.enum.animation import MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        with pytest.raises(ValueError, match="trigger"):
            shape.set_animation(MSO_ANIMATION_TYPE.FADE_IN, trigger="onBogus")

    def it_raises_for_negative_delay(self):
        from pptx.enum.animation import MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        with pytest.raises(ValueError, match="non-negative"):
            shape.set_animation(MSO_ANIMATION_TYPE.FADE_IN, delay=-5)

    def it_round_trips_through_save_and_reopen(self):
        import io

        from pptx.enum.animation import MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        shape.set_animation(MSO_ANIMATION_TYPE.FLY_IN, delay=300)
        orig_spid = shape.shape_id

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]
        shape2 = next(s for s in slide2.shapes if s.shape_id == orig_spid)
        effect = shape2.animation
        assert effect is not None
        assert effect.type is MSO_ANIMATION_TYPE.FLY_IN
        assert effect.delay == 300

    def it_enables_slide_has_animations_after_set(self):
        from pptx.enum.animation import MSO_ANIMATION_TYPE
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        assert slide.has_animations is False
        shape.set_animation(MSO_ANIMATION_TYPE.APPEAR)
        assert slide.has_animations is True


class DescribeSubshape(object):
    def it_knows_the_part_it_belongs_to(self, subshape_with_parent_):
        subshape, parent_ = subshape_with_parent_
        part = subshape.part
        assert part is parent_.part

    # fixtures ---------------------------------------------

    @pytest.fixture
    def subshape_with_parent_(self, request):
        parent_ = loose_mock(request, name="parent_")
        subshape = Subshape(parent_)
        return subshape, parent_


class Describe_CustomPropsDict(object):
    """Unit-test suite for `pptx.shapes.base._CustomPropsDict` (issue #582)."""

    def it_reports_zero_keys_for_a_bare_shape(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        props = shape.custom_props
        assert list(props) == []
        assert len(props) == 0
        assert "foo" not in props

    def it_can_set_and_get_a_key(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        shape.custom_props["department"] = "marketing"
        assert shape.custom_props["department"] == "marketing"
        assert shape.custom_props.get("department") == "marketing"

    def it_preserves_insertion_order_across_multiple_keys(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        props = shape.custom_props
        props["c"] = "1"
        props["a"] = "2"
        props["b"] = "3"
        assert list(props) == ["c", "a", "b"]
        assert props.items() == [("c", "1"), ("a", "2"), ("b", "3")]
        assert props.values() == ["1", "2", "3"]

    def it_overwrites_an_existing_value_in_place(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        shape.custom_props["k"] = "first"
        shape.custom_props["k2"] = "second"
        shape.custom_props["k"] = "updated"
        # -- overwrite is in-place, so order does not change --
        assert list(shape.custom_props) == ["k", "k2"]
        assert shape.custom_props["k"] == "updated"

    def it_returns_default_on_get_for_missing_key(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        assert shape.custom_props.get("missing") is None
        assert shape.custom_props.get("missing", "def") == "def"

    def it_raises_KeyError_on_subscript_of_missing_key(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        with pytest.raises(KeyError):
            shape.custom_props["missing"]

    def it_raises_KeyError_on_del_of_missing_key(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        with pytest.raises(KeyError):
            del shape.custom_props["missing"]

    def it_deletes_a_key_and_leaves_other_keys_intact(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        props = shape.custom_props
        props["a"] = "1"
        props["b"] = "2"
        props["c"] = "3"
        del props["b"]
        assert list(props) == ["a", "c"]
        assert "b" not in props

    def it_removes_the_extLst_when_the_last_prop_is_deleted(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        shape.custom_props["only"] = "1"
        del shape.custom_props["only"]
        cNvPr = shape._element._nvXxPr.cNvPr
        assert cNvPr.find(qn("a:extLst")) is None

    def it_removes_the_extLst_on_clear(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        shape.custom_props["a"] = "1"
        shape.custom_props["b"] = "2"
        shape.custom_props.clear()
        assert list(shape.custom_props) == []
        cNvPr = shape._element._nvXxPr.cNvPr
        assert cNvPr.find(qn("a:extLst")) is None

    def it_tolerates_clear_when_already_empty(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        shape.custom_props.clear()
        assert list(shape.custom_props) == []

    def it_rejects_non_string_keys(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        with pytest.raises(TypeError):
            shape.custom_props[42] = "x"  # type: ignore[index]

    def it_rejects_non_string_values(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        with pytest.raises(TypeError):
            shape.custom_props["k"] = 7  # type: ignore[assignment]

    def it_equates_to_a_dict_with_the_same_items(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        shape.custom_props["a"] = "1"
        shape.custom_props["b"] = "2"
        assert shape.custom_props == {"a": "1", "b": "2"}
        assert shape.custom_props != {"a": "1"}

    def it_handles_empty_string_values(self):
        shape = self._shape("p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}")
        shape.custom_props["blank"] = ""
        assert shape.custom_props["blank"] == ""
        assert list(shape.custom_props) == ["blank"]

    def it_is_available_on_every_slide_shape_type(self):
        """Smoke: every shape subtype exposes a working ``custom_props`` mapping."""
        for cxml in (
            "p:sp/p:nvSpPr/p:cNvPr{id=1,name=A}",
            "p:pic/p:nvPicPr/p:cNvPr{id=1,name=A}",
            "p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=1,name=A}",
            "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=1,name=A}",
            "p:grpSp/p:nvGrpSpPr/p:cNvPr{id=1,name=A}",
        ):
            shape = self._shape(cxml)
            shape.custom_props["k"] = "v"
            assert shape.custom_props["k"] == "v"

    def it_survives_a_full_save_reload_round_trip(self, tmp_path):
        """End-to-end: values set before save are visible after reopen (issue #582)."""
        from io import BytesIO

        from pptx.util import Inches

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        shape.custom_props["department"] = "marketing"
        shape.custom_props["owner"] = "alice"

        buf = BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        reopened = prs2.slides[0].shapes[0]
        assert reopened.custom_props == {"department": "marketing", "owner": "alice"}
        assert list(reopened.custom_props) == ["department", "owner"]

    # -- fixtures ------------------------------------------------------

    def _shape(self, cxml: str) -> BaseShape:
        return BaseShape(cast("ShapeElement", element(cxml)), None)


class Describe_PlaceholderFormat(object):
    def it_knows_its_idx(self, idx_get_fixture):
        placeholder_format, expected_value = idx_get_fixture
        assert placeholder_format.idx == expected_value

    def it_knows_its_type(self, type_get_fixture):
        placeholder_format, expected_value = type_get_fixture
        assert placeholder_format.type == expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[("p:ph", 0), ("p:ph{idx=42}", 42)])
    def idx_get_fixture(self, request):
        ph_cxml, expected_value = request.param
        placeholder_format = _PlaceholderFormat(element(ph_cxml))
        return placeholder_format, expected_value

    @pytest.fixture(
        params=[
            ("p:ph", PP_PLACEHOLDER.OBJECT),
            ("p:ph{type=pic}", PP_PLACEHOLDER.PICTURE),
        ]
    )
    def type_get_fixture(self, request):
        ph_cxml, expected_value = request.param
        placeholder_format = _PlaceholderFormat(element(ph_cxml))
        return placeholder_format, expected_value
