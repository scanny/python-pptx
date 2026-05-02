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


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from other tests that mutate `PartFactory.part_type_for`.

    In particular, `tests/opc/test_package.py::DescribePartFactory` overwrites
    `PartFactory.part_type_for[CT.PML_SLIDE]` with a Mock and does not restore it.
    Tests that load a real .pptx file (which walks the relationship graph and
    therefore the part-type registry) need to restore the default mapping.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.image import ImagePart
    from pptx.parts.media import MediaPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import (
        NotesMasterPart,
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    saved = dict(PartFactory.part_type_for)
    expected = {
        CT.PML_PRESENTATION_MAIN: PresentationPart,
        CT.PML_PRES_MACRO_MAIN: PresentationPart,
        CT.PML_TEMPLATE_MAIN: PresentationPart,
        CT.PML_SLIDESHOW_MAIN: PresentationPart,
        CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


class DescribeBaseShape(object):
    """Unit-test suite for `pptx.shapes.base.BaseShape` objects."""

    def it_provides_access_to_its_click_action(self, click_action_fixture):
        shape, ActionSetting_, cNvPr, click_action_ = click_action_fixture
        click_action = shape.click_action
        ActionSetting_.assert_called_once_with(cNvPr, shape)
        assert click_action is click_action_

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

    def it_knows_its_rotation_angle(self, rotation_get_fixture):
        shape, expected_value = rotation_get_fixture
        assert shape.rotation == expected_value

    def it_can_change_its_rotation_angle(self, rotation_set_fixture):
        shape, new_value, expected_xml = rotation_set_fixture
        shape.rotation = new_value
        assert shape._element.xml == expected_xml

    def it_provides_access_to_its_shadow(self, shadow_fixture):
        shape, ShadowFormat_, spPr, shadow_ = shadow_fixture

        shadow = shape.shadow

        ShadowFormat_.assert_called_once_with(spPr)
        assert shadow is shadow_

    def it_knows_the_part_it_belongs_to(self, part_fixture):
        shape, parent_ = part_fixture
        part = shape.part
        assert part is parent_.part

    def it_knows_it_doesnt_have_a_text_frame(self):
        shape = BaseShape(None, None)
        assert shape.has_text_frame is False

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

    def it_raises_when_duplicating_a_group_shape(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        gs = slide.shapes.add_group_shape()
        with pytest.raises(NotImplementedError, match="simple shapes"):
            gs.duplicate()

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
        self, _restore_part_factory
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
