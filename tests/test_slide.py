# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.slide` module."""

from __future__ import annotations

import pytest

from pptx.dml.fill import FillFormat
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.package import Package
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.shapes.base import BaseShape
from pptx.shapes.placeholder import LayoutPlaceholder, NotesSlidePlaceholder
from pptx.shapes.shapetree import (
    LayoutPlaceholders,
    LayoutShapes,
    MasterPlaceholders,
    MasterShapes,
    NotesSlidePlaceholders,
    NotesSlideShapes,
    SlidePlaceholders,
    SlideShapes,
)
from pptx.enum.transition import (
    PP_TRANSITION_SIDE_DIRECTION,
    PP_TRANSITION_SPEED,
    PP_TRANSITION_TYPE,
)
from pptx.oxml.ns import qn
from pptx.slide import (
    NotesMaster,
    NotesSlide,
    Slide,
    SlideLayout,
    SlideLayouts,
    SlideMaster,
    SlideMasters,
    Slides,
    Transition,
    _Background,
    _BaseMaster,
    _BaseSlide,
    _HeaderFooter,
)
from pptx.text.text import TextFrame

from .unitutil.cxml import element, xml
from .unitutil.mock import (
    call,
    class_mock,
    function_mock,
    instance_mock,
    loose_mock,
    method_mock,
    property_mock,
)


class Describe_BaseSlide(object):
    """Unit-test suite for `pptx.slide._BaseSlide` objects."""

    def it_knows_its_name(self, name_get_fixture):
        base_slide, expected_value = name_get_fixture
        assert base_slide.name == expected_value

    def it_can_change_its_name(self, name_set_fixture):
        base_slide, new_value, expected_xml = name_set_fixture
        base_slide.name = new_value
        assert base_slide._element.xml == expected_xml

    def it_provides_access_to_its_background(self, background_fixture):
        slide, _Background_, cSld, background_ = background_fixture

        background = slide.background

        _Background_.assert_called_once_with(cSld)
        assert background is background_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def background_fixture(self, _Background_, background_):
        sld = element("p:sld/p:cSld")
        slide = _BaseSlide(sld, None)
        cSld = sld.xpath("//p:cSld")[0]
        _Background_.return_value = background_
        return slide, _Background_, cSld, background_

    @pytest.fixture(params=[("p:sld/p:cSld", ""), ("p:sld/p:cSld{name=Foobar}", "Foobar")])
    def name_get_fixture(self, request):
        sld_cxml, expected_name = request.param
        base_slide = _BaseSlide(element(sld_cxml), None)
        return base_slide, expected_name

    @pytest.fixture(
        params=[
            ("p:sld/p:cSld", "foo", "p:sld/p:cSld{name=foo}"),
            ("p:sld/p:cSld{name=foo}", "bar", "p:sld/p:cSld{name=bar}"),
            ("p:sld/p:cSld{name=bar}", "", "p:sld/p:cSld"),
            ("p:sld/p:cSld{name=bar}", None, "p:sld/p:cSld"),
            ("p:sld/p:cSld", "", "p:sld/p:cSld"),
            ("p:sld/p:cSld", None, "p:sld/p:cSld"),
        ]
    )
    def name_set_fixture(self, request):
        xSld_cxml, new_value, expected_cxml = request.param
        base_slide = _BaseSlide(element(xSld_cxml), None)
        expected_xml = xml(expected_cxml)
        return base_slide, new_value, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _Background_(self, request):
        return class_mock(request, "pptx.slide._Background")

    @pytest.fixture
    def background_(self, request):
        return instance_mock(request, _Background)


class Describe_BaseMaster(object):
    """Unit-test suite for `pptx.slide._BaseMaster` objects."""

    def it_is_a_BaseSlide_subclass(self, subclass_fixture):
        base_master = subclass_fixture
        assert isinstance(base_master, _BaseSlide)

    def it_provides_access_to_its_placeholders(self, placeholders_fixture):
        base_master, MasterPlaceholders_, spTree, placeholders_ = placeholders_fixture
        placeholders = base_master.placeholders
        MasterPlaceholders_.assert_called_once_with(spTree, base_master)
        assert placeholders is placeholders_

    def it_provides_access_to_its_shapes(self, shapes_fixture):
        base_master, MasterShapes_, spTree, shapes_ = shapes_fixture
        shapes = base_master.shapes
        MasterShapes_.assert_called_once_with(spTree, base_master)
        assert shapes is shapes_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def placeholders_fixture(self, MasterPlaceholders_, placeholders_):
        master_element = element("p:sldMaster/p:cSld/p:spTree")
        base_master = _BaseMaster(master_element, None)
        spTree = master_element.xpath("//p:spTree")[0]
        return base_master, MasterPlaceholders_, spTree, placeholders_

    @pytest.fixture
    def shapes_fixture(self, MasterShapes_, shapes_):
        master_element = element("p:sldMaster/p:cSld/p:spTree")
        base_master = _BaseMaster(master_element, None)
        spTree = master_element.xpath("//p:spTree")[0]
        return base_master, MasterShapes_, spTree, shapes_

    @pytest.fixture
    def subclass_fixture(self):
        return _BaseMaster(None, None)

    # fixture components -----------------------------------

    @pytest.fixture
    def MasterPlaceholders_(self, request, placeholders_):
        return class_mock(request, "pptx.slide.MasterPlaceholders", return_value=placeholders_)

    @pytest.fixture
    def MasterShapes_(self, request, shapes_):
        return class_mock(request, "pptx.slide.MasterShapes", return_value=shapes_)

    @pytest.fixture
    def placeholders_(self, request):
        return instance_mock(request, MasterPlaceholders)

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, MasterShapes)


class DescribeNotesSlide(object):
    """Unit-test suite for `pptx.slide.NotesSlide` objects."""

    def it_can_clone_the_notes_master_placeholders(self, request, notes_master_, shapes_):
        placeholders = notes_master_.placeholders = (
            BaseShape(element("p:sp/p:nvSpPr/p:nvPr/p:ph{type=body}"), None),
            BaseShape(element("p:sp/p:nvSpPr/p:nvPr/p:ph{type=dt}"), None),
        )
        property_mock(request, NotesSlide, "shapes", return_value=shapes_)
        notes_slide = NotesSlide(None, None)

        notes_slide.clone_master_placeholders(notes_master_)

        assert shapes_.clone_placeholder.call_args_list == [call(placeholders[0])]

    def it_provides_access_to_its_shapes(self, shapes_fixture):
        notes_slide, NotesSlideShapes_, spTree, shapes_ = shapes_fixture
        shapes = notes_slide.shapes
        NotesSlideShapes_.assert_called_once_with(spTree, notes_slide)
        assert shapes is shapes_

    def it_provides_access_to_its_placeholders(self, placeholders_fixture):
        (
            notes_slide,
            NotesSlidePlaceholders_,
            spTree,
            placeholders_,
        ) = placeholders_fixture
        placeholders = notes_slide.placeholders
        NotesSlidePlaceholders_.assert_called_once_with(spTree, notes_slide)
        assert placeholders is placeholders_

    def it_provides_access_to_its_notes_placeholder(self, notes_ph_fixture):
        notes_slide, expected_value = notes_ph_fixture
        placeholder = notes_slide.notes_placeholder
        assert placeholder is expected_value

    def it_provides_access_to_its_notes_text_frame(self, notes_tf_fixture):
        notes_slide, expected_value = notes_tf_fixture
        text_frame = notes_slide.notes_text_frame
        assert text_frame is expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (("SLIDE_IMAGE", "BODY", "SLIDE_NUMBER"), 1),
            (("DATE", "SLIDE_IMAGE", "FOOTER"), None),
        ]
    )
    def notes_ph_fixture(self, request, placeholders_prop_):
        type_names, match = request.param
        notes_slide = NotesSlide(None, None)
        placeholders_ = []
        for type_name in type_names:
            placeholder_ = instance_mock(
                request, NotesSlidePlaceholder, name="%s-placeholder" % type_name
            )
            placeholder_.placeholder_format.type = getattr(PP_PLACEHOLDER, type_name)
            placeholders_.append(placeholder_)
        placeholders_prop_.return_value = placeholders_
        expected_value = None if match is None else placeholders_[match]
        return notes_slide, expected_value

    @pytest.fixture(params=[True, False])
    def notes_tf_fixture(self, request, notes_placeholder_prop_, placeholder_, text_frame_):
        has_text_frame = request.param
        notes_slide = NotesSlide(None, None)
        if has_text_frame:
            notes_placeholder_prop_.return_value = placeholder_
            placeholder_.text_frame = text_frame_
            expected_value = text_frame_
        else:
            notes_placeholder_prop_.return_value = None
            expected_value = None
        return notes_slide, expected_value

    @pytest.fixture
    def placeholders_fixture(self, NotesSlidePlaceholders_, placeholders_):
        notes = element("p:notes/p:cSld/p:spTree")
        notes_slide = NotesSlide(notes, None)
        spTree = notes.xpath("//p:spTree")[0]
        return notes_slide, NotesSlidePlaceholders_, spTree, placeholders_

    @pytest.fixture
    def shapes_fixture(self, NotesSlideShapes_, shapes_):
        notes = element("p:notes/p:cSld/p:spTree")
        notes_slide = NotesSlide(notes, None)
        spTree = notes.xpath("//p:spTree")[0]
        return notes_slide, NotesSlideShapes_, spTree, shapes_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def notes_master_(self, request):
        return instance_mock(request, NotesMaster)

    @pytest.fixture
    def notes_placeholder_prop_(self, request, placeholder_):
        return property_mock(request, NotesSlide, "notes_placeholder", return_value=placeholder_)

    @pytest.fixture
    def NotesSlidePlaceholders_(self, request, placeholders_):
        return class_mock(request, "pptx.slide.NotesSlidePlaceholders", return_value=placeholders_)

    @pytest.fixture
    def NotesSlideShapes_(self, request, shapes_):
        return class_mock(request, "pptx.slide.NotesSlideShapes", return_value=shapes_)

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, NotesSlidePlaceholder)

    @pytest.fixture
    def placeholders_(self, request):
        return instance_mock(request, NotesSlidePlaceholders)

    @pytest.fixture
    def placeholders_prop_(self, request, placeholders_):
        return property_mock(request, NotesSlide, "placeholders", return_value=placeholders_)

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, NotesSlideShapes)

    @pytest.fixture
    def text_frame_(self, request):
        return instance_mock(request, TextFrame)


class DescribeSlide(object):
    """Unit-test suite for `pptx.slide.Slide` objects."""

    def it_is_a_BaseSlide_subclass(self, subclass_fixture):
        slide = subclass_fixture
        assert isinstance(slide, _BaseSlide)

    def it_provides_access_to_its_background(self, background_fixture):
        slide, _BaseSlide_background_, background_ = background_fixture

        background = slide.background

        _BaseSlide_background_.assert_called_once_with()
        assert background is background_

    def it_knows_whether_it_follows_the_mstr_bkgd(self, follow_get_fixture):
        slide, expected_value = follow_get_fixture
        follows = slide.follow_master_background
        assert follows is expected_value

    def it_knows_whether_it_has_a_notes_slide(self, has_notes_slide_fixture):
        slide, expected_value = has_notes_slide_fixture
        assert slide.has_notes_slide is expected_value

    def it_knows_its_slide_id(self, slide_id_fixture):
        slide, expected_value = slide_id_fixture
        assert slide.slide_id == expected_value

    def it_provides_access_to_its_shapes(self, shapes_fixture):
        slide, SlideShapes_, spTree, shapes_ = shapes_fixture
        shapes = slide.shapes
        SlideShapes_.assert_called_once_with(spTree, slide)
        assert shapes is shapes_

    def it_provides_access_to_its_placeholders(self, placeholders_fixture):
        slide, SlidePlaceholders_, spTree, placeholders_ = placeholders_fixture
        placeholders = slide.placeholders
        SlidePlaceholders_.assert_called_once_with(spTree, slide)
        assert placeholders is placeholders_

    def it_provides_access_to_its_slide_layout(self, layout_fixture):
        slide, slide_layout_ = layout_fixture
        assert slide.slide_layout is slide_layout_

    def it_provides_access_to_its_notes_slide(self, notes_slide_fixture):
        slide, notes_slide_ = notes_slide_fixture
        assert slide.notes_slide is notes_slide_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def background_fixture(self, _BaseSlide_background_, background_):
        slide = Slide(None, None)
        _BaseSlide_background_.return_value = background_
        return slide, _BaseSlide_background_, background_

    @pytest.fixture(params=[("p:sld/p:cSld", True), ("p:sld/p:cSld/p:bg", False)])
    def follow_get_fixture(self, request):
        pSld_cxml, expected_value = request.param
        slide = Slide(element(pSld_cxml), None)
        return slide, expected_value

    @pytest.fixture
    def has_notes_slide_fixture(self, part_prop_, slide_part_):
        slide = Slide(None, None)
        expected_value = slide_part_.has_notes_slide = 42
        return slide, expected_value

    @pytest.fixture
    def layout_fixture(self, slide_layout_, part_prop_, slide_part_):
        slide = Slide(None, None)
        part_prop_.return_value = slide_part_
        slide_part_.slide_layout = slide_layout_
        return slide, slide_layout_

    @pytest.fixture
    def notes_slide_fixture(self, notes_slide_, part_prop_, slide_part_):
        slide = Slide(None, None)
        slide_part_.notes_slide = notes_slide_
        return slide, notes_slide_

    @pytest.fixture
    def placeholders_fixture(self, SlidePlaceholders_, placeholders_):
        sld = element("p:sld/p:cSld/p:spTree")
        slide = Slide(sld, None)
        spTree = sld.xpath("//p:spTree")[0]
        return slide, SlidePlaceholders_, spTree, placeholders_

    @pytest.fixture
    def shapes_fixture(self, SlideShapes_, shapes_):
        sld = element("p:sld/p:cSld/p:spTree")
        spTree = sld.xpath("//p:spTree")[0]
        slide = Slide(sld, None)
        return slide, SlideShapes_, spTree, shapes_

    @pytest.fixture
    def slide_id_fixture(self, part_prop_, slide_part_):
        slide = Slide(None, None)
        slide_id = 256
        slide_part_.slide_id = slide_id
        return slide, slide_id

    @pytest.fixture
    def subclass_fixture(self):
        return Slide(None, None)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _BaseSlide_background_(self, request):
        return property_mock(request, _BaseSlide, "background")

    @pytest.fixture
    def background_(self, request):
        return instance_mock(request, _Background)

    @pytest.fixture
    def notes_slide_(self, request):
        return instance_mock(request, NotesSlide)

    @pytest.fixture
    def part_prop_(self, request, slide_part_):
        return property_mock(request, Slide, "part", return_value=slide_part_)

    @pytest.fixture
    def placeholders_(self, request):
        return instance_mock(request, SlidePlaceholders)

    @pytest.fixture
    def SlidePlaceholders_(self, request, placeholders_):
        return class_mock(request, "pptx.slide.SlidePlaceholders", return_value=placeholders_)

    @pytest.fixture
    def SlideShapes_(self, request, shapes_):
        return class_mock(request, "pptx.slide.SlideShapes", return_value=shapes_)

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, SlideShapes)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)


class DescribeSlides(object):
    """Unit-test suite for `pptx.slide.Slides` objects."""

    def it_supports_indexed_access(self, getitem_fixture):
        slides, prs_part_, rId, slide_ = getitem_fixture
        slide = slides[0]
        prs_part_.related_slide.assert_called_once_with(rId)
        assert slide is slide_

    def it_raises_on_slide_index_out_of_range(self, getitem_raises_fixture):
        slides = getitem_raises_fixture
        with pytest.raises(IndexError):
            slides[2]

    def it_knows_the_index_of_a_slide_it_contains(self, index_fixture):
        slides, slide, expected_value = index_fixture
        index = slides.index(slide)
        assert index == expected_value

    def it_raises_on_slide_not_in_collection(self, raises_fixture):
        slides, slide = raises_fixture
        with pytest.raises(ValueError):
            slides.index(slide)

    def it_can_iterate_its_slides(self, iter_fixture):
        slides, related_slide_, calls, expected_value = iter_fixture
        slide_lst = [s for s in slides]
        assert related_slide_.call_args_list == calls
        assert slide_lst == expected_value

    def it_supports_len(self, len_fixture):
        slides, expected_value = len_fixture
        assert len(slides) == expected_value

    def it_can_add_a_new_slide(self, add_fixture):
        slides, slide_layout_, part_ = add_fixture[:3]
        clone_layout_placeholders_, expected_xml, slide_ = add_fixture[3:]

        slide = slides.add_slide(slide_layout_)

        part_.add_slide.assert_called_once_with(slide_layout_)
        clone_layout_placeholders_.assert_called_once_with(slide_layout_)
        assert slides._sldIdLst.xml == expected_xml
        assert slide is slide_

    def it_can_add_a_slide_cloned_from_another_presentation(
        self, part_prop_, slide_, slide_layout_
    ):
        slides = Slides(element("p:sldIdLst/p:sldId{r:id=rId1}"), None)
        part_ = part_prop_.return_value
        source_slide_ = slide_  # reuse mock for brevity
        part_.add_slide_from_external.return_value = ("rId2", slide_)
        expected_xml = xml("p:sldIdLst/(p:sldId{r:id=rId1},p:sldId{r:id=rId2,id=256})")

        new_slide = slides.add_slide_from_external(source_slide_, slide_layout_)

        part_.add_slide_from_external.assert_called_once_with(source_slide_, slide_layout_)
        assert slides._sldIdLst.xml == expected_xml
        assert new_slide is slide_

    def it_can_duplicate_a_slide(self, part_prop_, slide_):
        """Slides.duplicate() appends a new `p:sldId` referencing the cloned slide."""
        sldIdLst = element("p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257})")
        slides = Slides(sldIdLst, None)
        _slides = [Slide(element("p:sld"), None), Slide(element("p:sld"), None)]
        # -- .index() iterates, driving related_slide() once per element until match --
        part_prop_.return_value.related_slide.side_effect = _slides
        new_slide_ = slide_
        part_prop_.return_value.duplicate_slide.return_value = ("rIdNew", new_slide_)

        new_slide = slides.duplicate(_slides[0])

        part_prop_.return_value.duplicate_slide.assert_called_once_with(_slides[0])
        assert new_slide is new_slide_
        # -- the new p:sldId was appended to the end of p:sldIdLst --
        assert sldIdLst.xml == xml(
            "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},"
            "p:sldId{r:id=rIdNew,id=258})"
        )

    @pytest.mark.parametrize(
        ("index", "expected_cxml"),
        [
            # --- insert at beginning ---
            (
                0,
                "p:sldIdLst/(p:sldId{r:id=rIdNew,id=258},p:sldId{r:id=a,id=256},"
                "p:sldId{r:id=b,id=257})",
            ),
            # --- insert in the middle ---
            (
                1,
                "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=rIdNew,id=258},"
                "p:sldId{r:id=b,id=257})",
            ),
            # --- insert at end via explicit positive index ---
            (
                2,
                "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},"
                "p:sldId{r:id=rIdNew,id=258})",
            ),
            # --- negative index counts from the end ---
            (
                -1,
                "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},"
                "p:sldId{r:id=rIdNew,id=258})",
            ),
            # --- index beyond end clamps to last position ---
            (
                99,
                "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},"
                "p:sldId{r:id=rIdNew,id=258})",
            ),
            # --- sufficiently-negative index clamps to first position ---
            (
                -99,
                "p:sldIdLst/(p:sldId{r:id=rIdNew,id=258},p:sldId{r:id=a,id=256},"
                "p:sldId{r:id=b,id=257})",
            ),
        ],
    )
    def it_can_duplicate_a_slide_at_a_given_index(
        self, index, expected_cxml, part_prop_, slide_
    ):
        sldIdLst = element("p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257})")
        slides = Slides(sldIdLst, None)
        _slides = [Slide(element("p:sld"), None), Slide(element("p:sld"), None)]
        part_prop_.return_value.related_slide.side_effect = _slides
        part_prop_.return_value.duplicate_slide.return_value = ("rIdNew", slide_)

        slides.duplicate(_slides[0], index=index)

        assert sldIdLst.xml == xml(expected_cxml)

    def it_raises_on_duplicate_when_slide_not_in_collection(self, part_prop_):
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=a,id=256}")
        slides = Slides(sldIdLst, None)
        part_prop_.return_value.related_slide.return_value = Slide(element("p:sld"), None)
        stranger = Slide(element("p:sld"), None)

        with pytest.raises(ValueError):
            slides.duplicate(stranger)

        part_prop_.return_value.duplicate_slide.assert_not_called()

    def it_finds_a_slide_by_slide_id(self, get_fixture):
        slides, slide_id, default, prs_part_, expected_value = get_fixture
        slide = slides.get(slide_id, default)
        prs_part_.get_slide.assert_called_once_with(slide_id)
        assert slide is expected_value

    @pytest.mark.parametrize(
        ("src_idx", "new_idx", "expected_cxml"),
        [
            # --- move first slide to the middle ---
            (
                0,
                1,
                "p:sldIdLst/(p:sldId{r:id=b,id=256},p:sldId{r:id=a,id=257},"
                "p:sldId{r:id=c,id=258})",
            ),
            # --- move last slide to the first position ---
            (
                2,
                0,
                "p:sldIdLst/(p:sldId{r:id=c,id=258},p:sldId{r:id=a,id=257},"
                "p:sldId{r:id=b,id=256})",
            ),
            # --- move first slide to the end ---
            (
                0,
                2,
                "p:sldIdLst/(p:sldId{r:id=b,id=256},p:sldId{r:id=c,id=258},"
                "p:sldId{r:id=a,id=257})",
            ),
            # --- no-op when target index equals current index ---
            (
                1,
                1,
                "p:sldIdLst/(p:sldId{r:id=a,id=257},p:sldId{r:id=b,id=256},"
                "p:sldId{r:id=c,id=258})",
            ),
            # --- negative index counts from the end ---
            (
                0,
                -1,
                "p:sldIdLst/(p:sldId{r:id=b,id=256},p:sldId{r:id=c,id=258},"
                "p:sldId{r:id=a,id=257})",
            ),
            # --- index beyond end is clamped to last position ---
            (
                0,
                99,
                "p:sldIdLst/(p:sldId{r:id=b,id=256},p:sldId{r:id=c,id=258},"
                "p:sldId{r:id=a,id=257})",
            ),
            # --- sufficiently-negative index clamps to position 0 ---
            (
                2,
                -99,
                "p:sldIdLst/(p:sldId{r:id=c,id=258},p:sldId{r:id=a,id=257},"
                "p:sldId{r:id=b,id=256})",
            ),
        ],
    )
    def it_can_move_a_slide_to_a_new_position(self, src_idx, new_idx, expected_cxml, part_prop_):
        sldIdLst = element(
            "p:sldIdLst/(p:sldId{r:id=a,id=257},p:sldId{r:id=b,id=256}," "p:sldId{r:id=c,id=258})"
        )
        slides = Slides(sldIdLst, None)
        _slides = [Slide(element("p:sld"), None) for _ in range(3)]
        # -- .index() iterates, which drives related_slide() once per element --
        part_prop_.return_value.related_slide.side_effect = _slides

        slides.move_slide(_slides[src_idx], new_idx)

        assert sldIdLst.xml == xml(expected_cxml)

    def it_raises_on_move_slide_when_slide_not_in_collection(self, part_prop_):
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=a,id=256}")
        slides = Slides(sldIdLst, None)
        part_prop_.return_value.related_slide.return_value = Slide(element("p:sld"), None)
        stranger = Slide(element("p:sld"), None)
        with pytest.raises(ValueError):
            slides.move_slide(stranger, 0)

    @pytest.mark.parametrize(
        ("target_idx", "expected_rId", "expected_cxml"),
        [
            # --- delete first slide ---
            (
                0,
                "a",
                "p:sldIdLst/(p:sldId{r:id=b,id=257},p:sldId{r:id=c,id=258})",
            ),
            # --- delete middle slide ---
            (
                1,
                "b",
                "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=c,id=258})",
            ),
            # --- delete last slide ---
            (
                2,
                "c",
                "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257})",
            ),
        ],
    )
    def it_can_delete_a_slide(self, target_idx, expected_rId, expected_cxml, part_prop_):
        sldIdLst = element(
            "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},"
            "p:sldId{r:id=c,id=258})"
        )
        slides = Slides(sldIdLst, None)
        _slides = [Slide(element("p:sld"), None) for _ in range(3)]
        # -- .index() iterates, driving related_slide() once per element until match --
        part_prop_.return_value.related_slide.side_effect = _slides

        slides.delete(_slides[target_idx])

        assert sldIdLst.xml == xml(expected_cxml)
        part_prop_.return_value.drop_rel.assert_called_once_with(expected_rId)

    def it_can_delete_the_only_slide(self, part_prop_):
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=a,id=256}")
        slides = Slides(sldIdLst, None)
        only_slide = Slide(element("p:sld"), None)
        part_prop_.return_value.related_slide.return_value = only_slide

        slides.delete(only_slide)

        assert len(sldIdLst.sldId_lst) == 0
        part_prop_.return_value.drop_rel.assert_called_once_with("a")

    def it_raises_on_delete_when_slide_not_in_collection(self, part_prop_):
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=a,id=256}")
        slides = Slides(sldIdLst, None)
        part_prop_.return_value.related_slide.return_value = Slide(element("p:sld"), None)
        stranger = Slide(element("p:sld"), None)

        with pytest.raises(ValueError):
            slides.delete(stranger)

        part_prop_.return_value.drop_rel.assert_not_called()

    def it_allocates_a_unique_sldId_after_delete_move_and_insert(self, part_prop_):
        """Regression for Foundation F6.

        After a sequence of move_slide, delete, duplicate (with index), and
        add_slide_from_external, the `id` attribute on every `p:sldId` remains
        unique. In particular, the slide-id of a just-deleted slide is never
        reused while a higher-valued id is still in play: ``CT_SlideIdList._next_id``
        always returns ``max(used) + 1`` for the common case, so the only way to
        observe a collision would be a helper that assigns ids on its own --
        this test pins that invariant on the F6 API surface.
        """
        # --- starting layout: three slides at 256, 257, 258 ---
        sldIdLst = element(
            "p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},p:sldId{r:id=c,id=258})"
        )
        slides = Slides(sldIdLst, None)
        _slides = [Slide(element("p:sld"), None) for _ in range(3)]
        part_ = part_prop_.return_value
        # -- related_slide is called once per iteration of .index() / __iter__; the
        # -- scenario invokes index() six times on up-to-three-element collections, so
        # -- we return the canonical slide for each rId encountered.
        part_.related_slide.side_effect = lambda rId: {
            "a": _slides[0],
            "b": _slides[1],
            "c": _slides[2],
            "rIdDup": _slides[0],  # duplicate of slide a
            "rIdExt": _slides[1],  # external slide bound to layout
        }[rId]

        # --- 1. move slide c to position 0 -> (c,a,b) with ids (258,256,257) ---
        slides.move_slide(_slides[2], 0)

        # --- 2. delete slide a -> (c,b) with ids (258,257) ---
        slides.delete(_slides[0])

        # --- 3. duplicate slide b at the beginning. _next_id must skip the re-usable
        # --- but-deleted 256 and pick 259, because 258 is still in play. ---
        part_.duplicate_slide.return_value = ("rIdDup", _slides[0])
        slides.duplicate(_slides[1], index=0)

        # --- 4. add a slide cloned from another presentation; it appends and gets 260. ---
        part_.add_slide_from_external.return_value = ("rIdExt", _slides[1])
        slides.add_slide_from_external(_slides[1], None)  # type: ignore[arg-type]

        # --- assertion: every id is unique and monotonically above min-slide-id ---
        ids = [s.id for s in sldIdLst.sldId_lst]
        assert len(ids) == len(set(ids)), f"duplicate sldId values: {ids}"
        assert all(i >= 256 for i in ids)
        # --- the deleted slide's id (256) must not have been recycled while 258 is live ---
        assert 256 not in ids

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self, slide_layout_, part_prop_, slide_):
        slides = Slides(element("p:sldIdLst/p:sldId{r:id=rId1}"), None)
        part_ = part_prop_.return_value
        clone_layout_placeholders_ = slide_.shapes.clone_layout_placeholders
        expected_xml = xml("p:sldIdLst/(p:sldId{r:id=rId1},p:sldId{r:id=rId2,id=256})")
        part_.add_slide.return_value = "rId2", slide_
        return (
            slides,
            slide_layout_,
            part_,
            clone_layout_placeholders_,
            expected_xml,
            slide_,
        )

    @pytest.fixture(params=[True, False])
    def get_fixture(self, request, part_prop_, prs_part_, slide_):
        found = request.param
        slides = Slides(None, None)
        slide_id, default = 256, "foobar"
        expected_value = slide_ if found else default
        prs_part_.get_slide.return_value = slide_ if found else None
        return slides, slide_id, default, prs_part_, expected_value

    @pytest.fixture
    def getitem_fixture(self, prs_part_, slide_, part_prop_):
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=rId1}")
        slides = Slides(sldIdLst, None)
        prs_part_.related_slide.return_value = slide_
        return slides, prs_part_, "rId1", slide_

    @pytest.fixture
    def getitem_raises_fixture(self):
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=rId1}")
        slides = Slides(sldIdLst, None)
        return slides

    @pytest.fixture(params=[0, 1])
    def index_fixture(self, request, part_prop_):
        idx = request.param
        sldIdLst = element("p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})")
        slides = Slides(sldIdLst, None)
        _slides = [Slide(element("p:sld"), None), Slide(element("p:sld"), None)]
        part_prop_.return_value.related_slide.side_effect = _slides
        return slides, _slides[idx], idx

    @pytest.fixture
    def iter_fixture(self, part_prop_, slide_):
        sldIdLst = element("p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})")
        slides = Slides(sldIdLst, None)
        related_slide_ = part_prop_.return_value.related_slide
        related_slide_.return_value = slide_
        calls = [call("a"), call("b")]
        _slides = [slide_, slide_]
        return slides, related_slide_, calls, _slides

    @pytest.fixture(
        params=[
            ("p:sldIdLst", 0),
            ("p:sldIdLst/p:sldId{r:id=a}", 1),
            ("p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})", 2),
        ]
    )
    def len_fixture(self, request):
        sldIdLst_cxml, expected_value = request.param
        slides = Slides(element(sldIdLst_cxml), None)
        return slides, expected_value

    @pytest.fixture
    def raises_fixture(self):
        slides = Slides(element("p:sldIdLst"), None)
        slide = Slide(element("p:sld"), None)
        return slides, slide

    # fixture components ---------------------------------------------

    @pytest.fixture
    def part_prop_(self, request, prs_part_):
        return property_mock(request, Slides, "part", return_value=prs_part_)

    @pytest.fixture
    def prs_part_(self, request):
        return instance_mock(request, PresentationPart)

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)


class DescribeSlideLayout(object):
    """Unit-test suite for `pptx.slide.SlideLayout` objects."""

    def it_is_a_BaseSlide_subclass(self):
        slide_layout = SlideLayout(None, None)
        assert isinstance(slide_layout, _BaseSlide)

    def it_can_iterate_its_clonable_placeholders(self, cloneable_fixture):
        slide_layout, expected_placeholders = cloneable_fixture
        cloneable = list(slide_layout.iter_cloneable_placeholders())
        assert cloneable == expected_placeholders

    def it_provides_access_to_its_placeholders(self, LayoutPlaceholders_, placeholders_):
        sldLayout = element("p:sldLayout/p:cSld/p:spTree")
        spTree = sldLayout.xpath("//p:spTree")[0]
        slide_layout = SlideLayout(sldLayout, None)

        placeholders = slide_layout.placeholders

        LayoutPlaceholders_.assert_called_once_with(spTree, slide_layout)
        assert placeholders is placeholders_

    def it_provides_access_to_its_shapes(self, LayoutShapes_, shapes_):
        sldLayout = element("p:sldLayout/p:cSld/p:spTree")
        spTree = sldLayout.xpath("//p:spTree")[0]
        slide_layout = SlideLayout(sldLayout, None)

        shapes = slide_layout.shapes

        LayoutShapes_.assert_called_once_with(spTree, slide_layout)
        assert shapes is shapes_

    def it_provides_access_to_its_slide_master(self, slide_master_, part_prop_):
        part_prop_.return_value.slide_master = slide_master_
        slide_layout = SlideLayout(None, None)

        slide_master = slide_layout.slide_master

        assert slide_master is slide_master_

    def it_provides_access_to_its_header_footer(self):
        slide_layout = SlideLayout(element("p:sldLayout/p:cSld/p:spTree"), None)

        header_footer = slide_layout.header_footer

        assert isinstance(header_footer, _HeaderFooter)
        assert slide_layout.header_footer is header_footer

    def it_knows_which_slides_are_based_on_it(
        self,
        used_by_fixture,
        part_prop_,
        slide_layout_part_,
        package_,
        presentation_part_,
        presentation_,
    ):
        presentation_, slide_layout, expected_value = used_by_fixture
        part_prop_.return_value = slide_layout_part_
        slide_layout_part_.package = package_
        package_.presentation_part = presentation_part_
        presentation_part_.presentation = presentation_

        used_by_slides = slide_layout.used_by_slides

        assert used_by_slides == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ((PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.BODY), (0, 1)),
            ((PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.DATE), (0,)),
            ((PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.OBJECT), (1,)),
            ((PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER), ()),
        ]
    )
    def cloneable_fixture(self, request, placeholders_prop_, placeholder_, placeholder_2_):
        ph_types, expected_indices = request.param
        slide_layout = SlideLayout(None, None)
        placeholder_.element.ph_type = ph_types[0]
        placeholder_2_.element.ph_type = ph_types[1]
        _placeholders = (placeholder_, placeholder_2_)
        expected_placeholders = [_placeholders[idx] for idx in expected_indices]
        placeholders_prop_.return_value = _placeholders
        return slide_layout, expected_placeholders

    @pytest.fixture(params=[(), (0,), (1,), (0, 1)])
    def used_by_fixture(self, request, presentation_, slide_, slide_2_):
        used_by_idxs = request.param
        slides = (slide_, slide_2_)
        slide_layout = SlideLayout(None, None)
        for idx, s in enumerate(slides):
            s.slide_layout = slide_layout if idx in used_by_idxs else None
        presentation_.slides = slides
        expected_value = tuple(s for i, s in enumerate(slides) if i in used_by_idxs)
        return presentation_, slide_layout, expected_value

    # fixture components -----------------------------------

    @pytest.fixture
    def LayoutPlaceholders_(self, request, placeholders_):
        return class_mock(request, "pptx.slide.LayoutPlaceholders", return_value=placeholders_)

    @pytest.fixture
    def LayoutShapes_(self, request, shapes_):
        return class_mock(request, "pptx.slide.LayoutShapes", return_value=shapes_)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_prop_(self, request, slide_layout_part_):
        return property_mock(request, SlideLayout, "part", return_value=slide_layout_part_)

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, LayoutPlaceholder)

    @pytest.fixture
    def placeholder_2_(self, request):
        return instance_mock(request, LayoutPlaceholder)

    @pytest.fixture
    def placeholders_(self, request):
        return instance_mock(request, LayoutPlaceholders)

    @pytest.fixture
    def placeholders_prop_(self, request, placeholders_):
        return property_mock(request, SlideLayout, "placeholders", return_value=placeholders_)

    @pytest.fixture
    def presentation_(self, request):
        return instance_mock(request, Presentation)

    @pytest.fixture
    def presentation_part_(self, request):
        return instance_mock(request, PresentationPart)

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, LayoutShapes)

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_2_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_layout_part_(self, request):
        return instance_mock(request, SlideLayoutPart)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)


class DescribeSlideLayouts(object):
    """Unit-test suite for `pptx.slide.SlideLayouts` objects."""

    def it_supports_len(self, len_fixture):
        slide_layouts, expected_value = len_fixture
        assert len(slide_layouts) == expected_value

    def it_can_iterate_its_slide_layouts(self, part_prop_, slide_master_part_):
        sldLayoutIdLst = element("p:sldLayoutIdLst/(p:sldLayoutId{r:id=a},p:sldLayoutId{r:id=b})")
        _slide_layouts = [
            SlideLayout(element("p:sldLayout"), None),
            SlideLayout(element("p:sldLayout"), None),
        ]
        part_prop_.return_value = slide_master_part_
        related_slide_layout_ = slide_master_part_.related_slide_layout
        related_slide_layout_.side_effect = _slide_layouts
        slide_layouts = SlideLayouts(sldLayoutIdLst, None)

        slide_layout_lst = [sl for sl in slide_layouts]

        assert related_slide_layout_.call_args_list == [call("a"), call("b")]
        assert slide_layout_lst == _slide_layouts

    def it_supports_indexed_access(self, slide_layout_, part_prop_, slide_master_part_):
        part_prop_.return_value = slide_master_part_
        slide_master_part_.related_slide_layout.return_value = slide_layout_
        slide_layouts = SlideLayouts(element("p:sldLayoutIdLst/p:sldLayoutId{r:id=rId1}"), None)

        slide_layout = slide_layouts[0]

        slide_master_part_.related_slide_layout.assert_called_once_with("rId1")
        assert slide_layout is slide_layout_

    def but_it_raises_on_index_out_of_range(self, part_prop_):
        slide_layouts = SlideLayouts(element("p:sldLayoutIdLst/p:sldLayoutId{r:id=rId1}"), None)
        with pytest.raises(IndexError):
            slide_layouts[1]

    def it_can_find_a_slide_layout_by_name(self, _iter_, slide_layout_, slide_layout_2_):
        _iter_.return_value = iter((slide_layout_, slide_layout_2_))
        slide_layout_2_.name = "pick me!"
        slide_layouts = SlideLayouts(None, None)

        slide_layout = slide_layouts.get_by_name("pick me!")

        assert slide_layout is slide_layout_2_

    def but_it_returns_the_default_value_when_no_layout_has_that_name(
        self, _iter_, slide_layout_, slide_layout_2_
    ):
        _iter_.return_value = iter((slide_layout_, slide_layout_2_))
        slide_layout_2_.name = "not the droid you're looking for"
        slide_layouts = SlideLayouts(None, None)

        # ---default-default is None---
        slide_layout = slide_layouts.get_by_name("pick me!")
        assert slide_layout is None

        # ---but default can be specified---
        slide_layout = slide_layouts.get_by_name("pick me!", "default-value")
        assert slide_layout == "default-value"

    def it_knows_the_index_of_each_of_its_slide_layouts(
        self, _iter_, slide_layout_, slide_layout_2_
    ):
        _iter_.return_value = iter((slide_layout_, slide_layout_2_))
        slide_layouts = SlideLayouts(None, None)

        index = slide_layouts.index(slide_layout_2_)

        assert index == 1

    def but_it_raises_on_slide_layout_not_in_collection(
        self, _iter_, slide_layout_, slide_layout_2_
    ):
        _iter_.return_value = iter((slide_layout_,))
        slide_layouts = SlideLayouts(None, None)

        with pytest.raises(ValueError) as e:
            slide_layouts.index(slide_layout_2_)
        assert str(e.value) == "layout not in this SlideLayouts collection"

    def it_can_remove_an_unused_slide_layout(
        self, slide_layout_, index_, slide_master_, slide_master_part_
    ):
        slide_layout_.used_by_slides = ()
        index_.return_value = 0
        sldLayoutIdLst = element(
            "p:sldLayoutIdLst/(p:sldLayoutId{r:id=rId1},p:sldLayoutId{r:id=rId2})"
        )
        slide_layout_.slide_master = slide_master_
        slide_master_.part = slide_master_part_
        slide_layouts = SlideLayouts(sldLayoutIdLst, None)

        slide_layouts.remove(slide_layout_)

        assert slide_layouts._sldLayoutIdLst.xml == xml("p:sldLayoutIdLst/p:sldLayoutId{r:id=rId2}")
        slide_master_part_.drop_rel.assert_called_once_with("rId1")

    def but_it_raises_on_attempt_to_remove_slide_layout_in_use(self, slide_layout_, slide_):
        slide_layout_.used_by_slides = (slide_,)
        slide_layouts = SlideLayouts(None, None)

        with pytest.raises(ValueError):
            slide_layouts.remove(slide_layout_)

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("p:sldLayoutIdLst", 0),
            ("p:sldLayoutIdLst/p:sldLayoutId", 1),
            ("p:sldLayoutIdLst/(p:sldLayoutId,p:sldLayoutId)", 2),
        ]
    )
    def len_fixture(self, request):
        sldLayoutIdLst_cxml, expected_value = request.param
        slide_layouts = SlideLayouts(element(sldLayoutIdLst_cxml), None)
        return slide_layouts, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def index_(self, request):
        return method_mock(request, SlideLayouts, "index")

    @pytest.fixture
    def _iter_(self, request):
        return method_mock(request, SlideLayouts, "__iter__")

    @pytest.fixture
    def part_prop_(self, request):
        return property_mock(request, SlideLayouts, "part")

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_layout_2_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_master_part_(self, request):
        return instance_mock(request, SlideMasterPart)


class DescribeSlideMaster(object):
    """Unit-test suite for `pptx.slide.SlideMaster` objects."""

    def it_is_a_BaseMaster_subclass(self, subclass_fixture):
        slide_master = subclass_fixture
        assert isinstance(slide_master, _BaseMaster)

    def it_provides_access_to_its_slide_layouts(self, layouts_fixture):
        slide_master, SlideLayouts_, sldLayoutIdLst, slide_layouts_ = layouts_fixture
        slide_layouts = slide_master.slide_layouts
        SlideLayouts_.assert_called_once_with(sldLayoutIdLst, slide_master)
        assert slide_layouts is slide_layouts_

    def it_exposes_a_theme_colors_mapping(self, request):
        from pptx.dml.color import RGBColor

        _resolve_ = function_mock(
            request,
            "pptx.slide._resolve_theme_colors",
            return_value={"accent1": RGBColor(0x4F, 0x81, 0xBD)},
        )
        slide_master_part_ = instance_mock(request, SlideMasterPart)
        slide_master = SlideMaster(None, slide_master_part_)

        mapping = slide_master.theme_colors

        _resolve_.assert_called_once_with(slide_master_part_)
        assert mapping == {"accent1": RGBColor(0x4F, 0x81, 0xBD)}

    def it_provides_access_to_its_header_footer(self):
        slide_master = SlideMaster(element("p:sldMaster/p:cSld/p:spTree"), None)

        header_footer = slide_master.header_footer

        assert isinstance(header_footer, _HeaderFooter)
        # ---same instance returned on each access---
        assert slide_master.header_footer is header_footer

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def layouts_fixture(self, SlideLayouts_, slide_layouts_):
        sldMaster = element("p:sldMaster/p:sldLayoutIdLst")
        slide_master = SlideMaster(sldMaster, None)
        sldMasterIdLst = sldMaster.sldLayoutIdLst
        return slide_master, SlideLayouts_, sldMasterIdLst, slide_layouts_

    @pytest.fixture
    def subclass_fixture(self):
        return SlideMaster(None, None)

    # fixture components -----------------------------------

    @pytest.fixture
    def SlideLayouts_(self, request, slide_layouts_):
        return class_mock(request, "pptx.slide.SlideLayouts", return_value=slide_layouts_)

    @pytest.fixture
    def slide_layouts_(self, request):
        return instance_mock(request, SlideLayouts)


class DescribeSlideMasters(object):
    """Unit-test suite for `pptx.slide.SlideMasters` objects."""

    def it_knows_how_many_masters_it_contains(self, len_fixture):
        slide_masters, expected_value = len_fixture
        assert len(slide_masters) == expected_value

    def it_can_iterate_the_slide_masters(self, iter_fixture):
        slide_masters, related_slide_master_, calls, expected_values = iter_fixture
        _slide_masters = [sm for sm in slide_masters]
        assert related_slide_master_.call_args_list == calls
        assert _slide_masters == expected_values

    def it_supports_indexed_access(self, getitem_fixture):
        slide_masters, part_, slide_master_, rId = getitem_fixture
        slide_master = slide_masters[0]
        part_.related_slide_master.assert_called_once_with(rId)
        assert slide_master is slide_master_

    def it_raises_on_index_out_of_range(self, getitem_raises_fixture):
        slides = getitem_raises_fixture
        with pytest.raises(IndexError):
            slides[1]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getitem_fixture(self, part_, slide_master_, part_prop_):
        slide_masters = SlideMasters(element("p:sldMasterIdLst/p:sldMasterId{r:id=rId1}"), None)
        part_.related_slide_master.return_value = slide_master_
        return slide_masters, part_, slide_master_, "rId1"

    @pytest.fixture
    def getitem_raises_fixture(self, part_prop_):
        return SlideMasters(element("p:sldMasterIdLst/p:sldMasterId{r:id=rId1}"), None)

    @pytest.fixture
    def iter_fixture(self, part_prop_):
        sldMasterIdLst = element("p:sldMasterIdLst/(p:sldMasterId{r:id=a},p:sldMasterId{r:id=b})")
        slide_masters = SlideMasters(sldMasterIdLst, None)
        related_slide_master_ = part_prop_.return_value.related_slide_master
        calls = [call("a"), call("b")]
        _slide_masters = [
            SlideMaster(element("p:sldMaster"), None),
            SlideMaster(element("p:sldMaster"), None),
        ]
        related_slide_master_.side_effect = _slide_masters
        return slide_masters, related_slide_master_, calls, _slide_masters

    @pytest.fixture(
        params=[
            ("p:sldMasterIdLst", 0),
            ("p:sldMasterIdLst/p:sldMasterId", 1),
            ("p:sldMasterIdLst/(p:sldMasterId,p:sldMasterId)", 2),
        ]
    )
    def len_fixture(self, request):
        sldMasterIdLst_cxml, expected_value = request.param
        slide_masters = SlideMasters(element(sldMasterIdLst_cxml), None)
        return slide_masters, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, PresentationPart)

    @pytest.fixture
    def part_prop_(self, request, part_):
        return property_mock(request, SlideMasters, "part", return_value=part_)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)


class Describe_Background(object):
    """Unit-test suite for `pptx.slide._Background` objects."""

    @pytest.mark.parametrize(
        "cSld_xml, expected_cxml",
        (
            ("p:cSld{a:b=c}", "p:cSld{a:b=c}/p:bg/p:bgPr/(a:noFill,a:effectLst)"),
            (
                "p:cSld{a:b=c}/p:bg/p:bgRef",
                "p:cSld{a:b=c}/p:bg/p:bgPr/(a:noFill,a:effectLst)",
            ),
            ("p:cSld/p:bg/p:bgPr/a:solidFill", "p:cSld/p:bg/p:bgPr/a:solidFill"),
        ),
    )
    def it_provides_access_to_its_fill(self, request, cSld_xml, expected_cxml):
        fill_ = instance_mock(request, FillFormat)
        from_fill_parent_ = method_mock(
            request, FillFormat, "from_fill_parent", autospec=False, return_value=fill_
        )
        cSld = element(cSld_xml)
        background = _Background(cSld)

        fill = background.fill

        assert cSld.xml == xml(expected_cxml)
        from_fill_parent_.assert_called_once_with(cSld.xpath("p:bg/p:bgPr")[0])
        assert fill is fill_


class Describe_resolve_theme_colors(object):
    """Unit-test suite for `pptx.slide._resolve_theme_colors` helper."""

    def it_returns_an_empty_mapping_when_no_theme_is_related(self, request):
        from pptx.slide import _resolve_theme_colors

        master_part_ = loose_mock(request, name="slide_master_part_")
        master_part_.part_related_by.side_effect = KeyError("no theme")

        assert _resolve_theme_colors(master_part_) == {}

    def it_returns_the_srgb_sys_and_prst_values_of_the_clrScheme(self, request):
        from pptx.dml.color import RGBColor
        from pptx.oxml import parse_xml
        from pptx.slide import _resolve_theme_colors

        theme_xml = (
            "<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>"
            "  <a:themeElements>"
            "    <a:clrScheme name='Office'>"
            "      <a:dk1><a:sysClr val='windowText' lastClr='000000'/></a:dk1>"
            "      <a:lt1><a:sysClr val='window' lastClr='FFFFFF'/></a:lt1>"
            "      <a:dk2><a:srgbClr val='1F497D'/></a:dk2>"
            "      <a:lt2><a:srgbClr val='EEECE1'/></a:lt2>"
            "      <a:accent1><a:srgbClr val='4F81BD'/></a:accent1>"
            "      <a:accent2><a:prstClr val='cornflowerBlue'/></a:accent2>"
            "    </a:clrScheme>"
            "  </a:themeElements>"
            "</a:theme>"
        )
        master_xml = (
            "<p:sldMaster xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>"
            "  <p:clrMap bg1='lt1' tx1='dk1' bg2='lt2' tx2='dk2'/>"
            "</p:sldMaster>"
        )
        theme_elm = parse_xml(theme_xml)
        master_elm = parse_xml(master_xml)

        theme_part_ = loose_mock(request, name="theme_part_")
        theme_part_._element = theme_elm
        master_part_ = loose_mock(request, name="slide_master_part_")
        master_part_.part_related_by.return_value = theme_part_
        master_part_._element = master_elm

        colors = _resolve_theme_colors(master_part_)

        assert colors["dk1"] == RGBColor(0, 0, 0)
        assert colors["lt1"] == RGBColor(0xFF, 0xFF, 0xFF)
        assert colors["dk2"] == RGBColor(0x1F, 0x49, 0x7D)
        assert colors["accent1"] == RGBColor(0x4F, 0x81, 0xBD)
        # -- preset-color scheme entry resolves via PRESET_COLORS --
        assert colors["accent2"] == RGBColor(0x64, 0x95, 0xED)
        # -- clrMap alias applied --
        assert colors["bg1"] == colors["lt1"]
        assert colors["tx1"] == colors["dk1"]
        assert colors["bg2"] == colors["lt2"]
        assert colors["tx2"] == colors["dk2"]


class Describe_HeaderFooter(object):
    """Unit-test suite for `pptx.slide._HeaderFooter` objects."""

    @pytest.mark.parametrize(
        ("sldMaster_cxml", "flag_name", "expected"),
        [
            ("p:sldMaster/p:cSld/p:spTree", "slide_number_visible", True),
            ("p:sldMaster/p:cSld/p:spTree", "header_visible", True),
            ("p:sldMaster/p:cSld/p:spTree", "footer_visible", True),
            ("p:sldMaster/p:cSld/p:spTree", "date_visible", True),
            (
                "p:sldMaster/(p:cSld/p:spTree,p:hf{sldNum=0})",
                "slide_number_visible",
                False,
            ),
            (
                "p:sldMaster/(p:cSld/p:spTree,p:hf{hdr=0})",
                "header_visible",
                False,
            ),
            (
                "p:sldMaster/(p:cSld/p:spTree,p:hf{ftr=0})",
                "footer_visible",
                False,
            ),
            (
                "p:sldMaster/(p:cSld/p:spTree,p:hf{dt=0})",
                "date_visible",
                False,
            ),
        ],
    )
    def it_reads_the_visibility_flags(
        self, sldMaster_cxml: str, flag_name: str, expected: bool
    ):
        sldMaster = element(sldMaster_cxml)
        hf_proxy = _HeaderFooter(sldMaster)
        assert getattr(hf_proxy, flag_name) is expected

    @pytest.mark.parametrize(
        "flag_name,attr_name",
        [
            ("slide_number_visible", "sldNum"),
            ("header_visible", "hdr"),
            ("footer_visible", "ftr"),
            ("date_visible", "dt"),
        ],
    )
    def it_sets_the_flag_to_false_by_adding_a_hf_element(self, flag_name, attr_name):
        sldMaster = element("p:sldMaster/p:cSld/p:spTree")
        hf_proxy = _HeaderFooter(sldMaster)

        setattr(hf_proxy, flag_name, False)

        assert sldMaster.hf is not None
        assert getattr(sldMaster.hf, attr_name) is False

    def it_removes_the_hf_element_when_every_flag_is_True_again(self):
        sldMaster = element("p:sldMaster/(p:cSld/p:spTree,p:hf{sldNum=0})")
        hf_proxy = _HeaderFooter(sldMaster)

        hf_proxy.slide_number_visible = True

        assert sldMaster.hf is None
        assert hf_proxy.slide_number_visible is True

    def it_raises_TypeError_on_non_bool_assignment(self):
        sldMaster = element("p:sldMaster/p:cSld/p:spTree")
        hf_proxy = _HeaderFooter(sldMaster)

        with pytest.raises(TypeError, match="must be a bool"):
            hf_proxy.slide_number_visible = "no"  # pyright: ignore[reportAttributeAccessIssue]


class DescribeSlide_F8_transition_and_animations(object):
    """Unit-test suite for Foundation F8 additions to `pptx.slide.Slide`."""

    # -- Slide.has_animations -----------------------------------------

    def it_reports_False_for_has_animations_when_no_timing(self):
        slide = Slide(element("p:sld/p:cSld/p:spTree"), None)
        assert slide.has_animations is False

    def it_reports_False_for_has_animations_when_timing_but_no_tnLst(self):
        slide = Slide(
            element("p:sld/(p:cSld/p:spTree,p:timing)"),
            None,
        )
        assert slide.has_animations is False

    def it_reports_False_when_tnLst_is_empty(self):
        slide = Slide(
            element("p:sld/(p:cSld/p:spTree,p:timing/p:tnLst)"),
            None,
        )
        assert slide.has_animations is False

    def it_reports_True_when_timing_has_a_tnLst_child(self):
        slide = Slide(
            element("p:sld/(p:cSld/p:spTree,p:timing/p:tnLst/p:par/p:cTn{id=1})"),
            None,
        )
        assert slide.has_animations is True

    # -- Slide.timing_xml ---------------------------------------------

    def it_returns_None_for_timing_xml_when_no_timing(self):
        slide = Slide(element("p:sld/p:cSld/p:spTree"), None)
        assert slide.timing_xml is None

    def it_returns_the_xml_of_the_timing_subtree_when_present(self):
        slide = Slide(
            element("p:sld/(p:cSld/p:spTree,p:timing/p:tnLst)"),
            None,
        )
        timing_xml = slide.timing_xml
        assert timing_xml is not None
        assert "<p:timing" in timing_xml
        assert "<p:tnLst" in timing_xml

    # -- Slide.transition (proxy return) ------------------------------

    def it_returns_a_Transition_proxy(self):
        slide = Slide(element("p:sld/p:cSld/p:spTree"), None)
        assert isinstance(slide.transition, Transition)

    def it_returns_the_same_Transition_proxy_on_subsequent_access(self):
        slide = Slide(element("p:sld/p:cSld/p:spTree"), None)
        assert slide.transition is slide.transition


class DescribeTransition(object):
    """Unit-test suite for `pptx.slide.Transition` (Foundation F8 MVP)."""

    # -- .type read -------------------------------------------------------

    def it_reports_NONE_type_when_no_transition_element(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        assert transition.type is PP_TRANSITION_TYPE.NONE

    def it_reports_NONE_type_when_transition_has_no_variant_child(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition)")
        transition = Transition(sld)
        assert transition.type is PP_TRANSITION_TYPE.NONE

    @pytest.mark.parametrize(
        ("variant_tag", "expected"),
        [
            ("p:fade", PP_TRANSITION_TYPE.FADE),
            ("p:push", PP_TRANSITION_TYPE.PUSH),
            ("p:wipe", PP_TRANSITION_TYPE.WIPE),
            ("p:cover", PP_TRANSITION_TYPE.COVER),
            ("p:wheel", PP_TRANSITION_TYPE.WHEEL),
            ("p:dissolve", PP_TRANSITION_TYPE.DISSOLVE),
        ],
    )
    def it_reads_the_transition_type_for_p_variants(
        self, variant_tag: str, expected: PP_TRANSITION_TYPE
    ):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/%s)" % variant_tag)
        transition = Transition(sld)
        assert transition.type is expected

    # -- .type write ------------------------------------------------------

    def it_sets_a_p_variant(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.FADE
        assert transition.type is PP_TRANSITION_TYPE.FADE
        assert sld.transition is not None
        assert sld.transition.variant_tag == "p:fade"

    def it_sets_the_p14_morph_variant_wrapped_in_alt_content(self):
        # -- setting MORPH moves the p:transition inside an
        # -- mc:AlternateContent, so the direct-child descriptor is None.
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        assert transition.type is PP_TRANSITION_TYPE.MORPH
        assert sld.transition is None  # no direct child now
        assert sld.transition_effective is not None
        assert sld.transition_is_alt_content_wrapped is True
        assert sld.transition_effective.variant_tag == "p14:morph"

    def it_replaces_an_existing_variant(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.PUSH
        assert transition.type is PP_TRANSITION_TYPE.PUSH

    def it_removes_the_transition_element_when_set_to_NONE(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.NONE
        assert sld.transition is None

    def it_preserves_the_transition_element_when_clearing_but_attrs_remain(self):
        sld = element(
            "p:sld/(p:cSld/p:spTree,p:transition{advTm=2000}/p:fade)"
        )
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.NONE
        # -- transition still exists because @advTm is a meaningful attr --
        assert sld.transition is not None
        assert sld.transition.advTm == 2000
        assert transition.type is PP_TRANSITION_TYPE.NONE

    def it_rejects_non_enum_type_assignments(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        with pytest.raises(ValueError):
            transition.type = "fade"  # type: ignore[assignment]

    # -- .duration --------------------------------------------------------

    def it_returns_None_duration_when_no_transition(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        assert transition.duration is None

    def it_reads_and_writes_duration(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.duration = 800
        assert transition.duration == 800

    def it_removes_duration_when_set_to_None(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.duration = 800
        transition.duration = None
        assert transition.duration is None

    # -- .advance_on_click -----------------------------------------------

    def it_defaults_advance_on_click_to_True(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        assert transition.advance_on_click is True

    def it_reads_advance_on_click_when_explicitly_false(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition{advClick=0})")
        transition = Transition(sld)
        assert transition.advance_on_click is False

    def it_writes_advance_on_click(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.advance_on_click = False
        assert sld.transition is not None
        assert sld.transition.advClick is False

    def it_rejects_non_bool_advance_on_click(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        with pytest.raises(TypeError, match="must be a bool"):
            transition.advance_on_click = "yes"  # type: ignore[assignment]

    # -- .advance_after_time ----------------------------------------------

    def it_returns_None_advance_after_time_when_absent(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        assert transition.advance_after_time is None

    def it_reads_advance_after_time(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition{advTm=5000})")
        transition = Transition(sld)
        assert transition.advance_after_time == 5000

    def it_writes_advance_after_time(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.advance_after_time = 3000
        assert transition.advance_after_time == 3000

    def it_clears_advance_after_time_when_set_to_None(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition{advTm=5000})")
        transition = Transition(sld)
        transition.advance_after_time = None
        assert transition.advance_after_time is None

    def it_rejects_negative_advance_after_time(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        with pytest.raises(ValueError, match="non-negative"):
            transition.advance_after_time = -1

    # -- .speed -----------------------------------------------------------

    def it_reports_FAST_speed_when_no_transition(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        assert transition.speed is PP_TRANSITION_SPEED.FAST

    def it_reports_FAST_speed_when_spd_absent(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition)")
        transition = Transition(sld)
        assert transition.speed is PP_TRANSITION_SPEED.FAST

    @pytest.mark.parametrize(
        ("spd", "expected"),
        [
            ("slow", PP_TRANSITION_SPEED.SLOW),
            ("med", PP_TRANSITION_SPEED.MEDIUM),
            ("fast", PP_TRANSITION_SPEED.FAST),
        ],
    )
    def it_reads_an_explicit_spd(
        self, spd: str, expected: PP_TRANSITION_SPEED
    ):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition{spd=%s})" % spd)
        transition = Transition(sld)
        assert transition.speed is expected

    @pytest.mark.parametrize(
        ("value", "xml_value"),
        [
            (PP_TRANSITION_SPEED.SLOW, "slow"),
            (PP_TRANSITION_SPEED.MEDIUM, "med"),
            (PP_TRANSITION_SPEED.FAST, "fast"),
        ],
    )
    def it_writes_speed(
        self, value: PP_TRANSITION_SPEED, xml_value: str
    ):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.speed = value
        assert sld.transition is not None
        assert sld.transition.spd == xml_value

    def it_rejects_non_enum_speed_assignments(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        with pytest.raises(ValueError):
            transition.speed = "slow"  # type: ignore[assignment]

    # -- .wipe_direction --------------------------------------------------

    def it_returns_None_wipe_direction_when_no_transition(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        assert transition.wipe_direction is None

    def it_returns_None_wipe_direction_when_variant_is_not_wipe(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        transition = Transition(sld)
        assert transition.wipe_direction is None

    def it_defaults_wipe_direction_to_LEFT_when_dir_absent(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:wipe)")
        transition = Transition(sld)
        assert transition.wipe_direction is PP_TRANSITION_SIDE_DIRECTION.LEFT

    @pytest.mark.parametrize(
        ("dir_val", "expected"),
        [
            ("l", PP_TRANSITION_SIDE_DIRECTION.LEFT),
            ("u", PP_TRANSITION_SIDE_DIRECTION.UP),
            ("r", PP_TRANSITION_SIDE_DIRECTION.RIGHT),
            ("d", PP_TRANSITION_SIDE_DIRECTION.DOWN),
        ],
    )
    def it_reads_an_explicit_wipe_direction(
        self, dir_val: str, expected: PP_TRANSITION_SIDE_DIRECTION
    ):
        sld = element(
            "p:sld/(p:cSld/p:spTree,p:transition/p:wipe{dir=%s})" % dir_val
        )
        transition = Transition(sld)
        assert transition.wipe_direction is expected

    def it_writes_wipe_direction_on_existing_wipe(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:wipe)")
        transition = Transition(sld)
        transition.wipe_direction = PP_TRANSITION_SIDE_DIRECTION.UP
        assert transition.wipe_direction is PP_TRANSITION_SIDE_DIRECTION.UP
        wipe = sld.transition.find(qn("p:wipe"))
        assert wipe is not None
        assert wipe.get("dir") == "u"

    def it_creates_wipe_variant_when_setting_direction_with_no_transition(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.wipe_direction = PP_TRANSITION_SIDE_DIRECTION.RIGHT
        assert transition.type is PP_TRANSITION_TYPE.WIPE
        assert transition.wipe_direction is PP_TRANSITION_SIDE_DIRECTION.RIGHT

    def it_replaces_existing_variant_when_setting_wipe_direction(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        transition = Transition(sld)
        transition.wipe_direction = PP_TRANSITION_SIDE_DIRECTION.DOWN
        assert transition.type is PP_TRANSITION_TYPE.WIPE
        assert transition.wipe_direction is PP_TRANSITION_SIDE_DIRECTION.DOWN

    def it_rejects_non_enum_wipe_direction(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        with pytest.raises(ValueError):
            transition.wipe_direction = "left"  # type: ignore[assignment]


class DescribeTransition_morph_option_942(object):
    """Unit-test suite for Transition MORPH support (issue #942)."""

    # -- .morph_option read ----------------------------------------------

    def it_returns_None_morph_option_when_transition_is_not_MORPH(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        transition = Transition(sld)
        assert transition.morph_option is None

    def it_returns_None_morph_option_when_no_transition(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        assert transition.morph_option is None

    def it_defaults_morph_option_to_byObject_on_a_MORPH_transition(self):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        # -- slide with a wrapped MORPH transition but no @option attr --
        sld = parse_xml(
            "<p:sld %s>"
            "  <p:cSld><p:spTree/></p:cSld>"
            "  <mc:AlternateContent>"
            '    <mc:Choice Requires="p14">'
            '      <p:transition><p14:morph/></p:transition>'
            "    </mc:Choice>"
            "    <mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>"
            "  </mc:AlternateContent>"
            "</p:sld>" % nsdecls("p", "p14", "mc")
        )
        transition = Transition(sld)
        assert transition.morph_option == "byObject"

    @pytest.mark.parametrize("option", ["byObject", "byWord", "byChar"])
    def it_reads_an_explicit_morph_option(self, option: str):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        sld = parse_xml(
            "<p:sld %s>"
            "  <p:cSld><p:spTree/></p:cSld>"
            "  <mc:AlternateContent>"
            '    <mc:Choice Requires="p14">'
            '      <p:transition><p14:morph option="%s"/></p:transition>'
            "    </mc:Choice>"
            "    <mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>"
            "  </mc:AlternateContent>"
            "</p:sld>" % (nsdecls("p", "p14", "mc"), option)
        )
        transition = Transition(sld)
        assert transition.morph_option == option

    # -- .morph_option write ---------------------------------------------

    @pytest.mark.parametrize("option", ["byObject", "byWord", "byChar"])
    def it_writes_the_morph_option(self, option: str):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        transition.morph_option = option
        assert transition.morph_option == option

    def it_rejects_an_unknown_morph_option(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        with pytest.raises(ValueError, match="morph_option must be one of"):
            transition.morph_option = "byParagraph"

    def it_rejects_morph_option_write_when_not_a_MORPH_transition(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.FADE
        with pytest.raises(ValueError, match="only applies to a MORPH"):
            transition.morph_option = "byWord"

    # -- type = MORPH wraps in mc:AlternateContent ----------------------

    def it_wraps_in_alt_content_when_type_is_set_to_MORPH(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        assert sld.transition_is_alt_content_wrapped is True
        # -- mc:Fallback contains a p:transition/p:fade --
        ac = sld._transition_alt_content_lst[0]
        fallback = ac.find(qn("mc:Fallback"))
        assert fallback is not None
        fb_transition = fallback.find(qn("p:transition"))
        assert fb_transition is not None
        assert fb_transition.find(qn("p:fade")) is not None

    def it_wraps_an_existing_fade_transition_when_switching_to_MORPH(self):
        # -- the preceding p:transition's attrs (e.g. advClick) are
        # -- preserved on the mc:Choice/p:transition and copied onto the
        # -- mc:Fallback/p:transition --
        sld = element(
            "p:sld/(p:cSld/p:spTree,p:transition{advClick=0}/p:fade)"
        )
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        assert sld.transition_is_alt_content_wrapped is True
        effective = sld.transition_effective
        assert effective is not None
        assert effective.advClick is False
        # -- fallback carries the same advance-control --
        ac = sld._transition_alt_content_lst[0]
        fb_transition = ac.find(qn("mc:Fallback")).find(qn("p:transition"))
        assert fb_transition.get("advClick") == "0"

    def it_unwraps_when_switching_from_MORPH_to_a_plain_variant(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        assert sld.transition_is_alt_content_wrapped is True

        transition.type = PP_TRANSITION_TYPE.FADE
        assert sld.transition_is_alt_content_wrapped is False
        assert sld.transition is not None
        assert sld.transition.variant_tag == "p:fade"

    def it_removes_the_wrapper_when_set_to_NONE_from_MORPH(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        transition.type = PP_TRANSITION_TYPE.NONE
        assert transition.type is PP_TRANSITION_TYPE.NONE
        assert sld.transition_effective is None
        assert sld._transition_alt_content_lst == []

    # -- duration and advance still work on a wrapped MORPH -------------

    def it_reads_and_writes_duration_through_the_wrapper(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        transition.duration = 2000
        assert transition.duration == 2000

    def it_reads_and_writes_advance_on_click_through_the_wrapper(self):
        sld = element("p:sld/p:cSld/p:spTree")
        transition = Transition(sld)
        transition.type = PP_TRANSITION_TYPE.MORPH
        transition.advance_on_click = False
        assert transition.advance_on_click is False

    def it_reads_MORPH_as_the_type_of_a_wrapped_transition(self):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        sld = parse_xml(
            "<p:sld %s>"
            "  <p:cSld><p:spTree/></p:cSld>"
            "  <mc:AlternateContent>"
            '    <mc:Choice Requires="p14">'
            '      <p:transition><p14:morph option="byWord"/></p:transition>'
            "    </mc:Choice>"
            "    <mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>"
            "  </mc:AlternateContent>"
            "</p:sld>" % nsdecls("p", "p14", "mc")
        )
        transition = Transition(sld)
        assert transition.type is PP_TRANSITION_TYPE.MORPH
        assert transition.morph_option == "byWord"
