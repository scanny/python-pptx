"""Unit-test suite for `Presentation.view_props` and `first_slide_num` (issue #94).

Exercises :class:`.ViewProps` (view_type / show_comments / show_formatting /
per-view zoom) and the :attr:`.Presentation.first_slide_num` accessor end-
to-end against a real ``pptx.Presentation()`` — the default template
already carries a ``ppt/viewProps.xml`` part, so every property has a
non-trivial read path and a write path that must round-trip through
save + reopen.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation as open_presentation
from pptx.enum.presentation import PP_VIEW_TYPE
from pptx.presentation import ViewProps


class DescribeViewProps(object):
    """Unit-test suite for `pptx.presentation.ViewProps`."""

    def it_is_reachable_from_Presentation_and_returns_a_ViewProps(self):
        prs = open_presentation()

        vp = prs.view_props

        assert isinstance(vp, ViewProps)
        # -- same instance on subsequent access (lazyproperty) --
        assert prs.view_props is vp

    def it_reads_view_type_from_the_default_template(self):
        prs = open_presentation()

        # -- the default template stores `lastView="sldThumbnailView"` --
        assert prs.view_props.view_type == PP_VIEW_TYPE.SLIDE_THUMBNAIL

    def it_writes_view_type_and_round_trips_through_save(self):
        prs = open_presentation()
        prs.view_props.view_type = PP_VIEW_TYPE.OUTLINE

        reopened = _roundtrip(prs)

        assert reopened.view_props.view_type == PP_VIEW_TYPE.OUTLINE

    def it_rejects_a_non_PP_VIEW_TYPE_value_for_view_type(self):
        prs = open_presentation()

        with pytest.raises(TypeError, match="PP_VIEW_TYPE"):
            prs.view_props.view_type = "sldView"  # type: ignore[assignment]

    def it_exposes_show_comments_with_a_default_of_True(self):
        prs = open_presentation()

        # -- the default template omits `@showComments`, so schema default
        # -- applies.
        assert prs.view_props.show_comments is True

    def it_round_trips_show_comments_through_save(self):
        prs = open_presentation()
        prs.view_props.show_comments = False

        reopened = _roundtrip(prs)

        assert reopened.view_props.show_comments is False

    def it_rejects_a_non_bool_value_for_show_comments(self):
        prs = open_presentation()

        with pytest.raises(TypeError, match="show_comments must be a bool"):
            prs.view_props.show_comments = 0  # type: ignore[assignment]

    def it_returns_True_for_show_formatting_when_sorterViewPr_is_absent(self):
        prs = open_presentation()

        # -- default template does not write `p:sorterViewPr` --
        assert prs.view_props.show_formatting is True

    def it_round_trips_show_formatting_through_save(self):
        prs = open_presentation()
        prs.view_props.show_formatting = False

        reopened = _roundtrip(prs)

        assert reopened.view_props.show_formatting is False

    def it_rejects_a_non_bool_value_for_show_formatting(self):
        prs = open_presentation()

        with pytest.raises(TypeError, match="show_formatting must be a bool"):
            prs.view_props.show_formatting = 1  # type: ignore[assignment]

    @pytest.mark.parametrize(
        "attr_name",
        [
            "slide_view_zoom",
            "notes_view_zoom",
            "outline_view_zoom",
            "sorter_view_zoom",
        ],
    )
    def it_returns_1_point_0_zoom_when_the_element_chain_is_absent(self, attr_name: str):
        # -- a fresh default `p:viewPr` element has no per-view subtrees, so
        # -- every zoom-read path falls through to the schema default of 1.0.
        from pptx.oxml.viewprops import CT_ViewProperties
        from pptx.parts.viewprops import ViewPropsPart

        part = ViewPropsPart.default(None)  # type: ignore[arg-type]
        # -- ensure we started from a truly empty root element --
        assert part._element.slideViewPr is None  # pyright: ignore[reportPrivateUsage]
        assert isinstance(part._element, CT_ViewProperties)  # pyright: ignore[reportPrivateUsage]

        vp = ViewProps(part)

        assert getattr(vp, attr_name) == 1.0

    def it_reads_slide_view_zoom_from_the_default_template(self):
        prs = open_presentation()

        # -- the default template's slide-view zoom is 1.24 (sx n=124 d=100) --
        assert prs.view_props.slide_view_zoom == pytest.approx(1.24)

    @pytest.mark.parametrize(
        "attr_name",
        [
            "slide_view_zoom",
            "notes_view_zoom",
            "outline_view_zoom",
            "sorter_view_zoom",
        ],
    )
    def it_round_trips_per_view_zoom_through_save(self, attr_name: str):
        prs = open_presentation()
        setattr(prs.view_props, attr_name, 1.35)

        reopened = _roundtrip(prs)

        assert getattr(reopened.view_props, attr_name) == pytest.approx(1.35)

    @pytest.mark.parametrize(
        "bad_value",
        [0, -0.5, "1.0", None, True],
    )
    def it_rejects_a_bad_zoom_value(self, bad_value: object):
        prs = open_presentation()

        with pytest.raises((TypeError, ValueError)):
            prs.view_props.slide_view_zoom = bad_value  # type: ignore[assignment]

    def it_exposes_the_underlying_element_as_a_round_trip_escape_hatch(self):
        prs = open_presentation()

        element = prs.view_props.element

        assert element.tag.endswith("}viewPr")


class DescribePresentationFirstSlideNum(object):
    """Unit-test suite for :attr:`.Presentation.first_slide_num` (issue #94)."""

    def it_reads_the_schema_default_of_1_when_attribute_is_absent(self):
        prs = open_presentation()

        assert prs.first_slide_num == 1

    def it_round_trips_a_non_default_first_slide_num_through_save(self):
        prs = open_presentation()

        prs.first_slide_num = 7
        reopened = _roundtrip(prs)

        assert reopened.first_slide_num == 7

    def it_removes_the_attribute_when_set_back_to_1(self):
        prs = open_presentation()
        prs.first_slide_num = 5
        assert prs._element.get("firstSlideNum") == "5"

        prs.first_slide_num = 1

        # -- setting the schema default removes the attribute --
        assert prs._element.get("firstSlideNum") is None
        assert prs.first_slide_num == 1

    @pytest.mark.parametrize("bad_value", ["1", 1.5, None, True])
    def it_rejects_a_non_int_first_slide_num(self, bad_value: object):
        prs = open_presentation()

        with pytest.raises(TypeError, match="first_slide_num"):
            prs.first_slide_num = bad_value  # type: ignore[assignment]


def _roundtrip(prs):
    """Save `prs` to a BytesIO and re-open, returning the reopened Presentation."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return open_presentation(buf)
