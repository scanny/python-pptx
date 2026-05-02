# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.viewprops` module."""

from __future__ import annotations

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.viewprops import (
    CT_CommonViewProperties,
    CT_Scale2D,
    CT_ViewProperties,
)


class DescribeCT_ViewProperties(object):
    """Unit-test suite for `CT_ViewProperties` objects."""

    def it_can_construct_a_new_default_viewPr_element(self):
        viewPr = CT_ViewProperties.new_default()

        assert isinstance(viewPr, CT_ViewProperties)
        assert viewPr.tag == (
            "{http://schemas.openxmlformats.org/presentationml/2006/" "main}viewPr"
        )
        # -- no @lastView => schema default --
        assert viewPr.lastView == "sldView"
        # -- no @showComments => schema default --
        assert viewPr.showComments is True

    def it_exposes_lastView_read_write(self):
        viewPr = CT_ViewProperties.new_default()

        viewPr.lastView = "notesView"

        assert viewPr.lastView == "notesView"
        # -- round-trip through the XML to confirm persistence --
        xml = parse_xml(('<p:viewPr %s lastView="outlineView"/>' % nsdecls("p")).encode())
        assert xml.lastView == "outlineView"

    def it_reverts_lastView_to_schema_default_when_set_to_sldView(self):
        viewPr = CT_ViewProperties.new_default()
        viewPr.lastView = "handoutView"

        viewPr.lastView = "sldView"

        # -- assigning the default removes the attribute --
        assert "lastView" not in viewPr.attrib
        assert viewPr.lastView == "sldView"

    def it_exposes_showComments_read_write(self):
        viewPr = CT_ViewProperties.new_default()

        viewPr.showComments = False

        assert viewPr.showComments is False

    @pytest.mark.parametrize(
        "child_tag",
        [
            "normalViewPr",
            "slideViewPr",
            "outlineViewPr",
            "sorterViewPr",
            "notesViewPr",
        ],
    )
    def it_can_get_or_add_each_child_view_element(self, child_tag: str):
        viewPr = CT_ViewProperties.new_default()

        method = getattr(viewPr, "get_or_add_%s" % child_tag)
        child = method()

        # -- same node is returned on subsequent calls --
        assert method() is child
        # -- and it is an element with the right qname --
        assert child.tag.endswith("}" + child_tag)


class DescribeCT_SlideSorterViewProperties(object):
    """Unit-test suite for `CT_SlideSorterViewProperties` objects."""

    def it_returns_True_for_showFormatting_when_attribute_is_absent(self):
        viewPr = CT_ViewProperties.new_default()
        sorter = viewPr.get_or_add_sorterViewPr()

        assert sorter.showFormatting is True

    def it_round_trips_showFormatting(self):
        viewPr = CT_ViewProperties.new_default()
        sorter = viewPr.get_or_add_sorterViewPr()

        sorter.showFormatting = False

        assert sorter.showFormatting is False


class DescribeCT_Scale2D(object):
    """Unit-test suite for `CT_Scale2D` objects."""

    def it_reads_zoom_as_a_ratio_from_sx(self):
        scale = self._scale_with_ratio(124, 100)

        assert scale.zoom == pytest.approx(1.24)

    def it_returns_one_point_zero_when_sx_is_absent(self):
        scale = parse_xml("<p:scale %s/>" % nsdecls("p"))

        assert scale.zoom == 1.0

    def it_returns_one_point_zero_when_sx_denominator_is_zero(self):
        # -- mildly pathological input: d=0 would blow up naive division --
        scale = self._scale_with_ratio(0, 0)

        assert scale.zoom == 1.0

    def it_writes_zoom_into_both_sx_and_sy(self):
        scale = self._scale_with_ratio(100, 100)

        scale.set_zoom(1.5)

        assert scale.zoom == pytest.approx(1.5)
        # -- both axes are kept in lockstep --
        assert scale.sx is not None
        assert scale.sy is not None
        assert scale.sx.n == scale.sy.n
        assert scale.sx.d == scale.sy.d

    # -- helpers ---------------------------------------------------------

    def _scale_with_ratio(self, n: int, d: int) -> CT_Scale2D:
        xml = '<p:scale %s><a:sx n="%d" d="%d"/><a:sy n="%d" d="%d"/>' "</p:scale>" % (
            nsdecls("p", "a"),
            n,
            d,
            n,
            d,
        )
        return parse_xml(xml)


class DescribeCT_CommonViewProperties(object):
    """Unit-test suite for `CT_CommonViewProperties` objects."""

    def it_can_construct_a_new_default_cViewPr_with_unity_zoom(self):
        cViewPr = CT_CommonViewProperties.new_default()

        assert isinstance(cViewPr, CT_CommonViewProperties)
        scale = cViewPr.scale
        assert scale is not None
        assert scale.zoom == 1.0
