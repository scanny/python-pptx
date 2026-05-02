# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.viewprops` module."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.oxml.viewprops import CT_ViewProperties
from pptx.parts.viewprops import ViewPropsPart


class DescribeViewPropsPart(object):
    """Unit-test suite for `pptx.parts.viewprops.ViewPropsPart` objects."""

    def it_can_construct_a_default_view_props_part(self):
        view_props = ViewPropsPart.default(None)  # type: ignore[arg-type]

        assert isinstance(view_props, ViewPropsPart)
        assert view_props.content_type == CT.PML_VIEW_PROPS
        assert view_props.partname == "/ppt/viewProps.xml"
        assert isinstance(view_props._element, CT_ViewProperties)
        # -- schema defaults --
        assert view_props._element.lastView == "sldView"
        assert view_props._element.showComments is True
