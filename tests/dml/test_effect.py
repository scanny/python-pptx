"""Unit-test suite for `pptx.dml.effect` module."""

from __future__ import annotations

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.effect import (
    EffectFormat,
    GlowFormat,
    ReflectionFormat,
    ShadowFormat,
    SoftEdgeFormat,
)
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu

from ..unitutil.cxml import element, xml


class DescribeShadowFormat(object):
    def it_knows_whether_it_inherits(self, inherit_get_fixture):
        shadow, expected_value = inherit_get_fixture
        inherit = shadow.inherit
        assert inherit is expected_value

    def it_can_change_whether_it_inherits(self, inherit_set_fixture):
        shadow, value, expected_xml = inherit_set_fixture
        shadow.inherit = value
        assert shadow._element.xml == expected_xml

    def it_reports_blur_radius_as_None_when_no_outerShdw_is_defined(self):
        shadow = ShadowFormat(element("p:spPr"))
        assert shadow.blur_radius is None

    def it_reads_blur_radius_from_outerShdw(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst/a:outerShdw{blurRad=50800}"))
        assert shadow.blur_radius == Emu(50800)

    def it_writes_blur_radius_creating_outerShdw_and_effectLst(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst"))
        shadow.blur_radius = Emu(50800)
        outer = shadow._element.effectLst.outerShdw
        assert outer is not None
        assert outer.blurRad == Emu(50800)

    def it_reads_and_writes_distance(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst"))
        assert shadow.distance is None
        shadow.distance = Emu(38100)
        assert shadow.distance == Emu(38100)
        assert shadow._element.effectLst.outerShdw.dist == Emu(38100)

    def it_reads_and_writes_direction(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst"))
        assert shadow.direction is None
        shadow.direction = 270.0
        outer = shadow._element.effectLst.outerShdw
        assert outer.dir == 270.0

    def it_materializes_a_color_on_first_access(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst"))
        color = shadow.color
        assert isinstance(color, ColorFormat)
        outer = shadow._element.effectLst.outerShdw
        assert outer is not None
        assert outer.eg_colorChoice is not None

    def and_it_returns_the_same_ColorFormat_on_subsequent_access(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst"))
        assert shadow.color is shadow.color

    def it_leaves_xml_valid_after_round_trip(self):
        # -- start/expect XML that already declares `a:` so we can compare strings --
        shadow = ShadowFormat(element("p:spPr/a:effectLst/a:outerShdw"))
        shadow.blur_radius = Emu(50800)
        shadow.distance = Emu(38100)
        shadow.direction = 90.0
        expected = xml("p:spPr/a:effectLst/a:outerShdw{blurRad=50800,dist=38100,dir=5400000}")
        assert shadow._element.xml == expected

    # -- issue #446: inherit=False must also zero sibling `a:effectRef/@idx` -------

    def it_zeroes_sibling_effectRef_idx_when_breaking_inheritance(self):
        """Issue #446: setting `inherit = False` must also zero `p:style/a:effectRef/@idx`
        so a theme-inherited shadow doesn't leak through the empty `a:effectLst`."""
        sp = _sp_with_style_effectRef(idx="2")
        spPr = sp.xpath("./p:spPr")[0]
        shadow = ShadowFormat(spPr)

        shadow.inherit = False

        effectRef = sp.xpath("./p:style/a:effectRef")[0]
        assert effectRef.get("idx") == "0"
        assert spPr.xpath("./a:effectLst")  # -- effectLst also added

    def but_it_does_not_create_a_style_or_effectRef_when_none_exists(self):
        """No `p:style` on the parent → setting inherit=False only adds `a:effectLst`."""
        sp = _sp_without_style()
        spPr = sp.xpath("./p:spPr")[0]
        shadow = ShadowFormat(spPr)

        shadow.inherit = False

        assert not sp.xpath("./p:style")
        assert spPr.xpath("./a:effectLst")

    def and_it_is_a_no_op_on_a_parentless_spPr(self):
        """Detached `p:spPr` with no parent must not raise."""
        shadow = ShadowFormat(element("p:spPr"))
        # -- must not raise even though `self._element.getparent()` is None --
        shadow.inherit = False
        assert shadow._element.effectLst is not None

    def and_it_works_for_a_cxnSp_parent_too(self):
        """Connectors (`p:cxnSp`) also have a sibling `p:style/a:effectRef`."""
        cxnSp = _cxnSp_with_style_effectRef(idx="1")
        spPr = cxnSp.xpath("./p:spPr")[0]
        shadow = ShadowFormat(spPr)

        shadow.inherit = False

        effectRef = cxnSp.xpath("./p:style/a:effectRef")[0]
        assert effectRef.get("idx") == "0"

    def and_it_leaves_effectRef_untouched_when_restoring_inheritance(self):
        """`inherit = True` removes `a:effectLst` but must not touch `a:effectRef`."""
        sp = _sp_with_style_effectRef(idx="2")
        spPr = sp.xpath("./p:spPr")[0]
        # -- seed an explicit effectLst so that the True branch has something to remove --
        shadow = ShadowFormat(spPr)
        shadow.inherit = False  # -- also zeroes effectRef --
        # -- manually restore a non-zero idx to simulate post-edit state --
        sp.xpath("./p:style/a:effectRef")[0].set("idx", "2")

        shadow.inherit = True

        assert spPr.find(qn("a:effectLst")) is None
        # -- effectRef idx must still be "2", the True branch doesn't touch it --
        assert sp.xpath("./p:style/a:effectRef")[0].get("idx") == "2"

    def and_it_leaves_grpSpPr_behavior_unchanged(self):
        """Group shapes have no `p:style` sibling, so fix must be a silent no-op."""
        grpSp = _grpSp_without_style()
        grpSpPr = grpSp.xpath("./p:grpSpPr")[0]
        shadow = ShadowFormat(grpSpPr)

        shadow.inherit = False

        assert grpSpPr.xpath("./a:effectLst")

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("p:spPr", True),
            ("p:spPr/a:effectLst", False),
            ("p:grpSpPr", True),
            ("p:grpSpPr/a:effectLst", False),
        ]
    )
    def inherit_get_fixture(self, request):
        cxml, expected_value = request.param
        shadow = ShadowFormat(element(cxml))
        return shadow, expected_value

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", False, "p:spPr{a:b=c}/a:effectLst"),
            ("p:grpSpPr{a:b=c}", False, "p:grpSpPr{a:b=c}/a:effectLst"),
            ("p:spPr{a:b=c}/a:effectLst", True, "p:spPr{a:b=c}"),
            ("p:grpSpPr{a:b=c}/a:effectLst", True, "p:grpSpPr{a:b=c}"),
            ("p:spPr", True, "p:spPr"),
            ("p:grpSpPr", True, "p:grpSpPr"),
            ("p:spPr/a:effectLst", False, "p:spPr/a:effectLst"),
            ("p:grpSpPr/a:effectLst", False, "p:grpSpPr/a:effectLst"),
        ]
    )
    def inherit_set_fixture(self, request):
        cxml, value, expected_cxml = request.param
        shadow = ShadowFormat(element(cxml))
        expected_value = xml(expected_cxml)
        return shadow, value, expected_value


class DescribeEffectFormat(object):
    def it_reports_inherit_True_when_effectLst_is_absent(self):
        effect = EffectFormat(element("p:spPr"))
        assert effect.inherit is True

    def it_reports_inherit_False_when_effectLst_is_present(self):
        effect = EffectFormat(element("p:spPr/a:effectLst"))
        assert effect.inherit is False

    def it_can_break_inheritance_by_adding_an_empty_effectLst(self):
        effect = EffectFormat(element("p:spPr{a:b=c}"))
        effect.inherit = False
        assert effect._element.effectLst is not None

    def it_can_restore_inheritance_by_removing_effectLst(self):
        effect = EffectFormat(element("p:spPr/a:effectLst/a:glow"))
        effect.inherit = True
        assert effect._element.effectLst is None

    def it_exposes_shadow_glow_reflection_and_soft_edge(self):
        effect = EffectFormat(element("p:spPr"))
        assert isinstance(effect.shadow, ShadowFormat)
        assert isinstance(effect.glow, GlowFormat)
        assert isinstance(effect.reflection, ReflectionFormat)
        assert isinstance(effect.soft_edge, SoftEdgeFormat)

    def it_memoizes_each_sub_format(self):
        effect = EffectFormat(element("p:spPr"))
        assert effect.shadow is effect.shadow
        assert effect.glow is effect.glow
        assert effect.reflection is effect.reflection
        assert effect.soft_edge is effect.soft_edge


class DescribeGlowFormat(object):
    def it_reports_inherit_True_when_glow_is_absent(self, inherit_glow_fixture):
        glow, expected = inherit_glow_fixture
        assert glow.inherit is expected

    @pytest.fixture(
        params=[
            ("p:spPr", True),
            ("p:spPr/a:effectLst", True),
            ("p:spPr/a:effectLst/a:glow", False),
        ]
    )
    def inherit_glow_fixture(self, request):
        cxml, expected = request.param
        return GlowFormat(element(cxml)), expected

    def it_adds_glow_and_effectLst_when_clearing_inherit(self):
        glow = GlowFormat(element("p:spPr{a:b=c}"))
        glow.inherit = False
        assert glow._element.effectLst is not None
        assert glow._element.effectLst.glow is not None

    def it_removes_glow_but_not_effectLst_when_restoring_inherit(self):
        glow = GlowFormat(element("p:spPr/a:effectLst/a:glow{rad=50800}"))
        glow.inherit = True
        assert glow._element.effectLst is not None
        assert glow._element.effectLst.glow is None

    def it_reports_size_as_None_when_glow_absent(self):
        glow = GlowFormat(element("p:spPr"))
        assert glow.size is None

    def it_reads_and_writes_size(self):
        glow = GlowFormat(element("p:spPr/a:effectLst"))
        glow.size = Emu(50800)
        assert glow.size == Emu(50800)
        assert glow._element.effectLst.glow.rad == Emu(50800)

    def it_clears_size_to_zero_when_assigned_None(self):
        glow = GlowFormat(element("p:spPr/a:effectLst/a:glow{rad=50800}"))
        glow.size = None
        assert glow.size == Emu(0)

    def it_materializes_glow_color_on_first_access(self):
        glow = GlowFormat(element("p:spPr/a:effectLst"))
        color = glow.color
        assert isinstance(color, ColorFormat)
        assert glow._element.effectLst.glow is not None
        assert glow._element.effectLst.glow.eg_colorChoice is not None

    def it_produces_expected_xml_for_populated_glow(self):
        glow = GlowFormat(element("p:spPr/a:effectLst/a:glow"))
        glow.size = Emu(50800)
        expected = xml("p:spPr/a:effectLst/a:glow{rad=50800}")
        assert glow._element.xml == expected


class DescribeReflectionFormat(object):
    def it_reports_inherit(self):
        r = ReflectionFormat(element("p:spPr"))
        assert r.inherit is True
        r = ReflectionFormat(element("p:spPr/a:effectLst/a:reflection"))
        assert r.inherit is False

    def it_reads_blur_radius_and_distance_as_None_when_absent(self):
        r = ReflectionFormat(element("p:spPr"))
        assert r.blur_radius is None
        assert r.distance is None

    def it_writes_blur_radius_and_distance(self):
        r = ReflectionFormat(element("p:spPr/a:effectLst"))
        r.blur_radius = Emu(25400)
        r.distance = Emu(50800)
        refl = r._element.effectLst.reflection
        assert refl.blurRad == Emu(25400)
        assert refl.dist == Emu(50800)


class DescribeSoftEdgeFormat(object):
    def it_reports_inherit(self):
        s = SoftEdgeFormat(element("p:spPr"))
        assert s.inherit is True
        s = SoftEdgeFormat(element("p:spPr/a:effectLst/a:softEdge{rad=12700}"))
        assert s.inherit is False

    def it_reads_size_as_None_when_absent(self):
        s = SoftEdgeFormat(element("p:spPr"))
        assert s.size is None

    def it_writes_size(self):
        s = SoftEdgeFormat(element("p:spPr/a:effectLst"))
        s.size = Emu(12700)
        assert s.size == Emu(12700)
        assert s._element.effectLst.softEdge.rad == Emu(12700)

    def it_rejects_None_assignment_since_rad_is_required(self):
        s = SoftEdgeFormat(element("p:spPr/a:effectLst/a:softEdge{rad=12700}"))
        with pytest.raises(ValueError):
            s.size = None


# -- helpers -----------------------------------------------------------------


def _sp_with_style_effectRef(idx: str):
    """Return a `p:sp` element with `p:spPr` and a sibling `p:style/a:effectRef@idx`."""
    return parse_xml(
        f"<p:sp {nsdecls('a', 'p')}>\n"
        f"  <p:spPr/>\n"
        f"  <p:style>\n"
        f'    <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>\n'
        f'    <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>\n'
        f'    <a:effectRef idx="{idx}"><a:schemeClr val="accent1"/></a:effectRef>\n'
        f'    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>\n'
        f"  </p:style>\n"
        f"</p:sp>"
    )


def _sp_without_style():
    """Return a `p:sp` element with a `p:spPr` child but no sibling `p:style`."""
    return parse_xml(f"<p:sp {nsdecls('a', 'p')}>\n  <p:spPr/>\n</p:sp>")


def _cxnSp_with_style_effectRef(idx: str):
    """Return a `p:cxnSp` connector with `p:spPr` and sibling `p:style/a:effectRef@idx`."""
    return parse_xml(
        f"<p:cxnSp {nsdecls('a', 'p')}>\n"
        f"  <p:spPr/>\n"
        f"  <p:style>\n"
        f'    <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>\n'
        f'    <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>\n'
        f'    <a:effectRef idx="{idx}"><a:schemeClr val="accent1"/></a:effectRef>\n'
        f'    <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>\n'
        f"  </p:style>\n"
        f"</p:cxnSp>"
    )


def _grpSp_without_style():
    """Return a `p:grpSp` with a `p:grpSpPr` child (group shapes have no `p:style`)."""
    return parse_xml(
        f"<p:grpSp {nsdecls('a', 'p')}>\n  <p:grpSpPr/>\n</p:grpSp>"
    )
