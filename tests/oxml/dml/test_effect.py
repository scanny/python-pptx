"""Unit-test suite for `pptx.oxml.dml.effect` module."""

from __future__ import annotations

import pytest

from pptx.oxml.dml.effect import (
    CT_EffectList,
    CT_GlowEffect,
    CT_InnerShadowEffect,
    CT_OuterShadowEffect,
    CT_ReflectionEffect,
    CT_SoftEdgesEffect,
)
from pptx.util import Emu

from ...unitutil.cxml import element


class DescribeCT_EffectList(object):
    def it_can_construct_a_loose_effectLst(self):
        effectLst = CT_EffectList.new_effectLst()
        assert effectLst.tag.endswith("}effectLst")
        assert effectLst.blur is None
        assert effectLst.glow is None
        assert effectLst.innerShdw is None
        assert effectLst.outerShdw is None
        assert effectLst.prstShdw is None
        assert effectLst.reflection is None
        assert effectLst.softEdge is None

    def it_can_access_each_child_by_descriptor(self):
        effectLst = element("a:effectLst/(a:glow,a:innerShdw,a:outerShdw,a:reflection,a:softEdge)")
        assert isinstance(effectLst, CT_EffectList)
        assert isinstance(effectLst.glow, CT_GlowEffect)
        assert isinstance(effectLst.innerShdw, CT_InnerShadowEffect)
        assert isinstance(effectLst.outerShdw, CT_OuterShadowEffect)
        assert isinstance(effectLst.reflection, CT_ReflectionEffect)
        assert isinstance(effectLst.softEdge, CT_SoftEdgesEffect)

    def it_can_lazily_create_a_glow_child(self):
        effectLst = CT_EffectList.new_effectLst()
        glow = effectLst.get_or_add_glow()
        assert isinstance(glow, CT_GlowEffect)
        # subsequent call is idempotent
        assert effectLst.get_or_add_glow() is glow


class DescribeCT_GlowEffect(object):
    def it_reads_rad_as_zero_when_absent(self):
        glow = element("a:glow")
        assert glow.rad == Emu(0)

    def it_reads_rad_from_xml(self):
        glow = element("a:glow{rad=38100}")
        assert glow.rad == Emu(38100)


class DescribeCT_OuterShadowEffect(object):
    def it_reads_default_attrs_as_zero(self):
        outer = element("a:outerShdw")
        assert outer.blurRad == Emu(0)
        assert outer.dist == Emu(0)
        assert outer.dir == 0.0
        assert outer.rotWithShape is True

    def it_reads_explicit_attrs(self):
        outer = element("a:outerShdw{blurRad=50800,dist=38100,dir=5400000,rotWithShape=0}")
        assert outer.blurRad == Emu(50800)
        assert outer.dist == Emu(38100)
        assert outer.dir == 90.0
        assert outer.rotWithShape is False


class DescribeCT_InnerShadowEffect(object):
    def it_reads_default_attrs(self):
        inner = element("a:innerShdw")
        assert inner.blurRad == Emu(0)
        assert inner.dist == Emu(0)
        assert inner.dir == 0.0


class DescribeCT_PresetShadowEffect(object):
    def it_accepts_valid_preset_values(self):
        prst = element("a:prstShdw{prst=shdw5}")
        assert prst.prst == "shdw5"

    def it_rejects_invalid_preset_values(self):
        prst = element("a:prstShdw{prst=shdw5}")
        with pytest.raises(ValueError):
            prst.prst = "not-a-preset"


class DescribeCT_ReflectionEffect(object):
    def it_reads_defaults(self):
        refl = element("a:reflection")
        assert refl.blurRad == Emu(0)
        assert refl.dist == Emu(0)
        assert refl.dir == 0.0
        assert refl.sx == 1.0
        assert refl.sy == 1.0
        assert refl.rotWithShape is True


class DescribeCT_SoftEdgesEffect(object):
    def it_requires_rad(self):
        soft = element("a:softEdge{rad=12700}")
        assert soft.rad == Emu(12700)


class DescribeCT_BlurEffect(object):
    def it_reads_default_attrs(self):
        blur = element("a:blur")
        assert blur.rad == Emu(0)
        assert blur.grow is True
