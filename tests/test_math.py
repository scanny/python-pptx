"""Unit tests for math functionality."""

from __future__ import annotations

import pytest

from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.math import (
    CT_OMath, CT_R, CT_T, CT_F, CT_Num, CT_Den, CT_SSup, CT_E, CT_Rad,
    CT_Nary, CT_Char, CT_LimLoc, CT_Sub, CT_Sup
)


class TestCT_OMath:
    """Test CT_OMath element class."""

    def test_ct_omml_creation(self):
        """Test CT_OMath element creation."""
        omath = OxmlElement("m:oMath")
        assert isinstance(omath, CT_OMath)
        # The tag includes the full namespace URL
        assert "oMath" in omath.tag
        assert "http://schemas.openxmlformats.org/officeDocument/2006/math" in omath.tag

    def test_ct_omml_add_r(self):
        """Test adding text run to OMML element."""
        omath = OxmlElement("m:oMath")
        r = omath.add_r()
        assert isinstance(r, CT_R)
        assert r in omath

    def test_ct_omml_add_f(self):
        """Test adding fraction to OMML element."""
        omath = OxmlElement("m:oMath")
        f = omath.add_f()
        assert isinstance(f, CT_F)
        assert f in omath

    def test_ct_omml_add_sSup(self):
        """Test adding superscript to OMML element."""
        omath = OxmlElement("m:oMath")
        sup = omath.add_sSup()
        assert isinstance(sup, CT_SSup)
        assert sup in omath

    def test_ct_omml_add_rad(self):
        """Test adding radical to OMML element."""
        omath = OxmlElement("m:oMath")
        rad = omath.add_rad()
        assert isinstance(rad, CT_Rad)
        assert rad in omath

    def test_ct_omml_add_nary(self):
        """Test adding n-ary operator to OMML element."""
        omath = OxmlElement("m:oMath")
        nary = omath.add_nary()
        assert isinstance(nary, CT_Nary)
        assert nary in omath


class TestCT_R:
    """Test CT_R text run element class."""

    def test_ct_r_creation(self):
        """Test CT_R element creation."""
        r = OxmlElement("m:r")
        assert isinstance(r, CT_R)
        # Check for the namespace URL in the tag
        assert "r" in r.tag
        assert "http://schemas.openxmlformats.org/officeDocument/2006/math" in r.tag

    def test_ct_r_add_t(self):
        """Test adding text to run."""
        r = OxmlElement("m:r")
        t = r.add_t()
        assert isinstance(t, CT_T)
        assert t in r

    def test_ct_r_add_rPr(self):
        """Test adding run properties."""
        r = OxmlElement("m:r")
        rPr = r.add_rPr()
        assert rPr is not None


class TestCT_T:
    """Test CT_T text element class."""

    def test_ct_t_creation(self):
        """Test CT_T element creation."""
        t = OxmlElement("m:t")
        assert isinstance(t, CT_T)
        # Check for the namespace URL in the tag
        assert "t" in t.tag
        assert "http://schemas.openxmlformats.org/officeDocument/2006/math" in t.tag

    def test_ct_t_text_property(self):
        """Test text property getter and setter."""
        t = OxmlElement("m:t")

        # Test default
        assert t.content == ""

        # Test setter
        t.content = "Hello Math"
        assert t.content == "Hello Math"


class TestCT_F:
    """Test CT_F fraction element class."""

    def test_ct_f_creation(self):
        """Test CT_F element creation."""
        f = OxmlElement("m:f")
        assert isinstance(f, CT_F)
        # Check for the namespace URL in the tag
        assert "f" in f.tag
        assert "http://schemas.openxmlformats.org/officeDocument/2006/math" in f.tag

    def test_ct_f_add_num(self):
        """Test adding numerator to fraction."""
        f = OxmlElement("m:f")
        num = f.add_num()
        assert isinstance(num, CT_Num)
        assert num in f

    def test_ct_f_add_den(self):
        """Test adding denominator to fraction."""
        f = OxmlElement("m:f")
        den = f.add_den()
        assert isinstance(den, CT_Den)
        assert den in f


class TestCT_Num:
    """Test CT_Num numerator element class."""

    def test_ct_num_creation(self):
        """Test CT_Num element creation."""
        num = OxmlElement("m:num")
        assert isinstance(num, CT_Num)
        # Check for the namespace URL in the tag
        assert "num" in num.tag
        assert "http://schemas.openxmlformats.org/officeDocument/2006/math" in num.tag

    def test_ct_num_add_r(self):
        """Test adding run to numerator."""
        num = OxmlElement("m:num")
        r = num.add_r()
        assert isinstance(r, CT_R)
        assert r in num


class TestCT_Den:
    """Test CT_Den denominator element class."""

    def test_ct_den_creation(self):
        """Test CT_Den element creation."""
        den = OxmlElement("m:den")
        assert isinstance(den, CT_Den)
        # Check for the namespace URL in the tag
        assert "den" in den.tag
        assert "http://schemas.openxmlformats.org/officeDocument/2006/math" in den.tag

    def test_ct_den_add_r(self):
        """Test adding run to denominator."""
        den = OxmlElement("m:den")
        r = den.add_r()
        assert isinstance(r, CT_R)
        assert r in den


class TestOMMLIntegration:
    """Test OMML integration and XML generation."""

    def test_simple_fraction_xml(self):
        """Test generating XML for simple fraction."""
        omath = OxmlElement("m:oMath")

        f = omath.add_f()
        num = f.add_num()
        r_num = num.add_r()
        t_num = r_num.add_t()
        t_num.text = "x"

        den = f.add_den()
        r_den = den.add_r()
        t_den = r_den.add_t()
        t_den.text = "2"

        from pptx.oxml.xmlchemy import serialize_for_reading
        xml = serialize_for_reading(omath)

        assert "m:oMath" in xml
        assert "m:f" in xml
        assert "m:num" in xml
        assert "m:den" in xml
        assert "x" in xml
        assert "2" in xml

    def test_superscript_xml(self):
        """Test generating XML for superscript."""
        omath = OxmlElement("m:oMath")

        # Base x
        r_base = omath.add_r()
        t_base = r_base.add_t()
        t_base.text = "x"

        # Superscript 2
        sup = omath.add_sSup()
        e = sup.add_e()
        r_sup = e.add_r()
        t_sup = r_sup.add_t()
        t_sup.text = "2"

        from pptx.oxml.xmlchemy import serialize_for_reading
        xml = serialize_for_reading(omath)

        assert "m:oMath" in xml
        assert "m:sSup" in xml
        assert "x" in xml
        assert "2" in xml

    def test_radical_xml(self):
        """Test generating XML for radical."""
        omath = OxmlElement("m:oMath")

        rad = omath.add_rad()
        radPr = rad.add_radPr()
        deg = rad.add_deg()
        e = rad.add_e()
        r = e.add_r()
        t = r.add_t()
        t.text = "x"

        from pptx.oxml.xmlchemy import serialize_for_reading
        xml = serialize_for_reading(omath)

        assert "m:oMath" in xml
        assert "m:rad" in xml
        assert "m:radPr" in xml
        assert "m:deg" in xml
        assert "x" in xml

    def test_nary_xml(self):
        """Test generating XML for n-ary operator."""
        omath = OxmlElement("m:oMath")

        nary = omath.add_nary()
        naryPr = nary.add_naryPr()
        chr_elem = naryPr.add_chr()
        chr_elem.set("val", "∑")

        sub = nary.add_sub()
        r_sub = sub.add_r()
        t_sub = r_sub.add_t()
        t_sub.text = "i=1"

        sup = nary.add_sup()
        r_sup = sup.add_r()
        t_sup = r_sup.add_t()
        t_sup.text = "n"

        e = nary.add_e()
        r_e = e.add_r()
        t_e = r_e.add_t()
        t_e.text = "i"

        from pptx.oxml.xmlchemy import serialize_for_reading
        xml = serialize_for_reading(omath)

        assert "m:oMath" in xml
        assert "m:nary" in xml
        assert "∑" in xml
        assert "i=1" in xml
        assert "n" in xml
        assert "i" in xml
