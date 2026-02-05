"""Custom element classes for OMML (Office Math Markup Language) elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    pass


class CT_OMath(BaseOxmlElement):
    """`m:oMath` custom element class - root element for a math equation."""

    # Type hints for public API methods that will be created by manual implementations below
    # These match the pattern used by other python-pptx classes like text and table elements
    add_r: Callable[[], "CT_R"]        # Add a math run element
    add_f: Callable[[], "CT_F"]        # Add a fraction element
    add_sSup: Callable[[], "CT_SSup"]  # Add a superscript element
    add_rad: Callable[[], "CT_Rad"]    # Add a radical element
    add_nary: Callable[[], "CT_Nary"]  # Add an n-ary operator element

    # Declarative element definitions that create internal _add_* methods via metaclass
    # ZeroOrMore allows multiple instances and creates _add_r, _add_f, etc. methods
    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")
    rad: "ZeroOrMore" = ZeroOrMore("m:rad")
    nary: "ZeroOrMore" = ZeroOrMore("m:nary")

    # Manual implementations of public API methods that delegate to internal metaclass methods
    # This pattern matches how text and table classes handle ZeroOrMore/ZeroOrOne elements
    def add_r(self) -> "CT_R":
        """Add a math run element."""
        return self._add_r()

    def add_f(self) -> "CT_F":
        """Add a fraction element."""
        return self._add_f()

    def add_sSup(self) -> "CT_SSup":
        """Add a superscript element."""
        return self._add_sSup()

    def add_rad(self) -> "CT_Rad":
        """Add a radical element."""
        return self._add_rad()

    def add_nary(self) -> "CT_Nary":
        """Add an n-ary operator element."""
        return self._add_nary()


class CT_R(BaseOxmlElement):
    """`m:r` custom element class - math run (text container)."""

    add_t: Callable[[], "CT_T"]
    add_rPr: Callable[[], "CT_RPr"]

    t: "ZeroOrOne" = ZeroOrOne("m:t")
    rPr: "ZeroOrOne" = ZeroOrOne("m:rPr")

    def add_t(self) -> "CT_T":
        """Add a text element."""
        return self._add_t()

    def add_rPr(self) -> "CT_RPr":
        """Add run properties."""
        return self._add_rPr()


class CT_T(BaseOxmlElement):
    """`m:t` custom element class - text content."""

    @property
    def content(self) -> str:
        """Get text content."""
        return self.text or ""

    @content.setter
    def content(self, value: str):
        """Set text content."""
        # Use lxml's built-in text assignment
        object.__setattr__(self, 'text', value)


class CT_RPr(BaseOxmlElement):
    """`m:rPr` custom element class - run properties."""
    # Math run properties would be defined here
    pass


class CT_F(BaseOxmlElement):
    """`m:f` custom element class - fraction."""

    add_num: Callable[[], "CT_Num"]
    add_den: Callable[[], "CT_Den"]

    num: "ZeroOrOne" = ZeroOrOne("m:num")
    den: "ZeroOrOne" = ZeroOrOne("m:den")

    def add_num(self) -> "CT_Num":
        """Add numerator element."""
        return self._add_num()

    def add_den(self) -> "CT_Den":
        """Add denominator element."""
        return self._add_den()


class CT_Num(BaseOxmlElement):
    """`m:num` custom element class - fraction numerator."""

    add_r: Callable[[], "CT_R"]
    add_f: Callable[[], "CT_F"]
    add_sSup: Callable[[], "CT_SSup"]
    add_rad: Callable[[], "CT_Rad"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")
    rad: "ZeroOrMore" = ZeroOrMore("m:rad")

    def add_r(self) -> "CT_R":
        """Add a math run element."""
        return self._add_r()

    def add_f(self) -> "CT_F":
        """Add a fraction element."""
        return self._add_f()

    def add_sSup(self) -> "CT_SSup":
        """Add a superscript element."""
        return self._add_sSup()

    def add_rad(self) -> "CT_Rad":
        """Add a radical element."""
        return self._add_rad()


class CT_Den(BaseOxmlElement):
    """`m:den` custom element class - fraction denominator."""

    add_r: Callable[[], "CT_R"]
    add_f: Callable[[], "CT_F"]
    add_sSup: Callable[[], "CT_SSup"]
    add_rad: Callable[[], "CT_Rad"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")
    rad: "ZeroOrMore" = ZeroOrMore("m:rad")

    def add_r(self) -> "CT_R":
        """Add a math run element."""
        return self._add_r()

    def add_f(self) -> "CT_F":
        """Add a fraction element."""
        return self._add_f()

    def add_sSup(self) -> "CT_SSup":
        """Add a superscript element."""
        return self._add_sSup()

    def add_rad(self) -> "CT_Rad":
        """Add a radical element."""
        return self._add_rad()


class CT_SSup(BaseOxmlElement):
    """`m:sSup` custom element class - superscript."""

    add_e: Callable[[], "CT_E"]

    e: "ZeroOrOne" = ZeroOrOne("m:e")

    def add_e(self) -> "CT_E":
        """Add element."""
        return self._add_e()


class CT_E(BaseOxmlElement):
    """`m:e` custom element class - expression (base of superscript/subscript)."""

    add_r: Callable[[], "CT_R"]
    add_f: Callable[[], "CT_F"]
    add_sSup: Callable[[], "CT_SSup"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")

    def add_r(self) -> "CT_R":
        """Add a math run element."""
        return self._add_r()

    def add_f(self) -> "CT_F":
        """Add a fraction element."""
        return self._add_f()

    def add_sSup(self) -> "CT_SSup":
        """Add a superscript element."""
        return self._add_sSup()


class CT_Rad(BaseOxmlElement):
    """`m:rad` custom element class - radical (square root)."""

    add_radPr: Callable[[], "CT_RadPr"]
    add_deg: Callable[[], "CT_Deg"]
    add_e: Callable[[], "CT_E"]

    radPr: "ZeroOrOne" = ZeroOrOne("m:radPr")
    deg: "ZeroOrOne" = ZeroOrOne("m:deg")
    e: "ZeroOrOne" = ZeroOrOne("m:e")

    def add_radPr(self) -> "CT_RadPr":
        """Add radical properties."""
        return self._add_radPr()

    def add_deg(self) -> "CT_Deg":
        """Add degree."""
        return self._add_deg()

    def add_e(self) -> "CT_E":
        """Add element."""
        return self._add_e()


class CT_RadPr(BaseOxmlElement):
    """`m:radPr` custom element class - radical properties."""
    # Radical properties would be defined here
    pass


class CT_Deg(BaseOxmlElement):
    """`m:deg` custom element class - radical degree (for nth roots)."""

    add_r: Callable[[], "CT_R"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")

    def add_r(self) -> "CT_R":
        """Add a math run element."""
        return self._add_r()


class CT_Nary(BaseOxmlElement):
    """`m:nary` custom element class - n-ary operators (summation, integral, etc.)."""

    add_naryPr: Callable[[], "CT_NaryPr"]
    add_sub: Callable[[], "CT_Sub"]
    add_sup: Callable[[], "CT_Sup"]
    add_e: Callable[[], "CT_E"]

    naryPr: "ZeroOrOne" = ZeroOrOne("m:naryPr")
    sub: "ZeroOrOne" = ZeroOrOne("m:sub")
    sup: "ZeroOrOne" = ZeroOrOne("m:sup")
    e: "ZeroOrOne" = ZeroOrOne("m:e")

    def add_naryPr(self) -> "CT_NaryPr":
        """Add n-ary operator properties."""
        return self._add_naryPr()

    def add_sub(self) -> "CT_Sub":
        """Add subscript."""
        return self._add_sub()

    def add_sup(self) -> "CT_Sup":
        """Add superscript."""
        return self._add_sup()

    def add_e(self) -> "CT_E":
        """Add element."""
        return self._add_e()


class CT_NaryPr(BaseOxmlElement):
    """`m:naryPr` custom element class - n-ary operator properties."""

    add_chr: Callable[[], "CT_Char"]
    add_limLoc: Callable[[], "CT_LimLoc"]

    chr: "ZeroOrOne" = ZeroOrOne("m:chr")
    limLoc: "ZeroOrOne" = ZeroOrOne("m:limLoc")

    def add_chr(self) -> "CT_Char":
        """Add character element."""
        return self._add_chr()

    def add_limLoc(self) -> "CT_LimLoc":
        """Add limit location."""
        return self._add_limLoc()


class CT_Char(BaseOxmlElement):
    """`m:chr` custom element class - character for n-ary operators."""
    pass


class CT_LimLoc(BaseOxmlElement):
    """`m:limLoc` custom element class - limit location for n-ary operators."""
    pass


class CT_Sub(BaseOxmlElement):
    """`m:sub` custom element class - subscript for n-ary operators."""

    add_r: Callable[[], "CT_R"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")

    def add_r(self) -> "CT_R":
        """Add a math run element."""
        return self._add_r()


class CT_Sup(BaseOxmlElement):
    """`m:sup` custom element class - superscript for n-ary operators."""

    add_r: Callable[[], "CT_R"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")

    def add_r(self) -> "CT_R":
        """Add a math run element."""
        return self._add_r()
