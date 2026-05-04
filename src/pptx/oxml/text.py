"""Custom element classes for text-related XML elements"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Callable, cast

from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import (
    MSO_AUTO_SIZE,
    MSO_TEXT_STRIKE_TYPE,
    MSO_TEXT_UNDERLINE_TYPE,
    MSO_VERTICAL_ANCHOR,
    PP_AUTO_NUMBER_SCHEME,
    PP_PARAGRAPH_ALIGNMENT,
)
from pptx.exc import InvalidXmlError
from pptx.oxml import parse_xml
from pptx.oxml.dml.fill import CT_GradientFillProperties
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_Angle,
    ST_Coordinate32,
    ST_TextBaselinePercent,
    ST_TextBulletSizePercent,
    ST_TextBulletStartAtNum,
    ST_TextFontScalePercentOrPercentString,
    ST_TextFontScaleReductionPercent,
    ST_TextFontSize,
    ST_TextIndentLevelType,
    ST_TextSpacingPercentOrPercentString,
    ST_TextSpacingPoint,
    ST_TextTypeface,
    ST_TextWrappingType,
    XsdBoolean,
    XsdString,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OneAndOnlyOne,
    OneOrMore,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
    ZeroOrOneChoice,
)
from pptx.util import Centipoints, Emu, Length

if TYPE_CHECKING:
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.dml.color import CT_Color
    from pptx.oxml.dml.effect import CT_EffectList


class CT_RegularTextRun(BaseOxmlElement):
    """`a:r` custom element class"""

    get_or_add_rPr: Callable[[], CT_TextCharacterProperties]

    rPr: CT_TextCharacterProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:rPr", successors=("a:t",)
    )
    t: BaseOxmlElement = OneAndOnlyOne("a:t")  # pyright: ignore[reportAssignmentType]

    @property
    def text(self) -> str:
        """All text of (required) `a:t` child."""
        text = self.t.text
        # -- t.text is None when t element is empty, e.g. '<a:t/>' --
        return text or ""

    @text.setter
    def text(self, value: str):  # pyright: ignore[reportIncompatibleMethodOverride]
        self.t.text = self._escape_ctrl_chars(value)

    @staticmethod
    def _escape_ctrl_chars(s: str) -> str:
        """Return str after replacing each control character with a plain-text escape.

        For example, a BEL character (x07) would appear as "_x0007_". Horizontal-tab
        (x09) and line-feed (x0A) are not escaped. All other characters in the range
        x00-x1F are escaped.
        """
        return re.sub(r"([\x00-\x08\x0B-\x1F])", lambda match: "_x%04X_" % ord(match.group(1)), s)


class CT_TextBody(BaseOxmlElement):
    """`p:txBody` custom element class.

    Also used for `c:txPr` in charts and perhaps other elements.
    """

    add_p: Callable[[], CT_TextParagraph]
    p_lst: list[CT_TextParagraph]

    bodyPr: CT_TextBodyProperties = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:bodyPr"
    )
    p: CT_TextParagraph = OneOrMore("a:p")  # pyright: ignore[reportAssignmentType]

    def clear_content(self):
        """Remove all `a:p` children, but leave any others.

        cf. lxml `_Element.clear()` method which removes all children.
        """
        for p in self.p_lst:
            self.remove(p)

    @property
    def defRPr(self) -> CT_TextCharacterProperties:
        """`a:defRPr` element of required first `p` child, added with its ancestors if not present.

        Used when element is a ``c:txPr`` in a chart and the `p` element is used only to specify
        formatting, not content.
        """
        p = self.p_lst[0]
        pPr = p.get_or_add_pPr()
        defRPr = pPr.get_or_add_defRPr()
        return defRPr

    @property
    def is_empty(self) -> bool:
        """True if only a single empty `a:p` element is present."""
        ps = self.p_lst
        if len(ps) > 1:
            return False

        if not ps:
            raise InvalidXmlError("p:txBody must have at least one a:p")

        return ps[0].text == ""

    @classmethod
    def new(cls):
        """Return a new `p:txBody` element tree."""
        xml = cls._txBody_tmpl()
        txBody = parse_xml(xml)
        return txBody

    @classmethod
    def new_a_txBody(cls) -> CT_TextBody:
        """Return a new `a:txBody` element tree.

        Suitable for use in a table cell and possibly other situations.
        """
        xml = cls._a_txBody_tmpl()
        txBody = cast(CT_TextBody, parse_xml(xml))
        return txBody

    @classmethod
    def new_p_txBody(cls):
        """Return a new `p:txBody` element tree, suitable for use in an `p:sp` element."""
        xml = cls._p_txBody_tmpl()
        return parse_xml(xml)

    @classmethod
    def new_txPr(cls):
        """Return a `c:txPr` element tree.

        Suitable for use in a chart object like data labels or tick labels.
        """
        xml = (
            "<c:txPr %s>\n"
            "  <a:bodyPr/>\n"
            "  <a:lstStyle/>\n"
            "  <a:p>\n"
            "    <a:pPr>\n"
            "      <a:defRPr/>\n"
            "    </a:pPr>\n"
            "  </a:p>\n"
            "</c:txPr>\n"
        ) % nsdecls("c", "a")
        txPr = parse_xml(xml)
        return txPr

    def unclear_content(self):
        """Ensure p:txBody has at least one a:p child.

        Intuitively, reverse a ".clear_content()" operation to minimum conformance with spec
        (single empty paragraph).
        """
        if len(self.p_lst) > 0:
            return
        self.add_p()

    @classmethod
    def _a_txBody_tmpl(cls):
        return "<a:txBody %s>\n" "  <a:bodyPr/>\n" "  <a:p/>\n" "</a:txBody>\n" % (nsdecls("a"))

    @classmethod
    def _p_txBody_tmpl(cls):
        return (
            "<p:txBody %s>\n" "  <a:bodyPr/>\n" "  <a:p/>\n" "</p:txBody>\n" % (nsdecls("p", "a"))
        )

    @classmethod
    def _txBody_tmpl(cls):
        return (
            "<p:txBody %s>\n"
            "  <a:bodyPr/>\n"
            "  <a:lstStyle/>\n"
            "  <a:p/>\n"
            "</p:txBody>\n" % (nsdecls("a", "p"))
        )


class CT_TextBodyProperties(BaseOxmlElement):
    """`a:bodyPr` custom element class."""

    _add_noAutofit: Callable[[], BaseOxmlElement]
    _add_normAutofit: Callable[[], CT_TextNormalAutofit]
    _add_spAutoFit: Callable[[], BaseOxmlElement]
    _remove_eg_textAutoFit: Callable[[], None]

    noAutofit: BaseOxmlElement | None
    normAutofit: CT_TextNormalAutofit | None
    spAutoFit: BaseOxmlElement | None

    eg_textAutoFit = ZeroOrOneChoice(
        (Choice("a:noAutofit"), Choice("a:normAutofit"), Choice("a:spAutoFit")),
        successors=("a:scene3d", "a:sp3d", "a:flatTx", "a:extLst"),
    )
    rot: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rot", ST_Angle, default=0.0
    )
    upright: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "upright", XsdBoolean, default=False
    )
    lIns: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "lIns", ST_Coordinate32, default=Emu(91440)
    )
    tIns: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "tIns", ST_Coordinate32, default=Emu(45720)
    )
    rIns: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rIns", ST_Coordinate32, default=Emu(91440)
    )
    bIns: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "bIns", ST_Coordinate32, default=Emu(45720)
    )
    anchor: MSO_VERTICAL_ANCHOR | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "anchor", MSO_VERTICAL_ANCHOR
    )
    wrap: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "wrap", ST_TextWrappingType
    )

    @property
    def autofit(self):
        """The autofit setting for the text frame, a member of the `MSO_AUTO_SIZE` enumeration."""
        if self.noAutofit is not None:
            return MSO_AUTO_SIZE.NONE
        if self.normAutofit is not None:
            return MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        if self.spAutoFit is not None:
            return MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        return None

    @autofit.setter
    def autofit(self, value: MSO_AUTO_SIZE | None):
        if value is not None and value not in MSO_AUTO_SIZE:
            raise ValueError(
                f"only None or a member of the MSO_AUTO_SIZE enumeration can be assigned to"
                f" CT_TextBodyProperties.autofit, got {value}"
            )
        self._remove_eg_textAutoFit()
        if value == MSO_AUTO_SIZE.NONE:
            self._add_noAutofit()
        elif value == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE:
            self._add_normAutofit()
        elif value == MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT:
            self._add_spAutoFit()

    @property
    def font_scale(self) -> float:
        """Effective `fontScale` attribute of `a:normAutofit` as percent.

        Returns the default 100.0 when no `a:normAutofit` child is present or when
        `a:normAutofit` has no explicit `fontScale` attribute.
        """
        normAutofit = self.normAutofit
        if normAutofit is None:
            return 100.0
        return normAutofit.fontScale

    @font_scale.setter
    def font_scale(self, value: float):
        """Set the `fontScale` attribute of `a:normAutofit` to `value` percent.

        Ensures the `a:normAutofit` child element exists, replacing any other
        autofit-choice child (`a:noAutofit` or `a:spAutoFit`) that may be present.
        Assigning the default value of ``100.0`` removes the `fontScale`
        attribute from the element.
        """
        normAutofit = self._get_or_add_normAutofit()
        # -- assign via descriptor; default (100.0) clears the attribute --
        normAutofit.fontScale = value

    @property
    def line_space_reduction(self) -> float:
        """Effective `lnSpcReduction` attribute of `a:normAutofit` as percent.

        Returns the default 0.0 when no `a:normAutofit` child is present or when
        `a:normAutofit` has no explicit `lnSpcReduction` attribute.
        """
        normAutofit = self.normAutofit
        if normAutofit is None:
            return 0.0
        return normAutofit.lnSpcReduction

    @line_space_reduction.setter
    def line_space_reduction(self, value: float):
        """Set the `lnSpcReduction` attribute of `a:normAutofit` to `value` percent.

        Ensures the `a:normAutofit` child element exists, replacing any other
        autofit-choice child (`a:noAutofit` or `a:spAutoFit`) that may be present.
        Assigning the default value of ``0.0`` removes the `lnSpcReduction`
        attribute from the element.
        """
        normAutofit = self._get_or_add_normAutofit()
        normAutofit.lnSpcReduction = value

    def _get_or_add_normAutofit(self):
        """Return the `a:normAutofit` child, creating it if necessary.

        Any existing `a:noAutofit` or `a:spAutoFit` choice sibling is removed
        first because only one of the three autofit-choice elements may appear.
        """
        if self.normAutofit is None:
            self._remove_eg_textAutoFit()
            self._add_normAutofit()
        return self.normAutofit


class CT_TextCharacterProperties(BaseOxmlElement):
    """Custom element class for `a:rPr`, `a:defRPr`, and `a:endParaRPr`.

    'rPr' is short for 'run properties', and it corresponds to the |Font| proxy class.
    """

    get_or_add_effectLst: Callable[[], "CT_EffectList"]
    get_or_add_highlight: Callable[[], "CT_Color"]
    get_or_add_hlinkClick: Callable[[], CT_Hyperlink]
    get_or_add_latin: Callable[[], CT_TextFont]
    get_or_add_ea: Callable[[], CT_TextFont]
    get_or_add_cs: Callable[[], CT_TextFont]
    _remove_effectLst: Callable[[], None]
    _remove_highlight: Callable[[], None]
    _remove_latin: Callable[[], None]
    _remove_ea: Callable[[], None]
    _remove_cs: Callable[[], None]
    _remove_hlinkClick: Callable[[], None]

    eg_fillProperties = ZeroOrOneChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        ),
        successors=(
            "a:effectLst",
            "a:effectDag",
            "a:highlight",
            "a:uLnTx",
            "a:uLn",
            "a:uFillTx",
            "a:uFill",
            "a:latin",
            "a:ea",
            "a:cs",
            "a:sym",
            "a:hlinkClick",
            "a:hlinkMouseOver",
            "a:rtl",
            "a:extLst",
        ),
    )
    effectLst: "CT_EffectList | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:effectLst",
        successors=(
            "a:effectDag",
            "a:highlight",
            "a:uLnTx",
            "a:uLn",
            "a:uFillTx",
            "a:uFill",
            "a:latin",
            "a:ea",
            "a:cs",
            "a:sym",
            "a:hlinkClick",
            "a:hlinkMouseOver",
            "a:rtl",
            "a:extLst",
        ),
    )
    highlight: "CT_Color | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:highlight",
        successors=(
            "a:uLnTx",
            "a:uLn",
            "a:uFillTx",
            "a:uFill",
            "a:latin",
            "a:ea",
            "a:cs",
            "a:sym",
            "a:hlinkClick",
            "a:hlinkMouseOver",
            "a:rtl",
            "a:extLst",
        ),
    )
    latin: CT_TextFont | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:latin",
        successors=(
            "a:ea",
            "a:cs",
            "a:sym",
            "a:hlinkClick",
            "a:hlinkMouseOver",
            "a:rtl",
            "a:extLst",
        ),
    )
    ea: CT_TextFont | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:ea",
        successors=(
            "a:cs",
            "a:sym",
            "a:hlinkClick",
            "a:hlinkMouseOver",
            "a:rtl",
            "a:extLst",
        ),
    )
    cs: CT_TextFont | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:cs",
        successors=(
            "a:sym",
            "a:hlinkClick",
            "a:hlinkMouseOver",
            "a:rtl",
            "a:extLst",
        ),
    )
    hlinkClick: CT_Hyperlink | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:hlinkClick", successors=("a:hlinkMouseOver", "a:rtl", "a:extLst")
    )

    lang: MSO_LANGUAGE_ID | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "lang", MSO_LANGUAGE_ID
    )
    sz: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sz", ST_TextFontSize
    )
    b: bool | None = OptionalAttribute("b", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    i: bool | None = OptionalAttribute("i", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    u: MSO_TEXT_UNDERLINE_TYPE | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "u", MSO_TEXT_UNDERLINE_TYPE
    )
    strike: MSO_TEXT_STRIKE_TYPE | None = OptionalAttribute(
        "strike", MSO_TEXT_STRIKE_TYPE
    )  # pyright: ignore[reportAssignmentType]
    baseline: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "baseline", ST_TextBaselinePercent
    )

    def _new_gradFill(self):
        return CT_GradientFillProperties.new_gradFill()

    def add_hlinkClick(self, rId: str) -> CT_Hyperlink:
        """Add an `a:hlinkClick` child element with r:id attribute set to `rId`."""
        hlinkClick = self.get_or_add_hlinkClick()
        hlinkClick.rId = rId
        return hlinkClick


class CT_TextField(BaseOxmlElement):
    """`a:fld` field element, for either a slide number or date field."""

    get_or_add_rPr: Callable[[], CT_TextCharacterProperties]
    get_or_add_t: Callable[[], BaseOxmlElement]

    rPr: CT_TextCharacterProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:rPr", successors=("a:pPr", "a:t")
    )
    t: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:t", successors=()
    )
    id: str = RequiredAttribute("id", XsdString)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute("type", XsdString)  # pyright: ignore[reportAssignmentType]

    @property
    def text(self) -> str:
        """The text of the `a:t` child element."""
        t = self.t
        if t is None:
            return ""
        return t.text or ""

    @text.setter
    def text(self, value: str):
        t = self.get_or_add_t()
        t.text = value


class CT_TextFont(BaseOxmlElement):
    """Custom element class for `a:latin`, `a:ea`, `a:cs`, and `a:sym`.

    These occur as child elements of CT_TextCharacterProperties, e.g. `a:rPr`.
    """

    typeface: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "typeface", ST_TextTypeface
    )


class CT_TextLineBreak(BaseOxmlElement):
    """`a:br` line break element"""

    get_or_add_rPr: Callable[[], CT_TextCharacterProperties]

    rPr = ZeroOrOne("a:rPr", successors=())

    @property
    def text(self):  # pyright: ignore[reportIncompatibleMethodOverride]
        """Unconditionally a single vertical-tab character.

        A line break element can contain no text other than the implicit line feed it
        represents.
        """
        return "\v"


class CT_TextNormalAutofit(BaseOxmlElement):
    """`a:normAutofit` element specifying fit text to shape font reduction, etc."""

    fontScale = OptionalAttribute(
        "fontScale", ST_TextFontScalePercentOrPercentString, default=100.0
    )
    lnSpcReduction = OptionalAttribute(
        "lnSpcReduction", ST_TextFontScaleReductionPercent, default=0.0
    )


class CT_TextParagraph(BaseOxmlElement):
    """`a:p` custom element class"""

    get_or_add_endParaRPr: Callable[[], CT_TextCharacterProperties]
    get_or_add_pPr: Callable[[], CT_TextParagraphProperties]
    r_lst: list[CT_RegularTextRun]
    _add_br: Callable[[], CT_TextLineBreak]
    _add_fld: Callable[[], CT_TextField]
    _add_r: Callable[[], CT_RegularTextRun]

    pPr: CT_TextParagraphProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:pPr", successors=("a:r", "a:br", "a:fld", "a:endParaRPr")
    )
    r = ZeroOrMore("a:r", successors=("a:endParaRPr",))
    br = ZeroOrMore("a:br", successors=("a:endParaRPr",))
    fld = ZeroOrMore("a:fld", successors=("a:endParaRPr",))
    endParaRPr: CT_TextCharacterProperties | None = ZeroOrOne(
        "a:endParaRPr", successors=()
    )  # pyright: ignore[reportAssignmentType]

    def add_br(self) -> CT_TextLineBreak:
        """Return a newly appended `a:br` element."""
        return self._add_br()

    def add_fld(
        self,
        fld_type: str,
        text: str = "",
        fld_id: str | None = None,
    ) -> CT_TextField:
        """Return a newly appended `a:fld` (auto-refresh field) element.

        `fld_type` becomes the `a:fld/@type` value, e.g. `'slidenum'`, `'datetime'`,
        `'datetimeFigureOut'`, `'datetime1'` … `'datetime13'`, or `'footer'`. `text` is the
        placeholder-display string PowerPoint shows until the field is first rendered (a
        slide-number field typically uses `'#'` or `'‹#›'`). `fld_id` is the optional
        field-GUID — one is generated when omitted.
        """
        from uuid import uuid4

        guid = fld_id if fld_id is not None else "{%s}" % str(uuid4()).upper()
        fld = self._add_fld()
        fld.id = guid
        fld.type = fld_type
        if text:
            fld.text = text
        return fld

    def add_math_equation(self, omml_xml: str) -> BaseOxmlElement:
        """Append an `mc:AlternateContent` wrapping `omml_xml` to this `a:p`.

        `omml_xml` is caller-provided Office Math Markup Language (typically the output of
        Microsoft's `MML2OMML.XSL`) -- either a serialized `m:oMath` element or a `m:oMathPara`
        wrapper containing one or more `m:oMath` children. It must declare the `m` namespace
        (`http://schemas.openxmlformats.org/officeDocument/2006/math`) on its root element.

        The method wraps the fragment in the `mc:AlternateContent/mc:Choice[Requires="a14"]/a14:m`
        scaffolding that PowerPoint emits for an equation embedded inline in a paragraph,
        accompanied by a `mc:Fallback/a:r` run carrying the OMML reduced to its visible text
        characters (concatenated `m:t` contents). Any existing runs / line-breaks / fields in the
        paragraph are preserved -- the `mc:AlternateContent` is appended after them and before
        any `a:endParaRPr`.

        Returns the newly inserted `mc:AlternateContent` element. Raises `ValueError` when
        `omml_xml` cannot be parsed as XML, does not carry a recognisable `m:oMath` element, or
        its root is not `m:oMath` / `m:oMathPara`.
        """
        oMath_or_para = _parse_omml_fragment(omml_xml)
        fallback_text = _extract_fallback_text(oMath_or_para)

        # -- build the mc:AlternateContent scaffolding with the caller's OMML embedded --
        ac = parse_xml(_AC_WRAPPER_XML.format(fallback=fallback_text))
        # -- locate the a14:m element and append the (re-parsed) OMML fragment --
        a14_m = ac.find(".//{http://schemas.microsoft.com/office/drawing/2010/main}m")
        assert a14_m is not None  # pragma: no cover  (guaranteed by template)
        a14_m.append(oMath_or_para)

        # -- insert before any endParaRPr so schema order is preserved --
        endParaRPr = self.endParaRPr
        if endParaRPr is not None:
            endParaRPr.addprevious(ac)
        else:
            self.append(ac)
        return cast("BaseOxmlElement", ac)

    def add_r(self, text: str | None = None) -> CT_RegularTextRun:
        """Return a newly appended `a:r` element."""
        r = self._add_r()
        if text:
            r.text = text
        return r

    def append_text(self, text: str):
        """Append `a:r` and `a:br` elements to `p` based on `text`.

        Any `\n` or `\v` (vertical-tab) characters in `text` delimit `a:r` (run) elements and
        themselves are translated to `a:br` (line-break) elements. The vertical-tab character
        appears in clipboard text from PowerPoint at "soft" line-breaks (new-line, but not new
        paragraph).
        """
        for idx, r_str in enumerate(re.split("\n|\v", text)):
            # ---breaks are only added _between_ items, not at start---
            if idx > 0:
                self.add_br()
            # ---runs that would be empty are not added---
            if r_str:
                self.add_r(r_str)

    @property
    def content_children(self) -> tuple[CT_RegularTextRun | CT_TextLineBreak | CT_TextField, ...]:
        """Sequence containing text-container child elements of this `a:p` element.

        These include `a:r`, `a:br`, and `a:fld`.
        """
        return tuple(
            e for e in self if isinstance(e, (CT_RegularTextRun, CT_TextLineBreak, CT_TextField))
        )

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """str text contained in this paragraph."""
        # ---note this shadows the lxml _Element.text---
        return "".join([child.text for child in self.content_children])

    def _new_r(self):
        r_xml = "<a:r %s><a:t/></a:r>" % nsdecls("a")
        return parse_xml(r_xml)


# -- mc:AlternateContent template used by CT_TextParagraph.add_math_equation --
#
# The caller's OMML fragment is appended under the `a14:m` wrapper and the fallback
# run's `a:t` child is populated with the visible text extracted from the OMML.
_AC_WRAPPER_XML = (
    "<mc:AlternateContent "
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    "<mc:Choice "
    'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" '
    'Requires="a14">'
    "<a14:m/>"
    "</mc:Choice>"
    "<mc:Fallback>"
    '<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    "<a:t>{fallback}</a:t>"
    "</a:r>"
    "</mc:Fallback>"
    "</mc:AlternateContent>"
)


def _parse_omml_fragment(omml_xml: str) -> BaseOxmlElement:
    """Return parsed OMML `m:oMath` or `m:oMathPara` element from `omml_xml`.

    Raises `ValueError` on malformed input or when the root is not `m:oMath` / `m:oMathPara`.
    """
    from pptx.oxml.ns import qn

    try:
        elm = parse_xml(omml_xml)
    except Exception as exc:  # noqa: BLE001 -- any parse failure is user error
        raise ValueError(f"`omml_xml` is not well-formed XML: {exc}") from exc
    oMath_tag = qn("m:oMath")
    oMathPara_tag = qn("m:oMathPara")
    if elm.tag not in (oMath_tag, oMathPara_tag):
        raise ValueError(
            "`omml_xml` root element must be `m:oMath` or `m:oMathPara`, got " f"{elm.tag!r}"
        )
    # -- require at least one m:oMath somewhere (either the root, or inside m:oMathPara) --
    if elm.tag == oMathPara_tag and elm.find(oMath_tag) is None:
        raise ValueError("`omml_xml` must contain at least one `m:oMath` element")
    return cast("BaseOxmlElement", elm)


def _extract_fallback_text(oMath_or_para: BaseOxmlElement) -> str:
    """Return a best-effort plain-text rendering of an `m:oMath` subtree.

    Concatenates the text content of every `m:t` descendant, with XML-special characters
    escaped so the returned string is safe to embed in an `a:t` element's text content.
    Empty string is returned (producing an empty fallback run) when no `m:t` children are
    found.
    """
    from xml.sax.saxutils import escape as _xml_escape

    from pptx.oxml.ns import qn

    t_tag = qn("m:t")
    parts = [t.text or "" for t in oMath_or_para.iter(t_tag)]
    return _xml_escape("".join(parts))


class CT_TextParagraphProperties(BaseOxmlElement):
    """`a:pPr` custom element class."""

    get_or_add_defRPr: Callable[[], CT_TextCharacterProperties]
    get_or_add_buClr: Callable[[], "CT_TextBulletColor"]
    get_or_add_buFont: Callable[[], "CT_TextFont"]
    _add_lnSpc: Callable[[], CT_TextSpacing]
    _add_spcAft: Callable[[], CT_TextSpacing]
    _add_spcBef: Callable[[], CT_TextSpacing]
    _add_buSzPct: Callable[[], "CT_TextBulletSizePercent"]
    _add_buSzPts: Callable[[], "CT_TextBulletSizePoint"]
    _remove_lnSpc: Callable[[], None]
    _remove_spcAft: Callable[[], None]
    _remove_spcBef: Callable[[], None]
    _remove_buClr: Callable[[], None]
    _remove_buFont: Callable[[], None]
    _remove_buSzPct: Callable[[], None]
    _remove_buSzPts: Callable[[], None]
    _remove_eg_textBullet: Callable[[], None]
    _add_buNone: Callable[[], "CT_TextNoBullet"]
    _add_buChar: Callable[[], "CT_TextCharBullet"]
    _add_buAutoNum: Callable[[], "CT_TextAutonumberBullet"]

    buNone: "CT_TextNoBullet | None"
    buChar: "CT_TextCharBullet | None"
    buAutoNum: "CT_TextAutonumberBullet | None"

    _tag_seq = (
        "a:lnSpc",
        "a:spcBef",
        "a:spcAft",
        "a:buClrTx",
        "a:buClr",
        "a:buSzTx",
        "a:buSzPct",
        "a:buSzPts",
        "a:buFontTx",
        "a:buFont",
        "a:buNone",
        "a:buAutoNum",
        "a:buChar",
        "a:buBlip",
        "a:tabLst",
        "a:defRPr",
        "a:extLst",
    )
    lnSpc: CT_TextSpacing | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:lnSpc", successors=_tag_seq[1:]
    )
    spcBef: CT_TextSpacing | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:spcBef", successors=_tag_seq[2:]
    )
    spcAft: CT_TextSpacing | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:spcAft", successors=_tag_seq[3:]
    )
    buClr: "CT_TextBulletColor | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:buClr", successors=_tag_seq[5:]
    )
    buSzPct: "CT_TextBulletSizePercent | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:buSzPct", successors=_tag_seq[7:]
    )
    buSzPts: "CT_TextBulletSizePoint | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:buSzPts", successors=_tag_seq[8:]
    )
    buFont: "CT_TextFont | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:buFont", successors=_tag_seq[10:]
    )
    eg_textBullet = ZeroOrOneChoice(
        (Choice("a:buNone"), Choice("a:buAutoNum"), Choice("a:buChar")),
        successors=_tag_seq[13:],
    )
    defRPr: CT_TextCharacterProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:defRPr", successors=_tag_seq[16:]
    )
    lvl: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "lvl", ST_TextIndentLevelType, default=0
    )
    algn: PP_PARAGRAPH_ALIGNMENT | None = OptionalAttribute(
        "algn", PP_PARAGRAPH_ALIGNMENT
    )  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    @property
    def bullet_type(self) -> str | None:
        """One of ``"char"``, ``"autonum"``, ``"none"`` or |None|.

        |None| indicates no explicit bullet setting is present on this paragraph; the
        effective bullet is inherited from the style hierarchy.
        """
        if self.buChar is not None:
            return "char"
        if self.buAutoNum is not None:
            return "autonum"
        if self.buNone is not None:
            return "none"
        return None

    @property
    def bullet_char(self) -> str | None:
        """The single bullet character, or |None| when no `a:buChar` child is present."""
        buChar = self.buChar
        return None if buChar is None else buChar.char

    @property
    def bullet_number_scheme(self) -> PP_AUTO_NUMBER_SCHEME | None:
        """Numbering scheme, or |None| when no `a:buAutoNum` child is present."""
        buAutoNum = self.buAutoNum
        return None if buAutoNum is None else buAutoNum.type

    @property
    def bullet_number_start_at(self) -> int | None:
        """Start-at value for autonum bullet, or |None| when no `a:buAutoNum` child is present.

        When `a:buAutoNum` is present but `@startAt` is not, the default value of 1
        (per ECMA-376) is returned.
        """
        buAutoNum = self.buAutoNum
        if buAutoNum is None:
            return None
        return buAutoNum.startAt

    def clear_bullet(self) -> None:
        """Remove any `a:buNone`, `a:buAutoNum`, or `a:buChar` child.

        This causes the paragraph to inherit its bullet from the style hierarchy.
        """
        self._remove_eg_textBullet()

    def set_no_bullet(self) -> None:
        """Add an `a:buNone` child, replacing any existing bullet choice element.

        This explicitly suppresses any inherited bullet for this paragraph.
        """
        self._remove_eg_textBullet()
        self._add_buNone()

    def set_char_bullet(self, char: str) -> "CT_TextCharBullet":
        """Add an `a:buChar` child with `@char` set to `char`, replacing any existing bullet.

        Returns the newly added `a:buChar` element.
        """
        self._remove_eg_textBullet()
        buChar = self._add_buChar()
        buChar.char = char
        return buChar

    def set_auto_number_bullet(
        self, scheme: PP_AUTO_NUMBER_SCHEME, start_at: int | None = None
    ) -> "CT_TextAutonumberBullet":
        """Add an `a:buAutoNum` child, replacing any existing bullet choice element.

        `scheme` is a member of :class:`PP_AUTO_NUMBER_SCHEME`. If `start_at` is not
        |None|, the corresponding `@startAt` attribute is written.
        """
        self._remove_eg_textBullet()
        buAutoNum = self._add_buAutoNum()
        buAutoNum.type = scheme
        if start_at is not None:
            buAutoNum.startAt = start_at
        return buAutoNum

    @property
    def bullet_font(self) -> str | None:
        """Typeface of the bullet `a:buFont` child, or |None| when not present."""
        buFont = self.buFont
        return None if buFont is None else buFont.typeface

    @bullet_font.setter
    def bullet_font(self, value: str | None):
        if value is None:
            self._remove_buFont()
            return
        buFont = self.get_or_add_buFont()
        buFont.typeface = value

    @property
    def bullet_size_pct(self) -> float | None:
        """Size-percent value of the bullet `a:buSzPct` child, or |None|.

        Returns a float in range 0.25..4.0 (e.g. 0.75 for 75%).
        """
        buSzPct = self.buSzPct
        return None if buSzPct is None else buSzPct.val

    @bullet_size_pct.setter
    def bullet_size_pct(self, value: float | None):
        # -- assigning any size removes any conflicting size element --
        self._remove_buSzPts()
        self._remove_buSzPct()
        if value is None:
            return
        self._add_buSzPct().val = value

    @property
    def bullet_size_points(self) -> Length | None:
        """Size-in-points of the bullet `a:buSzPts` child, or |None|.

        Returns a |Length| value (EMU) derived from the centipoint XML value.
        """
        buSzPts = self.buSzPts
        if buSzPts is None:
            return None
        # -- XML @val is in centipoints; translate to a Length (EMU) --
        return Centipoints(buSzPts.val)

    @bullet_size_points.setter
    def bullet_size_points(self, value: Length | None):
        # -- assigning any size removes any conflicting size element --
        self._remove_buSzPct()
        self._remove_buSzPts()
        if value is None:
            return
        # -- stored attribute is in centipoints (hundredths of a point) --
        self._add_buSzPts().val = Emu(value).centipoints

    def clear_bullet_size(self) -> None:
        """Remove any `a:buSzPct` or `a:buSzPts` child.

        This causes the bullet size to inherit from the style hierarchy.
        """
        self._remove_buSzPct()
        self._remove_buSzPts()

    @property
    def line_spacing(self) -> float | Length | None:
        """The spacing between baselines of successive lines in this paragraph.

        A float value indicates a number of lines. A |Length| value indicates a fixed spacing.
        Value is contained in `./a:lnSpc/a:spcPts/@val` or `./a:lnSpc/a:spcPct/@val`. Value is
        |None| if no element is present.
        """
        lnSpc = self.lnSpc
        if lnSpc is None:
            return None
        if lnSpc.spcPts is not None:
            return lnSpc.spcPts.val
        return cast(CT_TextSpacingPercent, lnSpc.spcPct).val

    @line_spacing.setter
    def line_spacing(self, value: float | Length | None):
        self._remove_lnSpc()
        if value is None:
            return
        if isinstance(value, Length):
            self._add_lnSpc().set_spcPts(value)
        else:
            self._add_lnSpc().set_spcPct(value)

    @property
    def space_after(self) -> Length | None:
        """The EMU equivalent of the centipoints value in `./a:spcAft/a:spcPts/@val`."""
        spcAft = self.spcAft
        if spcAft is None:
            return None
        spcPts = spcAft.spcPts
        if spcPts is None:
            return None
        return spcPts.val

    @space_after.setter
    def space_after(self, value: Length | None):
        self._remove_spcAft()
        if value is not None:
            self._add_spcAft().set_spcPts(value)

    @property
    def space_before(self):
        """The EMU equivalent of the centipoints value in `./a:spcBef/a:spcPts/@val`."""
        spcBef = self.spcBef
        if spcBef is None:
            return None
        spcPts = spcBef.spcPts
        if spcPts is None:
            return None
        return spcPts.val

    @space_before.setter
    def space_before(self, value: Length | None):
        self._remove_spcBef()
        if value is not None:
            self._add_spcBef().set_spcPts(value)


class CT_TextSpacing(BaseOxmlElement):
    """Used for `a:lnSpc`, `a:spcBef`, and `a:spcAft` elements."""

    get_or_add_spcPct: Callable[[], CT_TextSpacingPercent]
    get_or_add_spcPts: Callable[[], CT_TextSpacingPoint]
    _remove_spcPct: Callable[[], None]
    _remove_spcPts: Callable[[], None]

    # this should actually be a OneAndOnlyOneChoice, but that's not
    # implemented yet.
    spcPct: CT_TextSpacingPercent | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:spcPct"
    )
    spcPts: CT_TextSpacingPoint | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:spcPts"
    )

    def set_spcPct(self, value: float):
        """Set spacing to `value` lines, e.g. 1.75 lines.

        A ./a:spcPts child is removed if present.
        """
        self._remove_spcPts()
        spcPct = self.get_or_add_spcPct()
        spcPct.val = value

    def set_spcPts(self, value: Length):
        """Set spacing to `value` points. A ./a:spcPct child is removed if present."""
        self._remove_spcPct()
        spcPts = self.get_or_add_spcPts()
        spcPts.val = value


class CT_TextSpacingPercent(BaseOxmlElement):
    """`a:spcPct` element, specifying spacing in thousandths of a percent in its `val` attribute."""

    val: float = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "val", ST_TextSpacingPercentOrPercentString
    )


class CT_TextSpacingPoint(BaseOxmlElement):
    """`a:spcPts` element, specifying spacing in centipoints in its `val` attribute."""

    val: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "val", ST_TextSpacingPoint
    )


class CT_TextNoBullet(BaseOxmlElement):
    """`a:buNone` element.

    Explicitly suppresses any inherited bullet on its parent paragraph-properties element.
    """


class CT_TextCharBullet(BaseOxmlElement):
    """`a:buChar` element.

    Specifies a single character used as the paragraph's bullet glyph, in its
    ``@char`` attribute.
    """

    char: str = RequiredAttribute("char", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TextAutonumberBullet(BaseOxmlElement):
    """`a:buAutoNum` element.

    Specifies an automatic-numbering scheme used as the paragraph's bullet, in its
    ``@type`` attribute, along with an optional ``@startAt`` starting ordinal.
    """

    type: PP_AUTO_NUMBER_SCHEME = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "type", PP_AUTO_NUMBER_SCHEME
    )
    startAt: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "startAt", ST_TextBulletStartAtNum, default=1
    )


class CT_TextBulletColor(BaseOxmlElement):
    """`a:buClr` element.

    Specifies an explicit color for the paragraph's bullet glyph. Contains a
    single child of the ``EG_ColorChoice`` group, e.g. ``a:srgbClr`` or
    ``a:schemeClr``.
    """

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )


class CT_TextBulletSizePercent(BaseOxmlElement):
    """`a:buSzPct` element.

    Specifies the bullet size as a percentage of the size of the text on the
    paragraph. Value is in its ``@val`` attribute, typically as a percent
    literal like ``"75%"``. Valid range is 25%..400%.
    """

    val: float = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "val", ST_TextBulletSizePercent
    )


class CT_TextBulletSizePoint(BaseOxmlElement):
    """`a:buSzPts` element.

    Specifies the bullet size as an absolute point size, in its ``@val``
    attribute. The XML integer value is in centipoints (hundredths of a
    point). Valid range is 1pt..4000pt.
    """

    val: Length = RequiredAttribute("val", ST_TextFontSize)  # pyright: ignore[reportAssignmentType]
