"""Text-related objects such as TextFrame and Paragraph."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, cast

from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_AUTO_SIZE, MSO_UNDERLINE, MSO_VERTICAL_ANCHOR, PP_AUTO_NUMBER_SCHEME
from pptx.enum.text import MSO_AUTO_SIZE, MSO_STRIKE, MSO_UNDERLINE, MSO_VERTICAL_ANCHOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.simpletypes import ST_TextWrappingType
from pptx.shapes import Subshape
from pptx.text.fonts import FontFiles
from pptx.text.layout import TextFitter
from pptx.util import Centipoints, Emu, Length, Pt, lazyproperty

if TYPE_CHECKING:
    from pptx.dml.color import ColorFormat
    from pptx.enum.text import (
        MSO_TEXT_STRIKE_TYPE,
        MSO_TEXT_UNDERLINE_TYPE,
        MSO_VERTICAL_ANCHOR,
        PP_PARAGRAPH_ALIGNMENT,
    )
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.text import (
        CT_RegularTextRun,
        CT_TextBody,
        CT_TextCharacterProperties,
        CT_TextParagraph,
        CT_TextParagraphProperties,
    )
    from pptx.types import ProvidesExtents, ProvidesPart


class TextFrame(Subshape):
    """The part of a shape that contains its text.

    Not all shapes have a text frame. Corresponds to the `p:txBody` element that can
    appear as a child element of `p:sp`. Not intended to be constructed directly.
    """

    def __init__(self, txBody: CT_TextBody, parent: ProvidesPart):
        super(TextFrame, self).__init__(parent)
        self._element = self._txBody = txBody
        self._parent = parent

    def add_paragraph(self):
        """
        Return new |_Paragraph| instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        p = self._txBody.add_p()
        return _Paragraph(p, self)

    @property
    def auto_size(self) -> MSO_AUTO_SIZE | None:
        """Resizing strategy used to fit text within this shape.

        Determins the type of automatic resizing used to fit the text of this shape within its
        bounding box when the text would otherwise extend beyond the shape boundaries. May be
        |None|, `MSO_AUTO_SIZE.NONE`, `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`, or
        `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`.
        """
        return self._bodyPr.autofit

    @auto_size.setter
    def auto_size(self, value: MSO_AUTO_SIZE | None):
        self._bodyPr.autofit = value

    def clear(self):
        """Remove all paragraphs except one empty one."""
        for p in self._txBody.p_lst[1:]:
            self._txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()

    def fit_text(
        self,
        font_family: str = "Calibri",
        max_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        font_file: str | None = None,
    ):
        """Fit text-frame text entirely within bounds of its shape.

        Make the text in this text frame fit entirely within the bounds of its shape by setting
        word wrap on and applying the "best-fit" font size to all the text it contains.

        :attr:`TextFrame.auto_size` is set to :attr:`MSO_AUTO_SIZE.NONE`. The font size will not
        be set larger than `max_size` points. If the path to a matching TrueType font is provided
        as `font_file`, that font file will be used for the font metrics. If `font_file` is |None|,
        best efforts are made to locate a font file with matchhing `font_family`, `bold`, and
        `italic` installed on the current system (usually succeeds if the font is installed).

        When the text contains a word too wide to fit the shape at `max_size` but that fits
        at a smaller point size, that smaller size is used (see issue #936).
        Raises :class:`pptx.exc.TextLayoutError` when no point size between 1 and `max_size`
        allows the text to fit the shape -- for example when a single word is too wide to fit
        the shape at the smallest considered size (see issue #773).
        """
        # ---no-op when empty as fit behavior not defined for that case---
        if self.text == "":
            return  # pragma: no cover

        font_size = self._best_fit_font_size(font_family, max_size, bold, italic, font_file)
        self._apply_fit(font_family, font_size, bold, italic)

    @property
    def margin_bottom(self) -> Length:
        """|Length| value representing the inset of text from the bottom text frame border.

        :meth:`pptx.util.Inches` provides a convenient way of setting the value, e.g.
        `text_frame.margin_bottom = Inches(0.05)`.
        """
        return self._bodyPr.bIns

    @margin_bottom.setter
    def margin_bottom(self, emu: Length):
        self._bodyPr.bIns = emu

    @property
    def margin_left(self) -> Length:
        """Inset of text from left text frame border as |Length| value."""
        return self._bodyPr.lIns

    @margin_left.setter
    def margin_left(self, emu: Length):
        self._bodyPr.lIns = emu

    @property
    def margin_right(self) -> Length:
        """Inset of text from right text frame border as |Length| value."""
        return self._bodyPr.rIns

    @margin_right.setter
    def margin_right(self, emu: Length):
        self._bodyPr.rIns = emu

    @property
    def margin_top(self) -> Length:
        """Inset of text from top text frame border as |Length| value."""
        return self._bodyPr.tIns

    @margin_top.setter
    def margin_top(self, emu: Length):
        self._bodyPr.tIns = emu

    @property
    def paragraphs(self) -> tuple[_Paragraph, ...]:
        """Sequence of paragraphs in this text frame.

        A text frame always contains at least one paragraph.
        """
        return tuple([_Paragraph(p, self) for p in self._txBody.p_lst])

    @property
    def text(self) -> str:
        """All text in this text-frame as a single string.

        Read/write. The return value contains all text in this text-frame. A line-feed character
        (`"\\n"`) separates the text for each paragraph. A vertical-tab character (`"\\v"`) appears
        for each line break (aka. soft carriage-return) encountered.

        The vertical-tab character is how PowerPoint represents a soft carriage return in clipboard
        text, which is why that encoding was chosen.

        Assignment replaces all text in the text frame. A new paragraph is added for each line-feed
        character (`"\\n"`) encountered. A line-break (soft carriage-return) is inserted for each
        vertical-tab character (`"\\v"`) encountered.

        Any control character other than newline, tab, or vertical-tab are escaped as plain-text
        like "_x001B_" (for ESC (ASCII 32) in this example).
        """
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, text: str):
        txBody = self._txBody
        txBody.clear_content()
        for p_text in text.split("\n"):
            p = txBody.add_p()
            p.append_text(p_text)

    @property
    def vertical_anchor(self) -> MSO_VERTICAL_ANCHOR | None:
        """Represents the vertical alignment of text in this text frame.

        |None| indicates the effective value should be inherited from this object's style hierarchy.
        """
        return self._txBody.bodyPr.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, value: MSO_VERTICAL_ANCHOR | None):
        bodyPr = self._txBody.bodyPr
        bodyPr.anchor = value

    @property
    def word_wrap(self) -> bool | None:
        """`True` when lines of text in this shape are wrapped to fit within the shape's width.

        Read-write. Valid values are True, False, or None. True and False turn word wrap on and
        off, respectively. Assigning None to word wrap causes any word wrap setting to be removed
        from the text frame, causing it to inherit this setting from its style hierarchy.
        """
        return {
            ST_TextWrappingType.SQUARE: True,
            ST_TextWrappingType.NONE: False,
            None: None,
        }[self._txBody.bodyPr.wrap]

    @word_wrap.setter
    def word_wrap(self, value: bool | None):
        if value not in (True, False, None):
            raise ValueError(  # pragma: no cover
                "assigned value must be True, False, or None, got %s" % value
            )
        self._txBody.bodyPr.wrap = {
            True: ST_TextWrappingType.SQUARE,
            False: ST_TextWrappingType.NONE,
            None: None,
        }[value]

    def _apply_fit(self, font_family: str, font_size: int, is_bold: bool, is_italic: bool):
        """Arrange text in this text frame to fit inside its extents.

        This is accomplished by setting auto size off, wrap on, and setting the font of
        all its text to `font_family`, `font_size`, `is_bold`, and `is_italic`.
        """
        self.auto_size = MSO_AUTO_SIZE.NONE
        self.word_wrap = True
        self._set_font(font_family, font_size, is_bold, is_italic)

    def _best_fit_font_size(
        self, family: str, max_size: int, bold: bool, italic: bool, font_file: str | None
    ) -> int:
        """Return font-size in points that best fits text in this text-frame.

        The best-fit font size is the largest integer point size not greater than `max_size` that
        allows all the text in this text frame to fit inside its extents when rendered using the
        font described by `family`, `bold`, and `italic`. If `font_file` is specified, it is used
        to calculate the fit, whether or not it matches `family`, `bold`, and `italic`.
        """
        if font_file is None:
            font_file = FontFiles.find(family, bold, italic)
        return TextFitter.best_fit_font_size(self.text, self._extents, max_size, font_file)

    @property
    def _bodyPr(self):
        return self._txBody.bodyPr

    @property
    def _extents(self) -> tuple[Length, Length]:
        """(cx, cy) 2-tuple representing the effective rendering area of this text-frame.

        Margins are taken into account.
        """
        parent = cast("ProvidesExtents", self._parent)
        return (
            Length(parent.width - self.margin_left - self.margin_right),
            Length(parent.height - self.margin_top - self.margin_bottom),
        )

    def _set_font(self, family: str, size: int, bold: bool, italic: bool):
        """Set the font properties of all the text in this text frame."""

        def iter_rPrs(txBody: CT_TextBody) -> Iterator[CT_TextCharacterProperties]:
            for p in txBody.p_lst:
                for elm in p.content_children:
                    yield elm.get_or_add_rPr()
                # generate a:endParaRPr for each <a:p> element
                yield p.get_or_add_endParaRPr()

        def set_rPr_font(
            rPr: CT_TextCharacterProperties, name: str, size: int, bold: bool, italic: bool
        ):
            f = Font(rPr)
            f.name, f.size, f.bold, f.italic = family, Pt(size), bold, italic

        txBody = self._element
        for rPr in iter_rPrs(txBody):
            set_rPr_font(rPr, family, size, bold, italic)


class Font(object):
    """Character properties object, providing font size, font name, bold, italic, etc.

    Corresponds to `a:rPr` child element of a run. Also appears as `a:defRPr` and
    `a:endParaRPr` in paragraph and `a:defRPr` in list style elements.
    """

    def __init__(self, rPr: CT_TextCharacterProperties):
        super(Font, self).__init__()
        self._element = self._rPr = rPr

    @property
    def bold(self) -> bool | None:
        """Get or set boolean bold value of |Font|, e.g. `paragraph.font.bold = True`.

        If set to |None|, the bold setting is cleared and is inherited from an enclosing shape's
        setting, or a setting in a style or master. Returns None if no bold attribute is present,
        meaning the effective bold value is inherited from a master or the theme.
        """
        return self._rPr.b

    @bold.setter
    def bold(self, value: bool | None):
        self._rPr.b = value

    @lazyproperty
    def color(self) -> ColorFormat:
        """The |ColorFormat| instance that provides access to the color settings for this font."""
        if self.fill.type != MSO_FILL.SOLID:
            self.fill.solid()
        return self.fill.fore_color

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this font.

        Provides access to fill properties such as fill color.
        """
        return FillFormat.from_fill_parent(self._rPr)

    @property
    def italic(self) -> bool | None:
        """Get or set boolean italic value of |Font| instance.

        Has the same behaviors as bold with respect to None values.
        """
        return self._rPr.i

    @italic.setter
    def italic(self, value: bool | None):
        self._rPr.i = value

    @property
    def language_id(self) -> MSO_LANGUAGE_ID | None:
        """Get or set the language id of this |Font| instance.

        The language id is a member of the :ref:`MsoLanguageId` enumeration. Assigning |None|
        removes any language setting, the same behavior as assigning `MSO_LANGUAGE_ID.NONE`.
        """
        lang = self._rPr.lang
        if lang is None:
            return MSO_LANGUAGE_ID.NONE
        return self._rPr.lang

    @language_id.setter
    def language_id(self, value: MSO_LANGUAGE_ID | None):
        if value == MSO_LANGUAGE_ID.NONE:
            value = None
        self._rPr.lang = value

    @property
    def name(self) -> str | None:
        """Get or set the Latin-script typeface name for this |Font| instance.

        Causes Latin-script text it controls to appear in the named font, if a matching font is
        found. Corresponds to the ``a:latin`` child of the run-properties element. Returns |None|
        if the typeface is currently inherited from the theme. Setting it to |None| removes any
        override of the theme typeface.

        To set the typeface used for East-Asian (CJK) text, use :attr:`name_ea`. To set the
        typeface used for complex-script text (e.g. Arabic, Hebrew, Thai), use :attr:`name_cs`.
        """
        latin = self._rPr.latin
        if latin is None:
            return None
        return latin.typeface

    @name.setter
    def name(self, value: str | None):
        if value is None:
            self._rPr._remove_latin()  # pyright: ignore[reportPrivateUsage]
        else:
            latin = self._rPr.get_or_add_latin()
            latin.typeface = value

    @property
    def name_ea(self) -> str | None:
        """Get or set the East-Asian typeface name for this |Font| instance.

        Corresponds to the ``a:ea`` child of the run-properties element and controls how
        East-Asian (CJK -- Chinese, Japanese, Korean) characters in the run are rendered.
        Returns |None| when no explicit setting is present; in that case the effective typeface
        is inherited from the theme or style hierarchy. Assigning |None| removes any override.
        """
        ea = self._rPr.ea
        if ea is None:
            return None
        return ea.typeface

    @name_ea.setter
    def name_ea(self, value: str | None):
        if value is None:
            self._rPr._remove_ea()  # pyright: ignore[reportPrivateUsage]
        else:
            ea = self._rPr.get_or_add_ea()
            ea.typeface = value

    @property
    def name_cs(self) -> str | None:
        """Get or set the complex-script typeface name for this |Font| instance.

        Corresponds to the ``a:cs`` child of the run-properties element and controls how
        complex-script characters (e.g. Arabic, Hebrew, Thai, Devanagari) in the run are
        rendered. Returns |None| when no explicit setting is present; in that case the effective
        typeface is inherited from the theme or style hierarchy. Assigning |None| removes any
        override.
        """
        cs = self._rPr.cs
        if cs is None:
            return None
        return cs.typeface

    @name_cs.setter
    def name_cs(self, value: str | None):
        if value is None:
            self._rPr._remove_cs()  # pyright: ignore[reportPrivateUsage]
        else:
            cs = self._rPr.get_or_add_cs()
            cs.typeface = value

    @property
    def size(self) -> Length | None:
        """Indicates the font height in English Metric Units (EMU).

        Read/write. |None| indicates the font size should be inherited from its style hierarchy,
        such as a placeholder or document defaults (usually 18pt). |Length| is a subclass of |int|
        having properties for convenient conversion into points or other length units. Likewise,
        the :class:`pptx.util.Pt` class allows convenient specification of point values::

            >>> font.size = Pt(24)
            >>> font.size
            304800
            >>> font.size.pt
            24.0
        """
        sz = self._rPr.sz
        if sz is None:
            return None
        return Centipoints(sz)

    @size.setter
    def size(self, emu: Length | None):
        if emu is None:
            self._rPr.sz = None
        else:
            sz = Emu(emu).centipoints
            self._rPr.sz = sz

    @property
    def strikethrough(self) -> bool | MSO_TEXT_STRIKE_TYPE | None:
        """Indicates the strikethrough setting for this font.

        Value is |True|, |False|, |None|, or a member of the :ref:`MsoTextStrikeType`
        enumeration. |None| is the default and indicates the strikethrough setting should
        be inherited from the style hierarchy, such as from a placeholder. |True|
        indicates single-line strikethrough. |False| indicates no strikethrough. A double
        strikethrough is indicated with `MSO_STRIKE.DOUBLE_LINE`.
        """
        strike = self._rPr.strike
        if strike is MSO_STRIKE.NONE:
            return False
        if strike is MSO_STRIKE.SINGLE_LINE:
            return True
        return strike

    @strikethrough.setter
    def strikethrough(self, value: bool | MSO_TEXT_STRIKE_TYPE | None):
        if value is True:
            value = MSO_STRIKE.SINGLE_LINE
        elif value is False:
            value = MSO_STRIKE.NONE
        self._element.strike = value

    @property
    def underline(self) -> bool | MSO_TEXT_UNDERLINE_TYPE | None:
        """Indicaties the underline setting for this font.

        Value is |True|, |False|, |None|, or a member of the :ref:`MsoTextUnderlineType`
        enumeration. |None| is the default and indicates the underline setting should be inherited
        from the style hierarchy, such as from a placeholder. |True| indicates single underline.
        |False| indicates no underline. Other settings such as double and wavy underlining are
        indicated with members of the :ref:`MsoTextUnderlineType` enumeration.
        """
        u = self._rPr.u
        if u is MSO_UNDERLINE.NONE:
            return False
        if u is MSO_UNDERLINE.SINGLE_LINE:
            return True
        return u

    @underline.setter
    def underline(self, value: bool | MSO_TEXT_UNDERLINE_TYPE | None):
        if value is True:
            value = MSO_UNDERLINE.SINGLE_LINE
        elif value is False:
            value = MSO_UNDERLINE.NONE
        self._element.u = value


class _Hyperlink(Subshape):
    """Text run hyperlink object.

    Corresponds to `a:hlinkClick` child element of the run's properties element (`a:rPr`).
    """

    def __init__(self, rPr: CT_TextCharacterProperties, parent: ProvidesPart):
        super(_Hyperlink, self).__init__(parent)
        self._rPr = rPr

    @property
    def address(self) -> str | None:
        """The URL of the hyperlink.

        Read/write. URL can be on http, https, mailto, or file scheme; others may work.
        """
        if self._hlinkClick is None:
            return None
        return self.part.target_ref(self._hlinkClick.rId)

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        if self._hlinkClick is not None:
            self._remove_hlinkClick()
        if url:
            self._add_hlinkClick(url)

    def _add_hlinkClick(self, url: str):
        rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
        self._rPr.add_hlinkClick(rId)

    @property
    def _hlinkClick(self) -> CT_Hyperlink | None:
        return self._rPr.hlinkClick

    def _remove_hlinkClick(self):
        assert self._hlinkClick is not None
        self.part.drop_rel(self._hlinkClick.rId)
        self._rPr._remove_hlinkClick()  # pyright: ignore[reportPrivateUsage]


class _BulletFormat(object):
    """Proxy object providing read/write access to the bullet of a paragraph.

    Corresponds to the bullet-related choice group of children of an ``a:pPr`` element.
    Not intended to be constructed directly; obtained via :attr:`._Paragraph.bullet`.
    """

    def __init__(self, pPr: CT_TextParagraphProperties):
        super(_BulletFormat, self).__init__()
        self._pPr = pPr

    @property
    def type(self) -> str | None:
        """One of ``"char"``, ``"autonum"``, ``"none"`` or |None|.

        * ``"char"`` — a literal character bullet (``<a:buChar>``)
        * ``"autonum"`` — an auto-numbered bullet (``<a:buAutoNum>``)
        * ``"none"`` — an explicit "no bullet" setting (``<a:buNone>``)
        * |None| — no explicit setting; effective bullet is inherited from the
          paragraph's style hierarchy (layout/master/theme).
        """
        return self._pPr.bullet_type

    @property
    def char(self) -> str | None:
        """The bullet character string, or |None| if the bullet is not a character bullet.

        Returns the value of ``a:buChar/@char`` when present; |None| otherwise.
        """
        return self._pPr.bullet_char

    @property
    def number_scheme(self) -> PP_AUTO_NUMBER_SCHEME | None:
        """The auto-number scheme, or |None| if the bullet is not an autonum bullet.

        Returns a member of :class:`PP_AUTO_NUMBER_SCHEME` when an ``a:buAutoNum`` child
        is present; |None| otherwise.
        """
        return self._pPr.bullet_number_scheme

    @property
    def start_at(self) -> int | None:
        """The ``startAt`` attribute for an autonum bullet, or |None|.

        Returns an integer (defaulting to 1) when an ``a:buAutoNum`` child is present;
        |None| otherwise.
        """
        return self._pPr.bullet_number_start_at

    @property
    def font(self) -> str | None:
        """The bullet-font typeface, or |None| when no ``a:buFont`` child is present.

        This corresponds to the ``typeface`` attribute of the paragraph's
        ``a:buFont`` element. Assigning a string sets the typeface; assigning
        |None| removes any explicit bullet-font setting, causing the bullet
        font to be inherited from the style hierarchy.

        Note that in order to render a character bullet such as a Wingdings
        glyph correctly, the matching bullet font must be set, e.g.
        ``paragraph.bullet.character("§"); paragraph.bullet.font = "Wingdings"``.
        """
        return self._pPr.bullet_font

    @font.setter
    def font(self, value: str | None):
        self._pPr.bullet_font = value

    @property
    def size_pct(self) -> float | None:
        """The bullet size as a fraction of the surrounding text, or |None|.

        When an ``a:buSzPct`` child is present, returns a float in range
        0.25..4.0 (e.g. ``0.75`` for 75%). Returns |None| when no ``a:buSzPct``
        child is present (which may mean an absolute ``a:buSzPts`` is in effect,
        or the bullet size is inherited).

        Assigning a float in range 0.25..4.0 writes ``a:buSzPct`` (replacing any
        existing ``a:buSzPts``). Assigning |None| removes any size element.
        """
        return self._pPr.bullet_size_pct

    @size_pct.setter
    def size_pct(self, value: float | None):
        self._pPr.bullet_size_pct = value

    @property
    def size_points(self) -> Length | None:
        """The bullet size in absolute points, or |None|.

        When an ``a:buSzPts`` child is present, returns a |Length| value (EMU)
        convertible to points via ``.pt``. Returns |None| when no ``a:buSzPts``
        child is present (which may mean an ``a:buSzPct`` is in effect, or the
        bullet size is inherited).

        Assigning a |Length| value (e.g. ``Pt(14)``) writes ``a:buSzPts``
        (replacing any existing ``a:buSzPct``). Assigning |None| removes any
        size element.
        """
        return self._pPr.bullet_size_points

    @size_points.setter
    def size_points(self, value: Length | None):
        self._pPr.bullet_size_points = value

    def clear_size(self) -> _BulletFormat:
        """Remove any explicit bullet-size setting.

        Removes both ``a:buSzPct`` and ``a:buSzPts`` children if present,
        causing the bullet size to be inherited from the style hierarchy.
        Returns self to support chaining.
        """
        self._pPr.clear_bullet_size()
        return self

    @lazyproperty
    def color(self) -> ColorFormat:
        """|ColorFormat| proxy for the bullet color.

        Provides access to the ``a:buClr`` color settings. The underlying
        ``a:buClr`` element is created on first access; ensuring it contains a
        solid RGB color requires assigning a value such as
        ``bullet.color.rgb = RGBColor(0xFF, 0, 0)`` or
        ``bullet.color.theme_color = MSO_THEME_COLOR.ACCENT_1``.

        Use :meth:`.clear_color` to remove the bullet-color setting entirely.
        """
        from pptx.dml.color import ColorFormat as _ColorFormat

        buClr = self._pPr.get_or_add_buClr()
        # -- ensure a concrete color choice is present so caller sees a usable
        # -- ColorFormat; default to black sRGB (caller typically overrides it
        # -- via .rgb or .theme_color).
        if buClr.eg_colorChoice is None:
            buClr.get_or_change_to_srgbClr()
        return _ColorFormat.from_colorchoice_parent(buClr)

    def clear_color(self) -> _BulletFormat:
        """Remove any explicit bullet-color setting.

        Removes any ``a:buClr`` child, causing the bullet color to be inherited
        from the style hierarchy. Returns self to support chaining.
        """
        self._pPr._remove_buClr()  # pyright: ignore[reportPrivateUsage]
        # -- invalidate the lazyproperty cache so a fresh ColorFormat is
        # -- created on the next access (over a newly-added buClr).
        self.__dict__.pop("color", None)
        return self

    def character(self, char: str) -> _BulletFormat:
        """Configure this paragraph to use `char` as its bullet character.

        `char` must be a single-character string. Any existing bullet setting is replaced.
        Returns self to support chaining.
        """
        if not isinstance(char, str) or len(char) != 1:  # pyright: ignore[reportUnnecessaryIsInstance]
            raise ValueError(
                f"`char` must be a single-character string, got {char!r}"
            )
        self._pPr.set_char_bullet(char)
        return self

    def auto_number(
        self,
        scheme: PP_AUTO_NUMBER_SCHEME,
        start_at: int | None = None,
    ) -> _BulletFormat:
        """Configure this paragraph to use `scheme` as its auto-number bullet.

        `scheme` must be a member of :class:`PP_AUTO_NUMBER_SCHEME` (raises |ValueError|
        otherwise). `start_at`, when specified, sets the starting ordinal (must be in
        range 1..32767). Any existing bullet setting is replaced. Returns self to support
        chaining.
        """
        # -- normalise to an enum member; raises ValueError for non-members --
        scheme_member = PP_AUTO_NUMBER_SCHEME(scheme)
        self._pPr.set_auto_number_bullet(scheme_member, start_at)
        return self

    def none(self) -> _BulletFormat:
        """Explicitly suppress any inherited bullet on this paragraph.

        Writes an ``<a:buNone/>`` element, replacing any existing bullet setting.
        Returns self to support chaining.
        """
        self._pPr.set_no_bullet()
        return self

    def clear(self) -> _BulletFormat:
        """Remove any explicit bullet setting, causing the paragraph to inherit.

        After calling this method, :attr:`.type` returns |None| and the paragraph's
        effective bullet is determined by its style hierarchy.
        """
        self._pPr.clear_bullet()
        return self


class _Paragraph(Subshape):
    """Paragraph object. Not intended to be constructed directly."""

    def __init__(self, p: CT_TextParagraph, parent: ProvidesPart):
        super(_Paragraph, self).__init__(parent)
        self._element = self._p = p

    def add_line_break(self):
        """Add line break at end of this paragraph."""
        self._p.add_br()

    def add_run(self) -> _Run:
        """Return a new run appended to the runs in this paragraph."""
        r = self._p.add_r()
        return _Run(r, self)

    @property
    def alignment(self) -> PP_PARAGRAPH_ALIGNMENT | None:
        """Horizontal alignment of this paragraph.

        The value |None| indicates the paragraph should 'inherit' its effective value from its
        style hierarchy. Assigning |None| removes any explicit setting, causing its inherited
        value to be used.
        """
        return self._pPr.algn

    @alignment.setter
    def alignment(self, value: PP_PARAGRAPH_ALIGNMENT | None):
        self._pPr.algn = value

    @lazyproperty
    def bullet(self) -> _BulletFormat:
        """|_BulletFormat| proxy for the bullet properties of this paragraph.

        Provides read/write access to the paragraph-level bullet: one can configure the
        paragraph to use a character bullet (:meth:`._BulletFormat.character`), an
        auto-numbered bullet (:meth:`._BulletFormat.auto_number`), no bullet
        (:meth:`._BulletFormat.none`), or remove any explicit bullet setting so the
        value is inherited (:meth:`._BulletFormat.clear`).
        """
        return _BulletFormat(self._pPr)

    def clear(self):
        """Remove all content from this paragraph.

        Paragraph properties are preserved. Content includes runs, line breaks, and fields.
        """
        for elm in self._element.content_children:
            self._element.remove(elm)
        return self

    @property
    def font(self) -> Font:
        """|Font| object containing default character properties for the runs in this paragraph.

        These character properties override default properties inherited from parent objects such
        as the text frame the paragraph is contained in and they may be overridden by character
        properties set at the run level.
        """
        return Font(self._defRPr)

    @property
    def level(self) -> int:
        """Indentation level of this paragraph.

        Read-write. Integer in range 0..8 inclusive. 0 represents a top-level paragraph and is the
        default value. Indentation level is most commonly encountered in a bulleted list, as is
        found on a word bullet slide.
        """
        return self._pPr.lvl

    @level.setter
    def level(self, level: int):
        self._pPr.lvl = level

    @property
    def line_spacing(self) -> int | float | Length | None:
        """The space between baselines in successive lines of this paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. A numeric value, e.g. `2` or `1.5`,
        indicates spacing is applied in multiples of line heights. A |Length| value such as
        `Pt(12)` indicates spacing is a fixed height. The |Pt| value class is a convenient way to
        apply line spacing in units of points.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.line_spacing

    @line_spacing.setter
    def line_spacing(self, value: int | float | Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.line_spacing = value

    @property
    def runs(self) -> tuple[_Run, ...]:
        """Sequence of runs in this paragraph."""
        return tuple(_Run(r, self) for r in self._element.r_lst)

    @property
    def space_after(self) -> Length | None:
        """The spacing to appear between this paragraph and the subsequent paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. |Length| objects provide convenience
        properties, such as `.pt` and `.inches`, that allow easy conversion to various length
        units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_after

    @space_after.setter
    def space_after(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_after = value

    @property
    def space_before(self) -> Length | None:
        """The spacing to appear between this paragraph and the prior paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. |Length| objects provide convenience
        properties, such as `.pt` and `.cm`, that allow easy conversion to various length units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_before

    @space_before.setter
    def space_before(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_before = value

    @property
    def text(self) -> str:
        """Text of paragraph as a single string.

        Read/write. This value is formed by concatenating the text in each run and field making up
        the paragraph, adding a vertical-tab character (`"\\v"`) for each line-break element
        (`<a:br>`, soft carriage-return) encountered.

        While the encoding of line-breaks as a vertical tab might be surprising at first, doing so
        is consistent with PowerPoint's clipboard copy behavior and allows a line-break to be
        distinguished from a paragraph boundary within the str return value.

        Assignment causes all content in the paragraph to be replaced. Each vertical-tab character
        (`"\\v"`) in the assigned str is translated to a line-break, as is each line-feed
        character (`"\\n"`). Contrast behavior of line-feed character in `TextFrame.text` setter.
        If line-feed characters are intended to produce new paragraphs, use `TextFrame.text`
        instead. Any other control characters in the assigned string are escaped as a hex
        representation like "_x001B_" (for ESC (ASCII 27) in this example).
        """
        return "".join(elm.text for elm in self._element.content_children)

    @text.setter
    def text(self, text: str):
        self.clear()
        self._element.append_text(text)

    @property
    def _defRPr(self) -> CT_TextCharacterProperties:
        """The element that defines the default run properties for runs in this paragraph.

        Causes the element to be added if not present.
        """
        return self._pPr.get_or_add_defRPr()

    @property
    def _pPr(self) -> CT_TextParagraphProperties:
        """Contains the properties for this paragraph.

        Causes the element to be added if not present.
        """
        return self._p.get_or_add_pPr()


class _Run(Subshape):
    """Text run object. Corresponds to `a:r` child element in a paragraph."""

    def __init__(self, r: CT_RegularTextRun, parent: ProvidesPart):
        super(_Run, self).__init__(parent)
        self._r = r

    @property
    def font(self):
        """|Font| instance containing run-level character properties for the text in this run.

        Character properties can be and perhaps most often are inherited from parent objects such
        as the paragraph and slide layout the run is contained in. Only those specifically
        overridden at the run level are contained in the font object.
        """
        rPr = self._r.get_or_add_rPr()
        return Font(rPr)

    @lazyproperty
    def hyperlink(self) -> _Hyperlink:
        """Proxy for any `a:hlinkClick` element under the run properties element.

        Created on demand, the hyperlink object is available whether an `a:hlinkClick` element is
        present or not, and creates or deletes that element as appropriate in response to actions
        on its methods and attributes.
        """
        rPr = self._r.get_or_add_rPr()
        return _Hyperlink(rPr, self)

    @property
    def text(self):
        """Read/write. A unicode string containing the text in this run.

        Assignment replaces all text in the run. The assigned value can be a 7-bit ASCII
        string, a UTF-8 encoded 8-bit string, or unicode. String values are converted to
        unicode assuming UTF-8 encoding.

        Any other control characters in the assigned string other than tab or newline
        are escaped as a hex representation. For example, ESC (ASCII 27) is escaped as
        "_x001B_". Contrast the behavior of `TextFrame.text` and `_Paragraph.text` with
        respect to line-feed and vertical-tab characters.
        """
        return self._r.text

    @text.setter
    def text(self, text: str):
        self._r.text = text
