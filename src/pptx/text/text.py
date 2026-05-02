"""Text-related objects such as TextFrame and Paragraph."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterator, NamedTuple, cast

from pptx.dml.color import ColorFormat as _ColorFormat
from pptx.dml.color import _Color  # pyright: ignore[reportPrivateUsage]  # noqa: PLC2701
from pptx.dml.effect import EffectFormat, ShadowFormat
from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import (
    MSO_AUTO_SIZE,
    MSO_STRIKE,
    MSO_UNDERLINE,
    MSO_VERTICAL_ANCHOR,
    PP_AUTO_NUMBER_SCHEME,
)
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_TextWrappingType
from pptx.oxml.text import CT_RegularTextRun, CT_TextField, CT_TextLineBreak
from pptx.shapes import Subshape
from pptx.text.fonts import FontFiles
from pptx.text.layout import TextFitter
from pptx.util import Centipoints, Emu, Length, Pt, lazyproperty

if TYPE_CHECKING:
    from pptx.dml.color import ColorFormat, RGBColor
    from pptx.enum.dml import MSO_THEME_COLOR
    from pptx.enum.text import (
        MSO_TEXT_STRIKE_TYPE,
        MSO_TEXT_UNDERLINE_TYPE,
        MSO_VERTICAL_ANCHOR,
        PP_PARAGRAPH_ALIGNMENT,
    )
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.text import (
        CT_TextBody,
        CT_TextCharacterProperties,
        CT_TextParagraph,
        CT_TextParagraphProperties,
    )
    from pptx.parts.slide import SlidePart
    from pptx.slide import Slide
    from pptx.types import ProvidesExtents, ProvidesPart


class TextFrameRect(NamedTuple):
    """Slide-relative rectangle PowerPoint allocates for rendering a shape's text.

    A 4-tuple ``(left, top, width, height)`` in English Metric Units (EMU).
    Each element is a |Length|, so ``.inches``, ``.pt``, ``.cm``, and ``.emu``
    are available as on any other length value. See
    :attr:`.BaseShape.text_frame_rect` for how the value is computed and how
    it differs from the shape's bounding box.

    .. versionadded:: 2026.05.0
    """

    left: Length
    top: Length
    width: Length
    height: Length


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
    def font_scale(self) -> float:
        """Font-scale percent applied to text in this text frame by autofit.

        Corresponds to the ``fontScale`` attribute of the ``a:normAutofit`` child of the
        ``a:bodyPr`` element. Returns a float percent in the range 1.0..100.0.

        A value of ``100.0`` (the default) indicates that the text has not been scaled
        down to fit. Lower values (e.g. ``85.0``) indicate PowerPoint has reduced the
        rendered font-size to that percent of its authored size in order to fit the
        text within the shape's bounds. Returns ``100.0`` when no ``a:normAutofit``
        child is present.

        Assigning a value adds an ``a:normAutofit`` child to ``a:bodyPr`` if one is
        not already present, replacing any other autofit-choice child (``a:noAutofit``
        or ``a:spAutoFit``). Useful for placeholder text that does not re-flow until
        the user edits it; emitting an explicit ``fontScale`` lets PowerPoint render
        the reduction on first display.

        Acceptable values are in the range 1.0..100.0. See issue #715.
        """
        return self._bodyPr.font_scale

    @font_scale.setter
    def font_scale(self, value: float):
        self._bodyPr.font_scale = value

    @property
    def line_space_reduction(self) -> float:
        """Line-spacing reduction percent applied by autofit.

        Corresponds to the ``lnSpcReduction`` attribute of the ``a:normAutofit`` child
        of the ``a:bodyPr`` element. Returns a float percent in the range 0.0..100.0.

        A value of ``0.0`` (the default) indicates no reduction. A value of ``20.0``
        indicates PowerPoint is reducing line spacing by 20% of its authored value.
        Returns ``0.0`` when no ``a:normAutofit`` child is present.

        Assigning a value adds an ``a:normAutofit`` child to ``a:bodyPr`` if one is
        not already present, replacing any other autofit-choice child (``a:noAutofit``
        or ``a:spAutoFit``).

        Acceptable values are in the range 0.0..100.0. See issue #715.
        """
        return self._bodyPr.line_space_reduction

    @line_space_reduction.setter
    def line_space_reduction(self, value: float):
        self._bodyPr.line_space_reduction = value

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

    def replace_text(self, find: str, replace: str) -> int:
        """Replace every occurrence of `find` in this text-frame with `replace`.

        Searches across runs so that a keyword broken across multiple `a:r`
        elements by PowerPoint edits (a common occurrence) is still replaced.
        Matches that would cross a line-break (`a:br`) or auto-refresh field
        (`a:fld`) boundary are not replaced — search is effectively scoped to
        each maximal consecutive group of `a:r` runs within a paragraph.

        When a match spans multiple runs, the formatting of the run in which
        the match starts is preserved for the replacement text. Any run fully
        contained within the match is removed; if a match ends partway
        through a run, that run's surviving suffix keeps its original
        formatting. Replacements do not overlap and occur left-to-right.

        Returns the number of replacements performed. `find` must be a
        non-empty string (raises `ValueError` otherwise). See issue #836.

        .. versionadded:: 2026.05.0
        """
        return sum(p.replace_text(find, replace) for p in self.paragraphs)

    @property
    def rotation(self) -> float:
        """Clockwise rotation in degrees applied to the text within this text frame.

        Read/write. Corresponds to the ``rot`` attribute of the ``a:bodyPr``
        element and rotates the text *inside* the shape (the shape itself is
        unaffected). Distinct from :attr:`Shape.rotation`, which rotates the
        whole shape via ``p:spPr/a:xfrm/@rot``.

        Returns ``0.0`` (the default) when the attribute is not present.
        Negative values assigned (e.g. ``-90``) are normalized to the
        equivalent positive rotation in the range ``[0, 360)``. Both integer
        and float values are accepted; PowerPoint stores the value in
        60000ths of a degree under the hood. See issue #133.
        """
        return self._bodyPr.rot

    @rotation.setter
    def rotation(self, value: float):
        self._bodyPr.rot = value

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

    def __init__(
        self,
        rPr: CT_TextCharacterProperties,
        parent: ProvidesPart | None = None,
    ):
        super(Font, self).__init__()
        self._element = self._rPr = rPr
        # -- retained only to support `effective_color`, which needs to walk
        # -- from the run up to the enclosing slide/layout/master parts; other
        # -- entry points (paragraph defRPr, chart text, etc.) can safely omit
        # -- it and get a None result from `effective_color`.
        self._parent = parent

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
        """The |ColorFormat| instance that provides access to the color settings for this font.

        Reading the returned |ColorFormat| (e.g. ``font.color.type``,
        ``font.color.rgb``) does *not* mutate the underlying XML: if no
        ``<a:solidFill>`` is present on the run's ``<a:rPr>``, the color
        appears as "not set" (``type`` is |None|) and the run continues to
        inherit its color from the style hierarchy (see issue #1111).
        Assigning ``font.color.rgb = ...`` or ``font.color.theme_color = ...``
        will create the ``<a:solidFill>`` on first write.
        """
        if self.fill.type == MSO_FILL.SOLID:
            return self.fill.fore_color
        return _FontColorFormat(self)

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this font.

        Provides access to fill properties such as fill color.
        """
        return FillFormat.from_fill_parent(self._rPr, self._parent)

    @lazyproperty
    def highlight_color(self) -> ColorFormat:
        """|ColorFormat| proxy for the text-highlight (text-background) color.

        Corresponds to `a:rPr/a:highlight` — the "text highlighting" swatch
        PowerPoint exposes on the Home ribbon as the text background-color
        marker. A |ColorFormat| is always returned, whether or not an
        `a:highlight` element is already present; the `a:highlight` element is
        created on first access and populated with a default black `a:srgbClr`
        when it has no color-choice child, so callers immediately see a usable
        :attr:`ColorFormat.rgb` / :attr:`ColorFormat.theme_color` surface.

        Typical use::

            run.font.highlight_color.rgb = RGBColor(0xFF, 0xFF, 0x00)
            run.font.highlight_color.theme_color = MSO_THEME_COLOR.ACCENT_1

        To remove an explicit highlight entirely (so the run inherits — which
        in practice means "no highlight"), call
        :meth:`clear_highlight_color`. See issue #675.
        """
        from pptx.dml.color import ColorFormat as _ColorFormat

        highlight = self._rPr.get_or_add_highlight()
        # -- ensure a concrete color-choice is present so callers get a usable
        # -- ColorFormat immediately; default to a black a:srgbClr that will be
        # -- overwritten by the caller's `.rgb = ...` / `.theme_color = ...`.
        # -- (same shape as `_BulletFormat.color` below; descriptor-typed
        # -- attrs aren't visible to pyright strict without extra hints.)
        if highlight.eg_colorChoice is None:  # pyright: ignore[reportUnnecessaryComparison]
            highlight.get_or_change_to_srgbClr()  # pyright: ignore[reportAttributeAccessIssue, reportUnknownMemberType]
        return _ColorFormat.from_colorchoice_parent(  # pyright: ignore[reportUnknownMemberType]
            highlight
        )

    def clear_highlight_color(self) -> Font:
        """Remove any explicit text-highlight color from this run.

        Removes any `a:highlight` child under the run's `a:rPr`, causing the
        text to render with no highlight (the inherited default). Returns
        self to support chaining. Safe to call when no `a:highlight` is
        present. See issue #675.
        """
        self._rPr._remove_highlight()  # pyright: ignore[reportPrivateUsage]
        # -- invalidate the lazyproperty cache so a fresh ColorFormat is
        # -- created on the next access (over a newly-added a:highlight).
        self.__dict__.pop("highlight_color", None)
        return self

    @lazyproperty
    def effect_format(self) -> EffectFormat:
        """|EffectFormat| instance providing access to visual effects on this font.

        Exposes the full `a:effectLst` family (`shadow`, `glow`, `reflection`,
        `soft_edge`) on this run's `a:rPr/a:effectLst`. See issue #546.

        .. versionadded:: 2026.05.0
        """
        return EffectFormat(self._rPr)

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """|ShadowFormat| instance providing access to text-shadow settings.

        Text-shadow is stored on the run's `a:rPr/a:effectLst/a:outerShdw`. The
        returned object exposes the full :class:`~pptx.dml.effect.ShadowFormat`
        API (``inherit``, ``blur_radius``, ``distance``, ``direction``,
        ``color``). A |ShadowFormat| object is always returned, even when no
        shadow is explicitly defined on this run (i.e. the run inherits its
        shadow from the style hierarchy). See issue #546.

        .. versionadded:: 2026.05.0
        """
        return ShadowFormat(self._rPr)

    @property
    def effective_color(self) -> RGBColor | None:
        """Resolved |RGBColor| for this run, walking the inheritance chain.

        Returns the RGB PowerPoint would render for this run's text color,
        tracing the style hierarchy until an explicit color is found:

        1. The run's own `a:rPr/a:solidFill`
        2. The enclosing paragraph's `a:pPr/a:defRPr/a:solidFill`
        3. The enclosing text body's `a:lstStyle/a:lvl{N}pPr/a:defRPr/a:solidFill`
           (level-matched to the paragraph's `@lvl`)
        4. The slide master's `p:txStyles` (``bodyStyle`` / ``titleStyle`` /
           ``otherStyle``) for the matching paragraph level
        5. `None` when no color can be resolved (e.g., missing theme context)

        Any ``<a:lumMod>`` / ``<a:lumOff>`` siblings on the resolved color
        element are applied via :meth:`ColorFormat.to_rgb`, and scheme colors
        are resolved against :attr:`SlideMaster.theme_colors`.

        Returns |None| when this |Font| was constructed without a part-aware
        parent (e.g., obtained from a chart `a:defRPr`), when the inheritance
        walk can't reach a slide-master, or when a scheme color is encountered
        but the target theme entry is missing. Callers that need finer-grained
        control can use :attr:`color` + :meth:`ColorFormat.to_rgb` directly.

        .. versionadded:: 2026.05.0
        """
        theme_colors = self._theme_colors
        for rPr_like in self._iter_inheritance_rPrs():
            rgb = _resolve_solid_fill_rgb(rPr_like, theme_colors)
            if rgb is not None:
                return rgb
        return None

    @property
    def effective_size(self) -> Length | None:
        """Resolved font size for this run, walking the inheritance chain.

        Returns the |Length| (in EMU) PowerPoint would render for this run's
        text, tracing the style hierarchy until an explicit ``sz`` attribute
        is found:

        1. The run's own ``a:rPr/@sz``
        2. The enclosing paragraph's ``a:pPr/a:defRPr/@sz``
        3. The enclosing text body's ``a:lstStyle/a:lvl{N}pPr/a:defRPr/@sz``
           (level-matched to the paragraph's ``@lvl``)
        4. The slide master's ``p:txStyles`` (``bodyStyle`` / ``otherStyle`` /
           ``titleStyle``) for the matching paragraph level
        5. The presentation's ``p:defaultTextStyle`` for the matching level
        6. |None| when no explicit size can be resolved (PowerPoint will
           typically fall back to its own default, around 18pt)

        Returns |None| when this |Font| was constructed without a part-aware
        parent (e.g., obtained from a chart ``a:defRPr``), when the
        inheritance walk can't reach a slide master, or when no ancestor in
        the chain declares an explicit size. Addresses issue #378.

        .. versionadded:: 2026.05.0
        """
        for rPr_like in self._iter_inheritance_rPrs():
            sz = rPr_like.get("sz")
            if sz is not None:
                try:
                    return Centipoints(int(sz))
                except (TypeError, ValueError):
                    continue
        return None

    @property
    def effective_bold(self) -> bool | None:
        """Resolved bold setting for this run, walking the inheritance chain.

        Returns |True| / |False| using the same chain described for
        :attr:`effective_size` (run rPr → paragraph defRPr → lstStyle →
        master ``p:txStyles`` → presentation ``p:defaultTextStyle``). Returns
        |None| when no ancestor in the chain declares an explicit ``b``
        attribute, the |Font| was constructed without a part-aware parent,
        or the walk can't reach a slide master. See issue #378.

        .. versionadded:: 2026.05.0
        """
        return self._effective_bool_attr("b")

    @property
    def effective_italic(self) -> bool | None:
        """Resolved italic setting for this run, walking the inheritance chain.

        Returns |True| / |False| using the same chain described for
        :attr:`effective_size`. Returns |None| when no ancestor declares an
        explicit ``i`` attribute. See issue #378.

        .. versionadded:: 2026.05.0
        """
        return self._effective_bool_attr("i")

    @property
    def effective_name(self) -> str | None:
        """Resolved Latin typeface name for this run, walking the chain.

        Returns the ``a:latin/@typeface`` value PowerPoint would render for
        this run's Latin-script text, using the same walk as
        :attr:`effective_size`. Returns |None| when no ancestor declares an
        explicit ``a:latin`` child (PowerPoint falls back to the theme's
        ``a:majorFont`` / ``a:minorFont`` in that case). See issue #378.

        .. versionadded:: 2026.05.0
        """
        for rPr_like in self._iter_inheritance_rPrs():
            latin = rPr_like.find(qn("a:latin"))
            if latin is not None:
                typeface = latin.get("typeface")
                if typeface:
                    return typeface
        return None

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

    @property
    def use_theme_hyperlink_color(self) -> bool | None:
        """Whether the theme hyperlink color applies to this run.

        Read-write tri-state. Addresses issue #940: in PowerPoint the theme
        color-scheme's `a:hlink` / `a:folHlink` color overrides any explicit
        `a:solidFill` color set on a run that carries an `a:hlinkClick`, making
        it impossible to recolor a hyperlinked run from python-pptx alone.

        Values:

        - |None| -- No override marker is present (default). The theme's
          hyperlink color applies in PowerPoint as usual.
        - |True| -- Same XML shape as |None| (no marker). Any pre-existing
          opt-out marker is cleared. Provided for symmetry.
        - |False| -- Writes a python-pptx-owned `a:extLst/a:ext` marker under
          the run's `a:hlinkClick` element to record the caller's intent that
          the run's explicit color be preferred. The marker round-trips
          through python-pptx and is ignored by PowerPoint.

        Setting the flag has no effect when the run does not yet carry a
        hyperlink; call `Run.hyperlink.address = ...` first. PowerPoint
        honors the run's explicit color only if the theme's hyperlink color
        is also replaced -- see
        :ref:`hyperlink-color <text-hyperlink-color-guide>` in the user guide
        for the recommended pattern.
        """
        hlinkClick = self._rPr.hlinkClick
        if hlinkClick is None:
            return None
        return False if hlinkClick.suppress_theme_color else None

    @use_theme_hyperlink_color.setter
    def use_theme_hyperlink_color(self, value: bool | None):
        hlinkClick = self._rPr.hlinkClick
        if hlinkClick is None:
            # -- no-op when the run has no hyperlink; nothing to annotate --
            return
        # -- `None` and `True` both mean "no override marker" --
        hlinkClick.suppress_theme_color = value is False

    @property
    def _slide_master(self):
        """The |SlideMaster| reachable from this Font's parent, or |None|.

        Used by :attr:`effective_color` to resolve scheme colors and walk up
        into `p:txStyles`. Returns |None| when no parent part is bound (e.g.
        the Font wraps a chart ``a:defRPr``) or when the containing part is
        not a slide / layout / master.
        """
        if self._parent is None:
            return None
        try:
            part = self._parent.part
        except AttributeError:
            return None
        # -- delayed to avoid circular import --
        from pptx.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart

        if isinstance(part, SlideMasterPart):
            return part.slide_master
        if isinstance(part, SlideLayoutPart):
            return part.slide_master
        if isinstance(part, SlidePart):
            return part.slide_layout.slide_master
        return None

    @property
    def _presentation_elm(self) -> Any:
        """The `p:presentation` element reachable from this Font, or |None|.

        Returns the root XML element of the presentation part (which carries
        the ``p:defaultTextStyle`` consulted by :attr:`effective_size` et al.)
        when the parent graph makes it reachable. Returns |None| for a Font
        constructed without a part-aware parent, or when the package has no
        presentation part bound.
        """
        if self._parent is None:
            return None
        try:
            part = self._parent.part
        except AttributeError:
            return None
        try:
            prs_part = part.package.presentation_part
        except AttributeError:
            return None
        if prs_part is None:
            return None
        try:
            return prs_part.element
        except AttributeError:
            return None

    @property
    def _theme_colors(self):
        """Mapping of scheme-color name to |RGBColor|, or |None|."""
        master = self._slide_master
        if master is None:
            return None
        return master.theme_colors

    def _iter_inheritance_rPrs(self) -> Iterator[Any]:
        """Yield rPr-like elements in style-inheritance order.

        Each yielded element is an ``a:rPr`` or ``a:defRPr`` carrying the
        character properties that PowerPoint consults for this run, ordered
        from most specific to least specific:

        1. The run's own ``a:rPr``
        2. The enclosing paragraph's ``a:pPr/a:defRPr``
        3. The enclosing text body's ``a:lstStyle/a:lvl{N}pPr/a:defRPr``
           (level-matched to the paragraph)
        4. The slide master's ``p:txStyles/p:{body,other,title}Style
           /a:lvl{N}pPr/a:defRPr``
        5. The presentation's ``p:defaultTextStyle/a:lvl{N}pPr/a:defRPr``

        Consumers that stop at the first match implement the inheritance
        semantics for a specific attribute (e.g. ``@sz``, ``@b``,
        ``a:latin``). Used by :attr:`effective_color`,
        :attr:`effective_size`, :attr:`effective_bold`, :attr:`effective_italic`,
        and :attr:`effective_name`.
        """
        # -- 1. direct run rPr --
        yield self._rPr

        # -- find enclosing a:p --
        p = self._rPr.getparent()
        while p is not None and p.tag != qn("a:p"):
            p = p.getparent()

        lvl = _paragraph_level(p) if p is not None else 0

        # -- 2. paragraph pPr/defRPr --
        if p is not None:
            pPr = p.find(qn("a:pPr"))
            if pPr is not None:
                defRPr = pPr.find(qn("a:defRPr"))
                if defRPr is not None:
                    yield defRPr

            # -- 3. txBody lstStyle/lvlNpPr/defRPr --
            txBody = p.getparent()
            if txBody is not None:
                lstStyle = txBody.find(qn("a:lstStyle"))
                if lstStyle is not None:
                    defRPr = _lvl_defRPr(lstStyle, lvl)
                    if defRPr is not None:
                        yield defRPr

        # -- 4. master p:txStyles --
        master = self._slide_master
        if master is not None:
            master_elm = master._element  # pyright: ignore[reportPrivateUsage]
            txStyles = master_elm.find(qn("p:txStyles"))
            if txStyles is not None:
                for style_tag in ("p:bodyStyle", "p:otherStyle", "p:titleStyle"):
                    style = txStyles.find(qn(style_tag))
                    if style is None:
                        continue
                    defRPr = _lvl_defRPr(style, lvl)
                    if defRPr is not None:
                        yield defRPr

        # -- 5. presentation p:defaultTextStyle --
        prs_elm = self._presentation_elm
        if prs_elm is not None:
            default_style = prs_elm.find(qn("p:defaultTextStyle"))
            if default_style is not None:
                defRPr = _lvl_defRPr(default_style, lvl)
                if defRPr is not None:
                    yield defRPr

    def _effective_bool_attr(self, attr_name: str) -> bool | None:
        """Return the first explicit boolean value of `attr_name` in the chain.

        `attr_name` is the local-name of an OOXML boolean attribute on
        ``a:rPr`` / ``a:defRPr`` — typically ``"b"`` (bold) or ``"i"``
        (italic). Returns |None| when no ancestor in the inheritance chain
        declares the attribute.
        """
        for rPr_like in self._iter_inheritance_rPrs():
            raw = rPr_like.get(attr_name)
            if raw is None:
                continue
            # -- OOXML boolean lexical space: "1"/"true" / "0"/"false" --
            if raw in ("1", "true"):
                return True
            if raw in ("0", "false"):
                return False
        return None


class _FontColorFormat(_ColorFormat):
    """|ColorFormat| that defers creating `<a:solidFill>` until first write.

    Reading attributes (`type`, `rgb`, `theme_color`, ...) reports the "no
    color set" state (|None| / `MSO_COLOR_TYPE.NOT_THEME_COLOR`) without
    writing any XML, so the run's color-inheritance chain is preserved. The
    first `rgb = ...` or `theme_color = ...` assignment promotes the run's
    fill to `a:solidFill` and delegates to the normal |ColorFormat|
    machinery. See issue #1111.
    """

    def __init__(self, font: "Font"):
        self._font = font
        # -- populate parent slots so read-side methods on ColorFormat /
        # -- _NoneColor work without any mutation. `_xFill` here is the rPr:
        # -- the read-side never dispatches on it, and the write-side is
        # -- overridden below to promote the fill before touching it.
        super().__init__(font._rPr, _Color(None))  # pyright: ignore[reportPrivateUsage]

    def _promote(self) -> None:
        """Ensure the run's fill is `a:solidFill` and re-sync state.

        Creates the `a:solidFill` child on first call and replaces the
        inner color + fill-parent bindings so subsequent reads/writes
        behave like a regular |ColorFormat|.
        """
        font = self._font
        if font.fill.type != MSO_FILL.SOLID:
            font.fill.solid()
        fresh = font.fill.fore_color
        # -- replace the cached `color` lazyproperty on `font` so later
        # -- `font.color` accesses return the newly-created solid-fill form.
        font.__dict__["color"] = fresh
        # -- `_xFill` / `_color` are private slots on the base ColorFormat;
        # -- the base class is untyped so these assignments are "Unknown"
        # -- from pyright's perspective.
        self._xFill = fresh._xFill  # pyright: ignore[reportUnknownMemberType]
        self._color = fresh._color  # pyright: ignore[reportUnknownMemberType]

    # -- Setter-only overrides: promote the fill, then delegate to the base
    # -- class's own setter via `fset`. The property getter is re-declared
    # -- so pyright-strict accepts the override-setter pairing.

    @property
    def rgb(self) -> "RGBColor | None":
        return self._color.rgb  # pyright: ignore[reportUnknownMemberType, reportUnknownVariableType]

    @rgb.setter
    def rgb(self, rgb: "RGBColor") -> None:
        self._promote()
        _ColorFormat.rgb.fset(self, rgb)  # pyright: ignore[reportOptionalCall]

    @property
    def theme_color(self) -> "MSO_THEME_COLOR":
        return self._color.theme_color  # pyright: ignore[reportUnknownMemberType, reportUnknownVariableType]

    @theme_color.setter
    def theme_color(self, mso_theme_color_idx: "MSO_THEME_COLOR") -> None:
        self._promote()
        _ColorFormat.theme_color.fset(self, mso_theme_color_idx)  # pyright: ignore[reportOptionalCall]


def _resolve_solid_fill_rgb(rPr_like, theme_colors):
    """Return resolved |RGBColor| for `rPr_like/a:solidFill`, or |None|.

    `rPr_like` is any element that may carry an `a:solidFill` child (`a:rPr`,
    `a:defRPr`, or similar). `theme_colors` is a scheme-color mapping as
    returned by :attr:`SlideMaster.theme_colors`, or |None|; it is required
    only when the fill references a scheme color. Returns |None| when the
    element has no `a:solidFill`, or when a scheme color can't be resolved.
    """
    if rPr_like is None:
        return None
    solidFill = rPr_like.find(qn("a:solidFill"))
    if solidFill is None:
        return None
    # -- solidFill has exactly one color-choice child --
    color_elm = next(iter(solidFill), None)
    if color_elm is None:
        return None
    # -- delayed to avoid circular import --
    from pptx.dml.color import _Color

    color = _Color(color_elm)
    try:
        return color.to_rgb(theme_colors)
    except (ValueError, NotImplementedError):
        return None


def _paragraph_level(p_elm):
    """Return the paragraph's `a:pPr/@lvl`, defaulting to 0."""
    pPr = p_elm.find(qn("a:pPr"))
    if pPr is None:
        return 0
    try:
        return int(pPr.get("lvl", "0"))
    except (TypeError, ValueError):
        return 0


def _iter_paragraph_text_chunks(p_elm: Any) -> Iterator[str]:
    """Yield text chunks for `p_elm` in document order.

    Walks the direct children of the `a:p` element. For each `a:r`, `a:br`,
    or `a:fld` child, yields that element's ``.text`` (which already encodes
    line-breaks as vertical-tab). For each `mc:AlternateContent` child, drops
    into its `mc:Fallback` subtree and yields the text of every `a:r`, `a:br`,
    and `a:fld` descendant there. The `mc:Choice` side is intentionally
    ignored -- it typically holds an `a14:m/m:oMath` equation whose rendered
    text is already reproduced under `mc:Fallback`, and yielding both would
    double-count.

    See issue #947: PowerPoint emits an inline math equation as an
    ``mc:AlternateContent`` child of the enclosing ``a:p`` with a plain-text
    ``mc:Fallback`` rendering; ``_Paragraph.text`` must surface that
    rendering so callers doing text extraction see the equation's characters
    rather than an empty string.
    """
    ac_tag = qn("mc:AlternateContent")
    fb_tag = qn("mc:Fallback")
    # -- first-class text children, matching `CT_TextParagraph.content_children`.
    text_types = (CT_RegularTextRun, CT_TextLineBreak, CT_TextField)

    for child in p_elm:
        if isinstance(child, text_types):
            yield child.text
        elif child.tag == ac_tag:
            fallback = child.find(fb_tag)
            if fallback is None:
                continue
            # -- yield text of every a:r / a:br / a:fld descendant of the
            # -- fallback subtree, preserving document order via `iter()`.
            for descendant in fallback.iter():
                if isinstance(descendant, text_types):
                    yield descendant.text


def _lvl_defRPr(style_elm, lvl):
    """Return the `a:defRPr` for paragraph `lvl` from a list-style-like element.

    `style_elm` is an `a:lstStyle` or one of the master `p:txStyles` children
    (`p:titleStyle`, `p:bodyStyle`, `p:otherStyle`), each of which contains
    `a:defPPr` and `a:lvl{1..9}pPr` children whose `a:defRPr` holds the
    default run properties for that indent level. `lvl` is 0-based
    (0 → `a:lvl1pPr`, 1 → `a:lvl2pPr`, ...). Returns |None| when the
    requested level is absent.
    """
    lvl_tag = "a:lvl{}pPr".format(lvl + 1)
    lvl_pPr = style_elm.find(qn(lvl_tag))
    if lvl_pPr is None:
        # -- lvl1pPr is sometimes written as a:defPPr on the master styles --
        if lvl == 0:
            defPPr = style_elm.find(qn("a:defPPr"))
            if defPPr is not None:
                return defPPr.find(qn("a:defRPr"))
        return None
    return lvl_pPr.find(qn("a:defRPr"))


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
        Returns |None| when no hyperlink is defined, or when the hyperlink is an internal
        slide-jump rather than an external URL -- see :attr:`target_slide`.
        """
        hlinkClick = self._hlinkClick
        if hlinkClick is None:
            return None
        rId = hlinkClick.rId
        # -- a slide-jump hyperlink (ppaction://hlinksldjump) uses an internal
        # -- relationship and has no external URL. Return None in that case so
        # -- `.address` stays strictly external-URL. Use `.target_slide` to
        # -- access the slide-jump target.
        if not rId:
            return None
        if hlinkClick.action == "ppaction://hlinksldjump":
            return None
        return self.part.target_ref(rId)

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        if self._hlinkClick is not None:
            self._remove_hlinkClick()
        if url:
            self._add_hlinkClick(url)

    @property
    def target_slide(self) -> "Slide | None":
        """The |Slide| this run jumps to when clicked, or |None|.

        Read/write. When the run's ``a:hlinkClick`` carries
        ``action="ppaction://hlinksldjump"`` and targets another slide in the
        presentation, returns that |Slide|. Returns |None| when the run has no
        hyperlink, when the hyperlink is an external URL, or when any other
        click action is present.

        Assigning a |Slide| adds (or replaces) an ``a:hlinkClick`` that
        navigates to that slide during a slide show -- equivalent to
        PowerPoint's "Insert > Hyperlink > Place in This Document". Assigning
        |None| removes any hyperlink on the run (whether slide-jump or URL);
        use :attr:`address` to set an external URL.

        .. versionadded:: 2026.05.0
        """
        hlinkClick = self._hlinkClick
        if hlinkClick is None:
            return None
        if hlinkClick.action != "ppaction://hlinksldjump":
            return None
        rId = hlinkClick.rId
        if not rId:
            return None
        slide_part = cast("SlidePart", self.part.related_part(rId))
        return slide_part.slide

    @target_slide.setter
    def target_slide(self, slide: "Slide | None"):
        # -- always clear any existing hlinkClick first; this mirrors
        # -- ActionSetting._clear_click_action so swapping from URL to
        # -- slide-jump (or vice-versa) correctly drops the prior rel.
        if self._hlinkClick is not None:
            self._remove_hlinkClick()
        if slide is None:
            return
        rId = self.part.relate_to(slide.part, RT.SLIDE)
        hlinkClick = self._rPr.get_or_add_hlinkClick()
        hlinkClick.action = "ppaction://hlinksldjump"
        hlinkClick.rId = rId

    @target_slide.deleter
    def target_slide(self):
        # -- `del run.hyperlink.target_slide` is a convenient alias for
        # -- assigning None; it removes any hyperlink on the run.
        self.target_slide = None

    def _add_hlinkClick(self, url: str):
        rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
        self._rPr.add_hlinkClick(rId)

    @property
    def _hlinkClick(self) -> CT_Hyperlink | None:
        return self._rPr.hlinkClick

    def _remove_hlinkClick(self):
        assert self._hlinkClick is not None
        rId = self._hlinkClick.rId
        if rId:
            self.part.drop_rel(rId)
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

    def add_field(self, field_type: str, text: str = "") -> _Field:
        """Append an auto-refresh field (`a:fld`) to the end of this paragraph.

        `field_type` is the `a:fld/@type` string PowerPoint uses to classify the field. Common
        values are `"slidenum"` (current slide number), `"datetime"`, `"datetimeFigureOut"`, or
        one of the formatted date variants `"datetime1"` … `"datetime13"`, and `"footer"`.
        `text` is the literal string stored alongside the field; PowerPoint uses this as the
        displayed value until the field is first refreshed on open — for a slide-number field,
        a short placeholder such as `"#"` is customary.
        """
        fld = self._p.add_fld(field_type, text)
        return _Field(fld, self)

    def add_line_break(self):
        """Add line break at end of this paragraph."""
        self._p.add_br()

    def add_math_equation(self, omml_xml: str) -> None:
        """Append an Office-Math (OMML) equation to this paragraph.

        `omml_xml` is a string of caller-provided OMML -- typically the output of Microsoft's
        ``MML2OMML.XSL`` transform (converting MathML to OMML) or another OMML producer. Its
        root element must be ``m:oMath`` or ``m:oMathPara`` and must declare the ``m``
        namespace (``http://schemas.openxmlformats.org/officeDocument/2006/math``) on the root.

        The OMML fragment is wrapped in the ``mc:AlternateContent/mc:Choice[Requires="a14"]
        /a14:m`` scaffolding PowerPoint emits for an equation embedded inline in a paragraph,
        paired with an ``mc:Fallback`` run that carries the OMML reduced to its visible text
        (concatenated ``m:t`` children) so pre-2010 consumers render *something* readable. Any
        existing runs, line-breaks, or fields already on the paragraph are preserved -- the
        equation is appended to the end of its content (before any ``a:endParaRPr``).

        The companion read-side is :attr:`BaseShape.math_equation_xml` / :attr:`has_math_equation`
        (see issue #126). Converting between OMML and LaTeX / MathML is **not** in scope -- the
        caller is responsible for producing the OMML. Raises ``ValueError`` when ``omml_xml`` is
        not well-formed XML or its root is neither ``m:oMath`` nor ``m:oMathPara``.
        """
        self._p.add_math_equation(omml_xml)

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

    def delete(self) -> None:
        """Remove this paragraph from its containing text frame.

        The paragraph's `a:p` element is removed from its parent `p:txBody` (or
        `a:txBody` in the case of a table cell). PowerPoint requires every text
        frame to contain at least one `a:p` child, so when this paragraph is the
        last one in the text frame a fresh empty `a:p` is added in its place to
        preserve that invariant. Subsequent use of this paragraph object is
        undefined; most operations will raise an exception. See issue #144.
        """
        txBody = self._p.getparent()
        if txBody is None:
            return
        txBody.remove(self._p)
        # -- PowerPoint requires a text frame to contain at least one `a:p`;
        # -- add a fresh empty paragraph when the deletion emptied the body.
        if not txBody.findall(qn("a:p")):
            txBody.add_p()  # pyright: ignore[reportAttributeAccessIssue]

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

    def replace_text(self, find: str, replace: str) -> int:
        """Replace every occurrence of `find` in this paragraph with `replace`.

        Searches across the paragraph's runs so that a keyword split across
        multiple `a:r` elements (for example after a PowerPoint edit that
        broke ``"{NAME}"`` into two runs) is still replaced. Matches that
        would cross an `a:br` (line-break) or `a:fld` (auto-refresh field)
        boundary are not replaced — search is scoped to each maximal
        consecutive group of `a:r` runs.

        When a match spans multiple runs, the formatting of the run in which
        the match starts is preserved for the replacement text; runs fully
        contained within the match are removed, and when a match ends partway
        through a run the surviving suffix keeps its original formatting.
        Replacements do not overlap and occur left-to-right.

        Returns the number of replacements performed. Raises `ValueError` if
        `find` is an empty string. See issue #836.
        """
        if not find:
            raise ValueError("`find` must be a non-empty string")

        # -- partition content children into maximal consecutive a:r groups,
        # -- separated by a:br / a:fld boundaries (within which replacement
        # -- is intentionally not attempted).
        groups: list[list[CT_RegularTextRun]] = [[]]
        for child in self._element.content_children:
            if isinstance(child, CT_RegularTextRun):
                groups[-1].append(child)
            else:
                if groups[-1]:
                    groups.append([])
        count = 0
        for run_group in groups:
            if run_group:
                count += _replace_in_runs(run_group, find, replace)
        return count

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

        When the paragraph contains an `mc:AlternateContent` wrapper (e.g. an inline OMML math
        equation), the text inside its `mc:Fallback` subtree is included in document order, so
        PowerPoint's downgrade-friendly plain-text rendering of the equation is surfaced for
        callers that just want a readable string. See issue #947.

        Assignment causes all content in the paragraph to be replaced. Each vertical-tab character
        (`"\\v"`) in the assigned str is translated to a line-break, as is each line-feed
        character (`"\\n"`). Contrast behavior of line-feed character in `TextFrame.text` setter.
        If line-feed characters are intended to produce new paragraphs, use `TextFrame.text`
        instead. Any other control characters in the assigned string are escaped as a hex
        representation like "_x001B_" (for ESC (ASCII 27) in this example).
        """
        return "".join(_iter_paragraph_text_chunks(self._element))

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

    def delete(self) -> None:
        """Remove this run from its containing paragraph.

        The run's `a:r` element is removed from its parent `a:p`. Other content
        (line-breaks, fields, sibling runs) is preserved. Paragraph-level
        properties (``a:pPr``, ``a:endParaRPr``) are unaffected. Subsequent use
        of this run object is undefined; most operations will raise an
        exception. See issue #144.
        """
        p = self._r.getparent()
        if p is None:
            return
        p.remove(self._r)

    @property
    def font(self) -> Font:
        """|Font| instance containing run-level character properties for the text in this run.

        Character properties can be and perhaps most often are inherited from parent objects such
        as the paragraph and slide layout the run is contained in. Only those specifically
        overridden at the run level are contained in the font object.
        """
        rPr = self._r.get_or_add_rPr()
        return Font(rPr, parent=self)

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


class _Field(Subshape):
    """Auto-refresh text field. Corresponds to an `a:fld` child element in a paragraph.

    The most common fields are a slide-number field (`type="slidenum"`) and a date field
    (`type="datetime"` / `"datetimeFigureOut"` / `"datetime1"` … `"datetime13"`). A `_Field`
    object is returned by :meth:`._Paragraph.add_field` and can also be iterated over via the
    runs-and-fields content of a paragraph in future APIs.
    """

    def __init__(self, fld: CT_TextField, parent: ProvidesPart):
        super(_Field, self).__init__(parent)
        self._fld = fld

    @property
    def field_type(self) -> str | None:
        """The `a:fld/@type` string, e.g. `"slidenum"`, `"datetime"`, or `"footer"`."""
        return self._fld.type

    @field_type.setter
    def field_type(self, value: str):
        self._fld.type = value

    @property
    def font(self) -> Font:
        """|Font| instance for the `a:rPr` run-properties child of this field."""
        rPr = self._fld.get_or_add_rPr()
        return Font(rPr)

    @property
    def text(self) -> str:
        """The placeholder-display text for this field (contents of `a:fld/a:t`).

        PowerPoint overwrites this text when it first renders the field — for a slide-number
        field, the number; for a date field, the rendered date. The value written here is what
        the viewer sees if it does not refresh the field itself.
        """
        return self._fld.text

    @text.setter
    def text(self, value: str):
        self._fld.text = value


def _replace_in_runs(
    runs: list[CT_RegularTextRun], find: str, replace: str
) -> int:
    """Replace every occurrence of `find` with `replace` within `runs`.

    `runs` is a non-empty list of consecutive `a:r` elements that share a
    parent `a:p`. Text is matched against the concatenation of the run
    texts, so matches may span adjacent runs. Returns the number of
    replacements performed.

    The run in which a match begins is retained (its formatting is
    preserved) and its `a:t` is rewritten to hold the text preceding the
    match plus the replacement. When a match spans into later runs, those
    runs are fully removed if entirely inside the match; a run that the
    match ends partway through keeps only its suffix text (and its own
    formatting). When repeat matches occur within a single run or a single
    span, they are all processed in left-to-right order by rebuilding that
    run's text in one pass.
    """
    # -- snapshot the text of each run and build the flattened string --
    run_texts = [r.text for r in runs]
    flat = "".join(run_texts)
    if find not in flat:
        return 0

    # -- map each character of `flat` to its source run index --
    run_index_at: list[int] = []
    for i, t in enumerate(run_texts):
        run_index_at.extend([i] * len(t))

    # -- collect non-overlapping match spans, left-to-right --
    matches: list[tuple[int, int]] = []  # (start, end) into `flat`
    search_start = 0
    while True:
        idx = flat.find(find, search_start)
        if idx < 0:
            break
        matches.append((idx, idx + len(find)))
        search_start = idx + len(find)
    if not matches:
        return 0

    # -- compute new text for each run and which runs to drop --
    # -- run_offsets[i] is the start offset of run i in `flat` --
    run_offsets: list[int] = []
    off = 0
    for t in run_texts:
        run_offsets.append(off)
        off += len(t)

    new_texts: list[str | None] = list(run_texts)  # None => drop run

    # -- walk matches in reverse so earlier indices stay valid --
    for start, end in reversed(matches):
        first_run = run_index_at[start] if start < len(run_index_at) else len(runs) - 1
        # -- `end` may equal len(flat); clamp to last run in that case --
        last_run = (
            run_index_at[end - 1] if end - 1 < len(run_index_at) else len(runs) - 1
        )
        first_text = new_texts[first_run]
        # -- first_text can't be None here: it contains the match start --
        assert first_text is not None
        prefix = first_text[: start - run_offsets[first_run]]
        # -- compute the surviving suffix of the last run (if partial) --
        last_text = new_texts[last_run]
        if last_text is None:
            # -- shouldn't happen for a match end, but guard anyway --
            suffix = ""
        else:
            suffix = last_text[end - run_offsets[last_run] :]

        if first_run == last_run:
            # -- match lies entirely within a single run; just splice --
            new_texts[first_run] = prefix + replace + suffix
        else:
            # -- match spans multiple runs: first run absorbs prefix +
            # -- replace; intermediate runs are dropped; last run keeps
            # -- only its suffix (preserving its own formatting).
            new_texts[first_run] = prefix + replace
            for mid in range(first_run + 1, last_run):
                new_texts[mid] = None
            new_texts[last_run] = suffix

    # -- apply changes back to the XML elements --
    # -- a run whose text becomes empty as a result of replacement is
    # -- removed (preserving an originally-empty run is unnecessary
    # -- after a targeted text rewrite).
    for run, new_text in zip(runs, new_texts):
        if new_text is None or (new_text == "" and run.text != ""):
            parent = run.getparent()
            if parent is not None:
                parent.remove(run)
        elif new_text != run.text:
            run.text = new_text

    return len(matches)
