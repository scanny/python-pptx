# encoding: utf-8

"""Text-related objects such as TextFrame and Paragraph."""

from pptx.compat import to_unicode
from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_AUTO_SIZE, MSO_UNDERLINE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.simpletypes import ST_TextWrappingType
from pptx.shapes import Subshape
from pptx.text.fonts import FontFiles
from pptx.text.layout import TextFitter
from pptx.util import Centipoints, Emu, lazyproperty, Pt


class TextFrame(Subshape):
    """The part of a shape that contains its text.

    Not all shapes have a text frame. Corresponds to the ``<p:txBody>`` element that can
    appear as a child element of ``<p:sp>``. Not intended to be constructed directly.
    """

    def __init__(self, txBody, parent):
        super(TextFrame, self).__init__(parent)
        self._element = self._txBody = txBody

    def add_paragraph(self):
        """
        Return new |_Paragraph| instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        p = self._txBody.add_p()
        return _Paragraph(p, self)

    @property
    def auto_size(self):
        """
        The type of automatic resizing that should be used to fit the text of
        this shape within its bounding box when the text would otherwise
        extend beyond the shape boundaries. May be |None|,
        ``MSO_AUTO_SIZE.NONE``, ``MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT``, or
        ``MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE``.
        """
        return self._bodyPr.autofit

    @auto_size.setter
    def auto_size(self, value):
        self._bodyPr.autofit = value

    def clear(self):
        """Remove all paragraphs except one empty one."""
        for p in self._txBody.p_lst[1:]:
            self._txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()

    def fit_text(
        self,
        font_family="Calibri",
        max_size=18,
        bold=False,
        italic=False,
        font_file=None,
    ):
        """Fit text-frame text entirely within bounds of its shape.

        Make the text in this text frame fit entirely within the bounds of
        its shape by setting word wrap on and applying the "best-fit" font
        size to all the text it contains. :attr:`TextFrame.auto_size` is set
        to :attr:`MSO_AUTO_SIZE.NONE`. The font size will not be set larger
        than *max_size* points. If the path to a matching TrueType font is
        provided as *font_file*, that font file will be used for the font
        metrics. If *font_file* is |None|, best efforts are made to locate
        a font file with matchhing *font_family*, *bold*, and *italic*
        installed on the current system (usually succeeds if the font is
        installed).
        """
        # ---no-op when empty as fit behavior not defined for that case---
        if self.text == "":
            return  # pragma: no cover

        font_size = self._best_fit_font_size(
            font_family, max_size, bold, italic, font_file
        )
        self._apply_fit(font_family, font_size, bold, italic)

    @property
    def margin_bottom(self):
        """
        |Length| value representing the inset of text from the bottom text
        frame border. :meth:`pptx.util.Inches` provides a convenient way of
        setting the value, e.g. ``text_frame.margin_bottom = Inches(0.05)``.
        """
        return self._bodyPr.bIns

    @margin_bottom.setter
    def margin_bottom(self, emu):
        self._bodyPr.bIns = emu

    @property
    def margin_left(self):
        """
        Inset of text from left text frame border as |Length| value.
        """
        return self._bodyPr.lIns

    @margin_left.setter
    def margin_left(self, emu):
        self._bodyPr.lIns = emu

    @property
    def margin_right(self):
        """
        Inset of text from right text frame border as |Length| value.
        """
        return self._bodyPr.rIns

    @margin_right.setter
    def margin_right(self, emu):
        self._bodyPr.rIns = emu

    @property
    def margin_top(self):
        """
        Inset of text from top text frame border as |Length| value.
        """
        return self._bodyPr.tIns

    @margin_top.setter
    def margin_top(self, emu):
        self._bodyPr.tIns = emu

    @property
    def paragraphs(self):
        """
        Immutable sequence of |_Paragraph| instances corresponding to the
        paragraphs in this text frame. A text frame always contains at least
        one paragraph.
        """
        return tuple([_Paragraph(p, self) for p in self._txBody.p_lst])

    @property
    def text(self):
        """Unicode/str containing all text in this text-frame.

        Read/write. The return value is a str (unicode) containing all text in this
        text-frame. A line-feed character (``"\\n"``) separates the text for each
        paragraph. A vertical-tab character (``"\\v"``) appears for each line break
        (aka. soft carriage-return) encountered.

        The vertical-tab character is how PowerPoint represents a soft carriage return
        in clipboard text, which is why that encoding was chosen.

        Assignment replaces all text in the text frame. The assigned value can be
        a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode. A bytes value
        (such as a Python 2 ``str``) is converted to unicode assuming UTF-8 encoding.
        A new paragraph is added for each line-feed character (``"\\n"``) encountered.
        A line-break (soft carriage-return) is inserted for each vertical-tab character
        (``"\\v"``) encountered.

        Any control character other than newline, tab, or vertical-tab are escaped as
        plain-text like "_x001B_" (for ESC (ASCII 32) in this example).
        """
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, text):
        txBody = self._txBody
        txBody.clear_content()
        for p_text in to_unicode(text).split("\n"):
            p = txBody.add_p()
            p.append_text(p_text)

    @property
    def vertical_anchor(self):
        """
        Read/write member of :ref:`MsoVerticalAnchor` enumeration or |None|,
        representing the vertical alignment of text in this text frame.
        |None| indicates the effective value should be inherited from this
        object's style hierarchy.
        """
        return self._txBody.bodyPr.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, value):
        bodyPr = self._txBody.bodyPr
        bodyPr.anchor = value

    @property
    def word_wrap(self):
        """
        Read-write setting determining whether lines of text in this shape
        are wrapped to fit within the shape's width. Valid values are True,
        False, or None. True and False turn word wrap on and off,
        respectively. Assigning None to word wrap causes any word wrap
        setting to be removed from the text frame, causing it to inherit this
        setting from its style hierarchy.
        """
        return {
            ST_TextWrappingType.SQUARE: True,
            ST_TextWrappingType.NONE: False,
            None: None,
        }[self._txBody.bodyPr.wrap]

    @word_wrap.setter
    def word_wrap(self, value):
        if value not in (True, False, None):
            raise ValueError(  # pragma: no cover
                "assigned value must be True, False, or None, got %s" % value
            )
        self._txBody.bodyPr.wrap = {
            True: ST_TextWrappingType.SQUARE,
            False: ST_TextWrappingType.NONE,
            None: None,
        }[value]

    def _apply_fit(self, font_family, font_size, is_bold, is_italic):
        """Arrange text in this text frame to fit inside its extents.

        This is accomplished by setting auto size off, wrap on, and setting the font of
        all its text to `font_family`, `font_size`, `is_bold`, and `is_italic`.
        """
        self.auto_size = MSO_AUTO_SIZE.NONE
        self.word_wrap = True
        self._set_font(font_family, font_size, is_bold, is_italic)

    def _best_fit_font_size(self, family, max_size, bold, italic, font_file):
        """
        Return the largest integer point size not greater than *max_size*
        that allows all the text in this text frame to fit inside its extents
        when rendered using the font described by *family*, *bold*, and
        *italic*. If *font_file* is specified, it is used to calculate the
        fit, whether or not it matches *family*, *bold*, and *italic*.
        """
        if font_file is None:
            font_file = FontFiles.find(family, bold, italic)
        return TextFitter.best_fit_font_size(
            self.text, self._extents, max_size, font_file
        )

    @property
    def _bodyPr(self):
        return self._txBody.bodyPr

    @property
    def _extents(self):
        """
        A (cx, cy) 2-tuple representing the effective rendering area for text
        within this text frame when margins are taken into account.
        """
        return (
            self._parent.width - self.margin_left - self.margin_right,
            self._parent.height - self.margin_top - self.margin_bottom,
        )

    def _set_font(self, family, size, bold, italic):
        """
        Set the font properties of all the text in this text frame to
        *family*, *size*, *bold*, and *italic*.
        """

        def iter_rPrs(txBody):
            for p in txBody.p_lst:
                for elm in p.content_children:
                    yield elm.get_or_add_rPr()
                # generate a:endParaRPr for each <a:p> element
                yield p.get_or_add_endParaRPr()

        def set_rPr_font(rPr, name, size, bold, italic):
            f = Font(rPr)
            f.name, f.size, f.bold, f.italic = family, Pt(size), bold, italic

        txBody = self._element
        for rPr in iter_rPrs(txBody):
            set_rPr_font(rPr, family, size, bold, italic)


class Font(object):
    """
    Character properties object, providing font size, font name, bold,
    italic, etc. Corresponds to ``<a:rPr>`` child element of a run. Also
    appears as ``<a:defRPr>`` and ``<a:endParaRPr>`` in paragraph and
    ``<a:defRPr>`` in list style elements.
    """

    def __init__(self, rPr):
        super(Font, self).__init__()
        self._element = self._rPr = rPr

    @property
    def bold(self):
        """
        Get or set boolean bold value of |Font|, e.g.
        ``paragraph.font.bold = True``. If set to |None|, the bold setting is
        cleared and is inherited from an enclosing shape's setting, or a
        setting in a style or master. Returns None if no bold attribute is
        present, meaning the effective bold value is inherited from a master
        or the theme.
        """
        return self._rPr.b

    @bold.setter
    def bold(self, value):
        self._rPr.b = value

    @lazyproperty
    def color(self):
        """
        The |ColorFormat| instance that provides access to the color settings
        for this font.
        """
        if self.fill.type != MSO_FILL.SOLID:
            self.fill.solid()
        return self.fill.fore_color

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this font, providing access to fill
        properties such as fill color.
        """
        return FillFormat.from_fill_parent(self._rPr)

    @property
    def italic(self):
        """
        Get or set boolean italic value of |Font| instance, with the same
        behaviors as bold with respect to None values.
        """
        return self._rPr.i

    @italic.setter
    def italic(self, value):
        self._rPr.i = value

    @property
    def language_id(self):
        """
        Get or set the language id of this |Font| instance. The language id
        is a member of the :ref:`MsoLanguageId` enumeration. Assigning |None|
        removes any language setting, the same behavior as assigning
        `MSO_LANGUAGE_ID.NONE`.
        """
        lang = self._rPr.lang
        if lang is None:
            return MSO_LANGUAGE_ID.NONE
        return self._rPr.lang

    @language_id.setter
    def language_id(self, value):
        if value == MSO_LANGUAGE_ID.NONE:
            value = None
        self._rPr.lang = value

    @property
    def name(self):
        """
        Get or set the typeface name for this |Font| instance, causing the
        text it controls to appear in the named font, if a matching font is
        found. Returns |None| if the typeface is currently inherited from the
        theme. Setting it to |None| removes any override of the theme
        typeface.
        """
        latin = self._rPr.latin
        if latin is None:
            return None
        return latin.typeface

    @name.setter
    def name(self, value):
        if value is None:
            self._rPr._remove_latin()
        else:
            latin = self._rPr.get_or_add_latin()
            latin.typeface = value

    @property
    def size(self):
        """
        Read/write |Length| value or |None|, indicating the font height in
        English Metric Units (EMU). |None| indicates the font size should be
        inherited from its style hierarchy, such as a placeholder or document
        defaults (usually 18pt). |Length| is a subclass of |int| having
        properties for convenient conversion into points or other length
        units. Likewise, the :class:`pptx.util.Pt` class allows convenient
        specification of point values::

            >> font.size = Pt(24)
            >> font.size
            304800
            >> font.size.pt
            24.0
        """
        sz = self._rPr.sz
        if sz is None:
            return None
        return Centipoints(sz)

    @size.setter
    def size(self, emu):
        if emu is None:
            self._rPr.sz = None
        else:
            sz = Emu(emu).centipoints
            self._rPr.sz = sz

    @property
    def underline(self):
        """
        Read/write. |True|, |False|, |None|, or a member of the
        :ref:`MsoTextUnderlineType` enumeration indicating the underline
        setting for this font. |None| is the default and indicates the
        underline setting should be inherited from the style hierarchy, such
        as from a placeholder. |True| indicates single underline. |False|
        indicates no underline. Other settings such as double and wavy
        underlining are indicated with members of the
        :ref:`MsoTextUnderlineType` enumeration.
        """
        u = self._rPr.u
        if u is MSO_UNDERLINE.NONE:
            return False
        if u is MSO_UNDERLINE.SINGLE_LINE:
            return True
        return u

    @underline.setter
    def underline(self, value):
        if value is True:
            value = MSO_UNDERLINE.SINGLE_LINE
        elif value is False:
            value = MSO_UNDERLINE.NONE
        self._element.u = value


class _Hyperlink(Subshape):
    """
    Text run hyperlink object. Corresponds to ``<a:hlinkClick>`` child
    element of the run's properties element (``<a:rPr>``).
    """

    def __init__(self, rPr, parent):
        super(_Hyperlink, self).__init__(parent)
        self._rPr = rPr

    @property
    def address(self):
        """
        Read/write. The URL of the hyperlink. URL can be on http, https,
        mailto, or file scheme; others may work.
        """
        if self._hlinkClick is None:
            return None
        return self.part.target_ref(self._hlinkClick.rId)

    @address.setter
    def address(self, url):
        # implements all three of add, change, and remove hyperlink
        if self._hlinkClick is not None:
            self._remove_hlinkClick()
        if url:
            self._add_hlinkClick(url)

    def _add_hlinkClick(self, url):
        rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
        self._rPr.add_hlinkClick(rId)

    @property
    def _hlinkClick(self):
        return self._rPr.hlinkClick

    def _remove_hlinkClick(self):
        assert self._hlinkClick is not None
        self.part.drop_rel(self._hlinkClick.rId)
        self._rPr._remove_hlinkClick()


class _Paragraph(Subshape):
    """Paragraph object. Not intended to be constructed directly."""

    def __init__(self, p, parent):
        super(_Paragraph, self).__init__(parent)
        self._element = self._p = p

    def add_line_break(self):
        """Add line break at end of this paragraph."""
        self._p.add_br()

    def add_run(self):
        """
        Return a new run appended to the runs in this paragraph.
        """
        r = self._p.add_r()
        return _Run(r, self)

    @property
    def alignment(self):
        """
        Horizontal alignment of this paragraph, represented by either
        a member of the enumeration :ref:`PpParagraphAlignment` or |None|.
        The value |None| indicates the paragraph should 'inherit' its
        effective value from its style hierarchy. Assigning |None| removes
        any explicit setting, causing its inherited value to be used.
        """
        return self._pPr.algn

    @alignment.setter
    def alignment(self, value):
        self._pPr.algn = value

    def clear(self):
        """
        Remove all content from this paragraph. Paragraph properties are
        preserved. Content includes runs, line breaks, and fields.
        """
        for elm in self._element.content_children:
            self._element.remove(elm)
        return self

    @property
    def font(self):
        """
        |Font| object containing default character properties for the runs in
        this paragraph. These character properties override default properties
        inherited from parent objects such as the text frame the paragraph is
        contained in and they may be overridden by character properties set at
        the run level.
        """
        return Font(self._defRPr)

    @property
    def level(self):
        """
        Read-write integer indentation level of this paragraph, having a
        range of 0-8 inclusive. 0 represents a top-level paragraph and is the
        default value. Indentation level is most commonly encountered in a
        bulleted list, as is found on a word bullet slide.
        """
        return self._pPr.lvl

    @level.setter
    def level(self, level):
        self._pPr.lvl = level

    @property
    def line_spacing(self):
        """
        Numeric or |Length| value specifying the space between baselines in
        successive lines of this paragraph. A value of |None| indicates no
        explicit value is assigned and its effective value is inherited from
        the paragraph's style hierarchy. A numeric value, e.g. `2` or `1.5`,
        indicates spacing is applied in multiples of line heights. A |Length|
        value such as ``Pt(12)`` indicates spacing is a fixed height. The
        |Pt| value class is a convenient way to apply line spacing in units
        of points.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.line_spacing

    @line_spacing.setter
    def line_spacing(self, value):
        pPr = self._p.get_or_add_pPr()
        pPr.line_spacing = value

    @property
    def runs(self):
        """
        Immutable sequence of |_Run| objects corresponding to the runs in
        this paragraph.
        """
        return tuple(_Run(r, self) for r in self._element.r_lst)

    @property
    def space_after(self):
        """
        |Length| value specifying the spacing to appear between this
        paragraph and the subsequent paragraph. A value of |None| indicates
        no explicit value is assigned and its effective value is inherited
        from the paragraph's style hierarchy. |Length| objects provide
        convenience properties, such as ``.pt`` and ``.inches``, that allow
        easy conversion to various length units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_after

    @space_after.setter
    def space_after(self, value):
        pPr = self._p.get_or_add_pPr()
        pPr.space_after = value

    @property
    def space_before(self):
        """
        |Length| value specifying the spacing to appear between this
        paragraph and the prior paragraph. A value of |None| indicates no
        explicit value is assigned and its effective value is inherited from
        the paragraph's style hierarchy. |Length| objects provide convenience
        properties, such as ``.pt`` and ``.cm``, that allow easy conversion
        to various length units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_before

    @space_before.setter
    def space_before(self, value):
        pPr = self._p.get_or_add_pPr()
        pPr.space_before = value

    @property
    def text(self):
        """str (unicode) representation of paragraph contents.

        Read/write. This value is formed by concatenating the text in each run and field
        making up the paragraph, adding a vertical-tab character (``"\\v"``) for each
        line-break element (`<a:br>`, soft carriage-return) encountered.

        While the encoding of line-breaks as a vertical tab might be surprising at
        first, doing so is consistent with PowerPoint's clipboard copy behavior and
        allows a line-break to be distinguished from a paragraph boundary within the str
        return value.

        Assignment causes all content in the paragraph to be replaced. Each vertical-tab
        character (``"\\v"``) in the assigned str is translated to a line-break, as is
        each line-feed character (``"\\n"``). Contrast behavior of line-feed character
        in `TextFrame.text` setter. If line-feed characters are intended to produce new
        paragraphs, use `TextFrame.text` instead. Any other control characters in the
        assigned string are escaped as a hex representation like "_x001B_" (for ESC
        (ASCII 27) in this example).

        The assigned value can be a 7-bit ASCII byte string (Python 2 str), a UTF-8
        encoded 8-bit byte string (Python 2 str), or unicode. Bytes values are converted
        to unicode assuming UTF-8 encoding.
        """
        return "".join(elm.text for elm in self._element.content_children)

    @text.setter
    def text(self, text):
        self.clear()
        self._element.append_text(to_unicode(text))

    @property
    def _defRPr(self):
        """
        The |CT_TextCharacterProperties| instance (<a:defRPr> element) that
        defines the default run properties for runs in this paragraph. Causes
        the element to be added if not present.
        """
        return self._pPr.get_or_add_defRPr()

    @property
    def _pPr(self):
        """
        The |CT_TextParagraphProperties| instance for this paragraph, the
        <a:pPr> element containing its paragraph properties. Causes the
        element to be added if not present.
        """
        return self._p.get_or_add_pPr()


class _Run(Subshape):
    """Text run object. Corresponds to ``<a:r>`` child element in a paragraph."""

    def __init__(self, r, parent):
        super(_Run, self).__init__(parent)
        self._r = r

    @property
    def font(self):
        """
        |Font| instance containing run-level character properties for the
        text in this run. Character properties can be and perhaps most often
        are inherited from parent objects such as the paragraph and slide
        layout the run is contained in. Only those specifically overridden at
        the run level are contained in the font object.
        """
        rPr = self._r.get_or_add_rPr()
        return Font(rPr)

    @lazyproperty
    def hyperlink(self):
        """
        |_Hyperlink| instance acting as proxy for any ``<a:hlinkClick>``
        element under the run properties element. Created on demand, the
        hyperlink object is available whether an ``<a:hlinkClick>`` element
        is present or not, and creates or deletes that element as appropriate
        in response to actions on its methods and attributes.
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
    def text(self, str):
        self._r.text = to_unicode(str)
