# encoding: utf-8

"""
Text-related objects such as TextFrame and Paragraph.
"""

from pptx.constants import MSO
from pptx.dml.core import ColorFormat, RGBColor
from pptx.enum import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.core import Element, get_or_add
from pptx.oxml.ns import namespaces, qn
from pptx.shapes import Subshape
from pptx.spec import ParagraphAlignment
from pptx.util import lazyproperty, to_unicode


# default namespace map for use in lxml calls
_nsmap = namespaces('a', 'r', 'p')


class TextFrame(Subshape):
    """
    The part of a shape that contains its text. Not all shapes have a text
    frame. Corresponds to the ``<p:txBody>`` element that can appear as a
    child element of ``<p:sp>``. Not intended to be constructed directly.
    """
    def __init__(self, txBody, parent):
        super(TextFrame, self).__init__(parent)
        self._txBody = txBody

    def add_paragraph(self):
        """
        Return new |_Paragraph| instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        # <a:p> elements are last in txBody, so can simply append new one
        p = Element('a:p')
        self._txBody.append(p)
        return _Paragraph(p, self)

    def clear(self):
        """
        Remove all paragraphs except one empty one.
        """
        p_list = self._txBody.xpath('./a:p', namespaces=_nsmap)
        for p in p_list[1:]:
            self._txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()

    @property
    def margin_bottom(self):
        """
        Inset of text from textframe border in EMU. ``pptx.util.Inches``
        provides a convenient way of setting the value, e.g.
        ``textframe.margin_bottom = Inches(0.05)``. Returns |None| if there
        is explicit margin setting, meaning the setting is inherited from a
        master or theme. Conversely, setting a margin to |None| removes any
        explicit setting at the shape level and restores inheritance of the
        effective value.
        """
        return self._bodyPr.bIns

    @margin_bottom.setter
    def margin_bottom(self, emu):
        self._bodyPr.bIns = emu

    @property
    def margin_left(self):
        return self._bodyPr.lIns

    @margin_left.setter
    def margin_left(self, emu):
        self._bodyPr.lIns = emu

    @property
    def margin_right(self):
        return self._bodyPr.rIns

    @margin_right.setter
    def margin_right(self, emu):
        self._bodyPr.rIns = emu

    @property
    def margin_top(self):
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
        return tuple([_Paragraph(p, self) for p in self._txBody[qn('a:p')]])

    def _set_text(self, text):
        """Replace all text in text frame with single run containing *text*"""
        self.clear()
        self.paragraphs[0].text = to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the text frame with the assigned expression. After assignment, the
    #: text frame contains exactly one paragraph containing a single run
    #: containing all the text. The assigned value can be a 7-bit ASCII
    #: string, a UTF-8 encoded 8-bit string, or unicode. String values are
    #: converted to unicode assuming UTF-8 encoding.
    text = property(None, _set_text)

    def _set_vertical_anchor(self, value):
        """
        Set ``anchor`` attribute of ``<a:bodyPr>`` element
        """
        value_map = {
            MSO.ANCHOR_TOP:    't',
            MSO.ANCHOR_MIDDLE: 'ctr',
            MSO.ANCHOR_BOTTOM: 'b'
        }
        bodyPr = get_or_add(self._txBody, 'a:bodyPr')
        bodyPr.set('anchor', value_map[value])

    #: Write-only. Assignment to *vertical_anchor* sets the vertical
    #: alignment of the text frame to top, middle, or bottom. Valid values are
    #: ``MSO.ANCHOR_TOP``, ``MSO.ANCHOR_MIDDLE``, or ``MSO.ANCHOR_BOTTOM``.
    #: The ``MSO`` name is imported from ``pptx.constants``.
    vertical_anchor = property(None, _set_vertical_anchor)

    @property
    def word_wrap(self):
        """
        Read-write value of the word wrap setting for this text frame, either
        True, False, or None. Assignment to *word_wrap* sets the wrapping
        behavior. True and False turn word wrap on and off, respectively.
        Assigning None to word wrap causes its word wrap setting to be
        removed entirely and the text frame wrapping behavior to be inherited
        from a parent element.
        """
        value_map = {'square': True, 'none': False, None: None}
        bodyPr = get_or_add(self._txBody, 'a:bodyPr')
        value = bodyPr.get('wrap')
        return value_map[value]

    @word_wrap.setter
    def word_wrap(self, value):
        value_map = {True: 'square', False: 'none'}
        bodyPr = get_or_add(self._txBody, 'a:bodyPr')
        if value is None:
            if 'wrap' in bodyPr.attrib:
                del bodyPr.attrib['wrap']
            return
        bodyPr.set('wrap', value_map[value])

    @property
    def _bodyPr(self):
        return self._txBody.bodyPr


class _Font(object):
    """
    Character properties object, providing font size, font name, bold,
    italic, etc. Corresponds to ``<a:rPr>`` child element of a run. Also
    appears as ``<a:defRPr>`` and ``<a:endParaRPr>`` in paragraph and
    ``<a:defRPr>`` in list style elements.
    """
    def __init__(self, rPr):
        super(_Font, self).__init__()
        self._rPr = rPr

    @property
    def bold(self):
        """
        Get or set boolean bold value of |_Font|, e.g.
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

    @property
    def color(self):
        """
        The |ColorFormat| instance that provides access to the color settings
        for this font.
        """
        if not hasattr(self, '_color'):
            self._color = _FontColor(self._rPr)
        return self._color

    @property
    def italic(self):
        """
        Get or set boolean italic value of |_Font| instance, with the same
        behaviors as bold with respect to None values.
        """
        return self._rPr.i

    @italic.setter
    def italic(self, value):
        self._rPr.i = value

    def _set_size(self, centipoints):
        # handle float centipoints value gracefully
        centipoints = int(centipoints)
        self._rPr.set('sz', str(centipoints))

    #: Set the font size. In PresentationML, font size is expressed in
    #: hundredths of a point (centipoints). The :class:`pptx.util.Pt` class
    #: allows convenient conversion to centipoints from float or integer point
    #: values, e.g. ``Pt(12.5)``. I'm pretty sure I just made up the word
    #: *centipoint*, but it seems apt :).
    size = property(None, _set_size)


class _FontColor(ColorFormat):
    """
    Provides access to font color settings.
    """
    def __init__(self, rPr):
        # note that rPr is not always an actual <a:rPr> element, but must
        # always be an instance of CT_TextCharacterProperties, so will behave
        # as though it were one
        super(_FontColor, self).__init__()
        self._rPr = rPr

    @property
    def brightness(self):
        """
        Read/write float value between -1.0 and 1.0 indicating the brightness
        adjustment for this color, e.g. -0.25 is 25% darker and 0.4 is 40%
        lighter. 0 means no brightness adjustment.
        """
        if self._color_elm is None:
            return 0
        lumMod, lumOff = self._color_elm.lumMod, self._color_elm.lumOff
        # a tint is lighter, a shade is darker
        # only tints have lumOff child
        if lumOff is not None:
            val = lumOff.val
            brightness = float(val) / 100000
            return brightness
        # which leaves shades, if lumMod is present
        if lumMod is not None:
            val = lumMod.val
            brightness = -1.0 + float(val)/100000
            return brightness
        # there's no brightness adjustment if no lum{Mod|Off} elements
        return 0

    @brightness.setter
    def brightness(self, value):
        self._validate_brightness_value(value)
        if value > 0:
            self._tint(value)
        elif value < 0:
            self._shade(value)
        else:
            self._color_elm.clear_lum()

    @property
    def rgb(self):
        """
        |RGBColor| value of this color, or None if no RGB color is explicitly
        defined for this font. Setting this value to an |RGBColor| instance
        cause its type to change to MSO_COLOR_TYPE.RGB. If the color was a
        theme color with a brightness adjustment, the brightness adjustment
        is removed when changing it to an RGB color.
        """
        if self._srgbClr is None:
            return None
        return RGBColor.from_string(self._srgbClr.val)

    @rgb.setter
    def rgb(self, rgb):
        if not isinstance(rgb, RGBColor):
            raise TypeError('assigned value must be type RGBColor')
        solidFill = self._rPr.get_or_change_to_solidFill()
        srgbClr = solidFill.get_or_change_to_srgbClr()
        srgbClr.val = str(rgb)

    @property
    def theme_color(self):
        """
        Theme color value of this color, one of those defined in the
        MSO_THEME_COLOR enumeration, e.g. MSO_THEME_COLOR.ACCENT_1. None if
        no theme color is explicitly defined for this font. Setting this to a
        value in MSO_THEME_COLOR causes the color's type to change to
        ``MSO_COLOR_TYPE.SCHEME``.
        """
        if self._schemeClr is None:
            return None
        return MSO_THEME_COLOR.from_xml(self._schemeClr.val)

    @theme_color.setter
    def theme_color(self, mso_theme_color_idx):
        solidFill = self._rPr.get_or_change_to_solidFill()
        schemeClr = solidFill.get_or_change_to_schemeClr()
        schemeClr.val = MSO_THEME_COLOR.to_xml(mso_theme_color_idx)

    @property
    def type(self):
        """
        Read-only. A value from MSO_COLOR_TYPE, either RGB or SCHEME,
        corresponding to the way this color is defined, or None if no color
        is defined at the level of this font.
        """
        if self._srgbClr is not None:
            return MSO_COLOR_TYPE.RGB
        if self._schemeClr is not None:
            return MSO_COLOR_TYPE.SCHEME
        return None

    @property
    def _color_elm(self):
        """
        srgbClr or schemeClr child of <a:solidFill>, None if neither is
        present.
        """
        srgbClr = self._srgbClr
        if srgbClr is not None:
            return srgbClr
        schemeClr = self._schemeClr
        if schemeClr is not None:
            return schemeClr
        return None

    @property
    def _schemeClr(self):
        """
        schemeClr child of <a:solidFill> if present, None otherwise
        """
        solidFill = self._rPr.solidFill
        if solidFill is None:
            return None
        return solidFill.schemeClr

    def _shade(self, value):
        lumMod_val = 100000 - int(abs(value) * 100000)
        color_elm = self._color_elm.clear_lum()
        color_elm.add_lumMod(lumMod_val)

    def _tint(self, value):
        lumOff_val = int(value * 100000)
        lumMod_val = 100000 - lumOff_val
        color_elm = self._color_elm.clear_lum()
        color_elm.add_lumMod(lumMod_val)
        color_elm.add_lumOff(lumOff_val)

    @property
    def _srgbClr(self):
        """
        srgbClr child of <a:solidFill> if present, None otherwise
        """
        solidFill = self._rPr.solidFill
        if solidFill is None:
            return None
        return solidFill.srgbClr

    def _validate_brightness_value(self, value):
        if value < -1.0 or value > 1.0:
            raise ValueError('brightness must be number in range -1.0 to 1.0')
        if self._color_elm is None:
            msg = (
                "can't set brightness when color.type is None. Set color.rgb"
                " or .theme_color first."
            )
            raise ValueError(msg)


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
        if url is not None:
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
        self._rPr.hlinkClick = None


class _Paragraph(Subshape):
    """
    Paragraph object. Not intended to be constructed directly.
    """
    def __init__(self, p, parent):
        super(_Paragraph, self).__init__(parent)
        self._p = p

    def add_run(self):
        """
        Return a new run appended to the runs in this paragraph.
        """
        r = self._p.add_r()
        return _Run(r, self)

    @property
    def alignment(self):
        """
        Horizontal alignment of this paragraph, represented by a constant
        value like ``PP.ALIGN_CENTER``. Its value can be |None|, meaning the
        paragraph has no alignment setting and its effective value is
        inherited from a higher-level object.
        """
        return ParagraphAlignment.from_text_align_type(self._pPr.algn)

    @alignment.setter
    def alignment(self, alignment):
        algn = ParagraphAlignment.to_text_align_type(alignment)
        self._pPr.algn = algn

    def clear(self):
        """
        Remove all runs from this paragraph. Paragraph properties are
        preserved.
        """
        self._p.remove_child_r_elms()

    @property
    def font(self):
        """
        |_Font| object containing default character properties for the runs in
        this paragraph. These character properties override default properties
        inherited from parent objects such as the text frame the paragraph is
        contained in and they may be overridden by character properties set at
        the run level.
        """
        return _Font(self._defRPr)

    @property
    def level(self):
        """
        Read-write integer indentation level of this paragraph, having a
        range of 0-8 inclusive. 0 represents a top-level paragraph and is the
        default value. Indentation level is most commonly encountered in a
        bulleted list, as is found on a word bullet slide.
        """
        # return self._pPr.lvl
        return int(self._pPr.get('lvl', 0))

    @level.setter
    def level(self, level):
        if not isinstance(level, int) or level < 0 or level > 8:
            msg = "paragraph level must be integer between 0 and 8 inclusive"
            raise ValueError(msg)
        self._pPr.set('lvl', str(level))

    @property
    def runs(self):
        """
        Immutable sequence of |_Run| instances corresponding to the runs in
        this paragraph.
        """
        xpath = './a:r'
        r_elms = self._p.xpath(xpath, namespaces=_nsmap)
        runs = []
        for r in r_elms:
            runs.append(_Run(r, self))
        return tuple(runs)

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

    def _set_text(self, text):
        """Replace runs with single run containing *text*"""
        self.clear()
        r = self.add_run()
        r.text = to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the paragraph. After assignment, the paragraph containins exactly
    #: one run containing the text value of the assigned expression. The
    #: assigned value can be a 7-bit ASCII string, a UTF-8 encoded 8-bit
    #: string, or unicode. String values are converted to unicode assuming
    #: UTF-8 encoding.
    text = property(None, _set_text)


class _Run(Subshape):
    """
    Text run object. Corresponds to ``<a:r>`` child element in a paragraph.
    """
    def __init__(self, r, parent):
        super(_Run, self).__init__(parent)
        self._r = r

    @property
    def font(self):
        """
        |_Font| instance containing run-level character properties for the
        text in this run. Character properties can be and perhaps most often
        are inherited from parent objects such as the paragraph and slide
        layout the run is contained in. Only those specifically overridden at
        the run level are contained in the font object.
        """
        rPr = self._r.get_or_add_rPr()
        return _Font(rPr)

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
        """
        Read/Write. Text contained in the run. A regular text run is required
        to contain exactly one ``<a:t>`` (text) element. Assignment to *text*
        replaces the text currently contained in the run. The assigned value
        can be a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode.
        String values are converted to unicode assuming UTF-8 encoding.
        """
        return self._r.t.text

    @text.setter
    def text(self, str):
        """Set the text of this run to *str*."""
        self._r.t._setText(to_unicode(str))
