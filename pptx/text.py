# encoding: utf-8

"""
Text-related objects such as TextFrame and Paragraph.
"""

from .dml.fill import FillFormat
from .enum.dml import MSO_FILL
from .enum.text import MSO_ANCHOR
from .opc.constants import RELATIONSHIP_TYPE as RT
from .oxml.shared import Element, get_or_add
from .oxml.ns import _nsmap
from .shapes import Subshape
from .util import Centipoints, Emu, lazyproperty, to_unicode


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
        is no explicit margin setting, meaning the setting is inherited from
        a master or theme. Conversely, setting a margin to |None| removes any
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
        return tuple([_Paragraph(p, self) for p in self._txBody.p_lst])

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
        anchor = MSO_ANCHOR.to_xml(value)
        bodyPr = get_or_add(self._txBody, 'a:bodyPr')
        bodyPr.set('anchor', anchor)

    #: Write-only. Assignment to *vertical_anchor* sets the vertical
    #: alignment of the text frame to top, middle, or bottom. Valid values are
    #: ``MSO_ANCHOR.TOP``, ``MSO_ANCHOR.MIDDLE``, or ``MSO_ANCHOR.BOTTOM``.
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
        Get or set boolean italic value of |_Font| instance, with the same
        behaviors as bold with respect to None values.
        """
        return self._rPr.i

    @italic.setter
    def italic(self, value):
        self._rPr.i = value

    @property
    def name(self):
        """
        Get or set the typeface name for this |_Font| instance, causing the
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
        Height of the font in English Metric Units (EMU). The value is
        an instance of |BaseLength|, a subclass of |int| having properties
        for convenient conversion into points or other length units.
        Likewise, the :class:`pptx.util.Pt` class allows convenient
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
        sz = Emu(emu).centipoints
        self._rPr.sz = sz


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
        value like ``PP_ALIGN.CENTER``. Its value can be |None|, meaning the
        paragraph has no alignment setting and its effective value is
        inherited from a higher-level object.
        """
        return self._pPr.algn

    @alignment.setter
    def alignment(self, alignment):
        self._pPr.algn = alignment

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
        """
        Set the text of this run to *str*.
        """
        self._r.t.text = to_unicode(str)
