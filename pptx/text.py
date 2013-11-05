# encoding: utf-8

"""
Text-related objects such as TextFrame and Paragraph.
"""

from pptx.constants import MSO
from pptx.oxml.core import child, Element, get_or_add, SubElement
from pptx.oxml.ns import namespaces, qn
from pptx.spec import ParagraphAlignment
from pptx.util import to_unicode


# default namespace map for use in lxml calls
_nsmap = namespaces('a', 'r', 'p')


class TextFrame(object):
    """
    The part of a shape that contains its text. Not all shapes have a text
    frame. Corresponds to the ``<p:txBody>`` element that can appear as a
    child element of ``<p:sp>``. Not intended to be constructed directly.
    """
    def __init__(self, txBody):
        super(TextFrame, self).__init__()
        self._txBody = txBody

    @property
    def paragraphs(self):
        """
        Immutable sequence of |_Paragraph| instances corresponding to the
        paragraphs in this text frame. A text frame always contains at least
        one paragraph.
        """
        return tuple([_Paragraph(p) for p in self._txBody[qn('a:p')]])

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
        value_map = {MSO.ANCHOR_TOP: 't', MSO.ANCHOR_MIDDLE: 'ctr',
                     MSO.ANCHOR_BOTTOM: 'b'}
        bodyPr = get_or_add(self._txBody, 'a:bodyPr')
        bodyPr.set('anchor', value_map[value])

    #: Write-only. Assignment to *vertical_anchor* sets the vertical
    #: alignment of the text frame to top, middle, or bottom. Valid values are
    #: ``MSO.ANCHOR_TOP``, ``MSO.ANCHOR_MIDDLE``, or ``MSO.ANCHOR_BOTTOM``.
    #: The ``MSO`` name is imported from ``pptx.constants``.
    vertical_anchor = property(None, _set_vertical_anchor)

    def _set_word_wrap(self, value):
        """
        Set ``wrap`` attribution of ``<a:bodyPr>`` element. Can be
        one of True, False, or None.
        """
        bodyPr = get_or_add(self._txBody, 'a:bodyPr')

        if value is None:
            del bodyPr.attrib['wrap']
            return

        value_map = {True: 'square', False: 'none'}
        bodyPr.set('wrap', value_map[value])

    def _get_word_wrap(self):
        """
        Return the value of the word_wrap setting. Possible return values
        are True, False, and None.
        """
        value_map = {'square': True, 'none': False, None: None}
        bodyPr = get_or_add(self._txBody, 'a:bodyPr')
        value = bodyPr.get('wrap')
        return value_map[value]

    #: Read-write. Assignment to *word_wrap* sets the wrapping behavior
    #: of the text frame. The valid values are True, False, and None. True
    #: and False turn word wrap on and off, and None will set it to inherit
    #: the wrapping behavior from its parent element.
    word_wrap = property(_get_word_wrap, _set_word_wrap)

    def add_paragraph(self):
        """
        Return new |_Paragraph| instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        # <a:p> elements are last in txBody, so can simply append new one
        p = Element('a:p')
        self._txBody.append(p)
        return _Paragraph(p)

    def clear(self):
        """
        Remove all paragraphs except one empty one.
        """
        p_list = self._txBody.xpath('./a:p', namespaces=_nsmap)
        for p in p_list[1:]:
            self._txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()


class _Paragraph(object):
    """
    Paragraph object. Not intended to be constructed directly.
    """
    def __init__(self, p):
        super(_Paragraph, self).__init__()
        self._p = p

    def add_run(self):
        """
        Return a new run appended to the runs in this paragraph.
        """
        r = self._p.add_r()
        return _Run(r)

    @property
    def alignment(self):
        """
        Horizontal alignment of this paragraph, represented by a constant
        value like ``PP.ALIGN_CENTER``. Its value can be |None|, meaning the
        paragraph has no alignment setting and its effective value is
        inherited from a higher-level object.
        """
        algn = self._p.get_algn()
        return ParagraphAlignment.from_text_align_type(algn)

    @alignment.setter
    def alignment(self, alignment):
        algn = ParagraphAlignment.to_text_align_type(alignment)
        self._p.set_algn(algn)

    def clear(self):
        """Remove all runs from this paragraph."""
        # retain pPr if present
        pPr = child(self._p, 'a:pPr')
        self._p.clear()
        if pPr is not None:
            self._p.insert(0, pPr)

    @property
    def font(self):
        """
        |_Font| object containing default character properties for the runs in
        this paragraph. These character properties override default properties
        inherited from parent objects such as the text frame the paragraph is
        contained in and they may be overridden by character properties set at
        the run level.
        """
        # A _Font instance is created on first access if it doesn't exist.
        # This can cause "litter" <a:pPr> and <a:defRPr> elements to be
        # included in the XML if the _Font element is referred to but not
        # populated with values.
        if not hasattr(self._p, 'pPr'):
            pPr = Element('a:pPr')
            self._p.insert(0, pPr)
        if not hasattr(self._p.pPr, 'defRPr'):
            SubElement(self._p.pPr, 'a:defRPr')
        return _Font(self._p.pPr.defRPr)

    @property
    def level(self):
        """
        Read-write integer indentation level of this paragraph, having a
        range of 0-8 inclusive. 0 represents a top-level paragraph and is the
        default value. Indentation level is most commonly encountered in a
        bulleted list, as is found on a word bullet slide.
        """
        if not hasattr(self._p, 'pPr'):
            return 0
        return int(self._p.pPr.get('lvl', 0))

    @level.setter
    def level(self, level):
        if not isinstance(level, int) or level < 0 or level > 8:
            msg = "paragraph level must be integer between 0 and 8 inclusive"
            raise ValueError(msg)
        if not hasattr(self._p, 'pPr'):
            pPr = Element('a:pPr')
            self._p.insert(0, pPr)
        self._p.pPr.set('lvl', str(level))

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
            runs.append(_Run(r))
        return tuple(runs)

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
        b = self._rPr.get('b')
        if b is None:
            return None
        return True if b in ('true', '1') else False

    @bold.setter
    def bold(self, bool):
        if bool is None:
            if 'b' in self._rPr.attrib:
                del self._rPr.attrib['b']
        else:
            self._rPr.set('b', '1' if bool else '0')

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


class _Run(object):
    """
    Text run object. Corresponds to ``<a:r>`` child element in a paragraph.
    """
    def __init__(self, r):
        super(_Run, self).__init__()
        self._r = r

    @property
    def font(self):
        """
        |_Font| object containing run-level character properties for the text
        in this run. Character properties can and perhaps most often are
        inherited from parent objects such as the paragraph and slide layout
        the run is contained in. Only those specifically assigned at the run
        level are contained in the |_Font| object.
        """
        if not hasattr(self._r, 'rPr'):
            self._r.insert(0, Element('a:rPr'))
        return _Font(self._r.rPr)

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
