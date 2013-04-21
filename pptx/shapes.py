# -*- coding: utf-8 -*-
#
# shapes.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
Classes that implement PowerPoint shapes such as picture, textbox, and table.
"""

import os

try:
    from PIL import Image as PIL_Image
except ImportError:
    import Image as PIL_Image

from pptx.constants import MSO
from pptx.oxml import (
    _get_or_add, qn, _Element, _SubElement, CT_GraphicalObjectFrame)
from pptx.spec import namespaces, slide_ph_basenames
from pptx.spec import RT_IMAGE
from pptx.spec import (
    PH_TYPE_BODY, PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ,
    PH_TYPE_SLDNUM, PH_TYPE_SUBTITLE, PH_TYPE_TITLE, PH_ORIENT_HORZ,
    PH_ORIENT_VERT, PH_SZ_FULL)
from pptx.util import Collection, Px

# default namespace map for use in lxml calls
_nsmap = namespaces('a', 'r', 'p')


def _child(element, child_tagname, nsmap=None):
    """
    Return direct child of *element* having *child_tagname* or :class:`None`
    if no such child element is present.
    """
    # use default nsmap if not specified
    if nsmap is None:
        nsmap = _nsmap
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


def _to_unicode(text):
    """
    Return *text* as a unicode string.

    *text* can be a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or
    unicode. String values are converted to unicode assuming UTF-8 encoding.
    Unicode values are returned unchanged.
    """
    # both str and unicode inherit from basestring
    if not isinstance(text, basestring):
        tmpl = 'expected UTF-8 encoded string or unicode, got %s value %s'
        raise TypeError(tmpl % (type(text), text))
    # return unicode strings unchanged
    if isinstance(text, unicode):
        return text
    # otherwise assume UTF-8 encoding, which also works for ASCII
    return unicode(text, 'utf-8')


# ============================================================================
# Shapes
# ============================================================================

class BaseShape(object):
    """
    Base class for shape objects. Both :class:`Shape` and :class:`Picture`
    inherit from :class:`BaseShape`.
    """
    def __init__(self, shape_element):
        super(BaseShape, self).__init__()
        self._element = shape_element
        # e.g. nvSpPr for shape, nvPicPr for pic, etc.
        self.__nvXxPr = shape_element.xpath('./*[1]', namespaces=_nsmap)[0]

    @property
    def has_textframe(self):
        """
        True if this shape has a txBody element and can contain text.
        """
        return _child(self._element, 'p:txBody') is not None

    @property
    def id(self):
        """
        Id of this shape. Note that ids are constrained to positive integers.
        """
        return int(self.__nvXxPr.cNvPr.get('id'))

    @property
    def is_placeholder(self):
        """
        True if this shape is a placeholder. A shape is a placeholder if it
        has a <p:ph> element.
        """
        return _child(self.__nvXxPr.nvPr, 'p:ph') is not None

    @property
    def name(self):
        """Name of this shape."""
        return self.__nvXxPr.cNvPr.get('name')

    def _set_text(self, text):
        """Replace all text in shape with single run containing *text*"""
        if not self.has_textframe:
            raise TypeError("cannot set text of shape with no text frame")
        self.textframe.text = _to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: by the shape, resulting in a text frame containing exactly one
    #: paragraph, itself containing a single run. The assigned value can be a
    #: 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode. String
    #: values are converted to unicode assuming UTF-8 encoding.
    text = property(None, _set_text)

    @property
    def textframe(self):
        """
        TextFrame instance for this shape. Raises :class:`ValueError` if shape
        has no text frame. Use :attr:`has_textframe` to check whether a shape
        has a text frame.
        """
        txBody = _child(self._element, 'p:txBody')
        if txBody is None:
            raise ValueError('shape has no text frame')
        return TextFrame(txBody)

    @property
    def _is_title(self):
        """
        True if this shape is a title placeholder.
        """
        ph = _child(self.__nvXxPr.nvPr, 'p:ph')
        if ph is None:
            return False
        # idx defaults to 0 when idx attr is absent
        ph_idx = ph.get('idx', '0')
        # title placeholder is identified by idx of 0
        return ph_idx == '0'


class ShapeCollection(BaseShape, Collection):
    """
    Sequence of shapes. Corresponds to CT_GroupShape in pml schema. Note that
    while spTree in a slide is a group shape, the group shape is recursive in
    that a group shape can include other group shapes within it.
    """
    _NVGRPSPPR = qn('p:nvGrpSpPr')
    _GRPSPPR = qn('p:grpSpPr')
    _SP = qn('p:sp')
    _GRPSP = qn('p:grpSp')
    _GRAPHICFRAME = qn('p:graphicFrame')
    _CXNSP = qn('p:cxnSp')
    _PIC = qn('p:pic')
    _CONTENTPART = qn('p:contentPart')
    _EXTLST = qn('p:extLst')

    def __init__(self, spTree, slide=None):
        super(ShapeCollection, self).__init__(spTree)
        self.__spTree = spTree
        self.__slide = slide
        self.__shapes = self._values
        # unmarshal shapes
        for elm in spTree.iterchildren():
            # log.debug('elm.tag == %s', elm.tag[60:])
            if elm.tag in (self._NVGRPSPPR, self._GRPSPPR, self._EXTLST):
                continue
            elif elm.tag == self._SP:
                shape = Shape(elm)
            elif elm.tag == self._PIC:
                shape = Picture(elm)
            elif elm.tag == self._GRPSP:
                shape = ShapeCollection(elm)
            elif elm.tag == self._GRAPHICFRAME:
                shape = Table(elm)
            elif elm.tag == self._CONTENTPART:
                msg = "first time 'contentPart' shape encountered in the "\
                      "wild, please let developer know and send example"
                raise ValueError(msg)
            else:
                shape = BaseShape(elm)
            self.__shapes.append(shape)

    @property
    def placeholders(self):
        """
        Immutable sequence containing the placeholder shapes in this shape
        collection, sorted in *idx* order.
        """
        placeholders =\
            [Placeholder(sp) for sp in self.__shapes if sp.is_placeholder]
        placeholders.sort(key=lambda ph: ph.idx)
        return tuple(placeholders)

    @property
    def title(self):
        """The title shape in collection or None if no title placeholder."""
        for shape in self.__shapes:
            if shape._is_title:
                return shape
        return None

    def add_picture(self, file, left, top, width=None, height=None):
        """
        Add picture shape displaying image in *file*, where *file* can be
        either a path to a file (a string) or a file-like object.
        """
        image = self.__package._images.add_image(file)
        rel = self.__slide._add_relationship(RT_IMAGE, image)
        pic = self.__pic(rel._rId, file, left, top, width, height)
        self.__spTree.append(pic)
        picture = Picture(pic)
        self.__shapes.append(picture)
        return picture

    def add_table(self, rows, cols, left, top, width, height):
        """
        Add table shape with the specified number of *rows* and *cols* at the
        specified position with the specified size.
        """
        id = self.__next_shape_id
        name = 'Table %d' % (id-1)
        graphicFrame = CT_GraphicalObjectFrame(id, name, rows, cols, left,
                                               top, width, height)
        self.__spTree.append(graphicFrame)
        table = Table(graphicFrame)
        self.__shapes.append(table)
        return table

    def add_textbox(self, left, top, width, height):
        """
        Add text box shape of specified size at specified position.
        """
        id = self.__next_shape_id
        name = 'TextBox %d' % (id-1)
        sp = self.__sp(id, name, left, top, width, height, is_textbox=True)
        self.__spTree.append(sp)
        shape = Shape(sp)
        self.__shapes.append(shape)
        return shape

    def _clone_layout_placeholders(self, slidelayout):
        """
        Add placeholder shapes based on those in *slidelayout*. Z-order of
        placeholders is preserved. Latent placeholders (date, slide number,
        and footer) are not cloned.
        """
        latent_ph_types = (PH_TYPE_DT, PH_TYPE_SLDNUM, PH_TYPE_FTR)
        for sp in slidelayout.shapes:
            if not sp.is_placeholder:
                continue
            ph = Placeholder(sp)
            if ph.type in latent_ph_types:
                continue
            self.__clone_layout_placeholder(ph)

    def __clone_layout_placeholder(self, layout_ph):
        """
        Add a new placeholder shape based on the slide layout placeholder
        *layout_ph*.
        """
        id = self.__next_shape_id
        ph_type = layout_ph.type
        orient = layout_ph.orient
        shapename = self.__next_ph_name(ph_type, id, orient)

        sp = self.__new_placeholder_sp(layout_ph, id, ph_type, orient,
                                       shapename)
        self.__spTree.append(sp)
        shape = Shape(sp)
        self.__shapes.append(shape)
        return shape

    def __new_placeholder_sp(self, layout_ph, id, ph_type, orient, shapename):
        """
        Assemble a new ``<p:sp>`` element based on the specified parameters.
        """
        # form XML hierarchy
        sp = _Element('p:sp', _nsmap)
        _SubElement(sp, 'p:nvSpPr')
        _SubElement(sp.nvSpPr, 'p:cNvPr')
        sp.nvSpPr.cNvPr.set('id', str(id))
        sp.nvSpPr.cNvPr.set('name', shapename)
        _SubElement(sp.nvSpPr, 'p:cNvSpPr')
        _SubElement(sp.nvSpPr.cNvSpPr, 'a:spLocks')
        sp.nvSpPr.cNvSpPr[qn('a:spLocks')].set('noGrp', '1')

        _SubElement(sp.nvSpPr, 'p:nvPr')
        ph = _SubElement(sp.nvSpPr.nvPr, 'p:ph')
        if ph_type != PH_TYPE_OBJ:
            ph.set('type', ph_type)
        if layout_ph.orient != PH_ORIENT_HORZ:
            ph.set('orient', layout_ph.orient)
        if layout_ph.sz != PH_SZ_FULL:
            ph.set('sz', layout_ph.sz)
        if layout_ph.idx != 0:
            ph.set('idx', str(layout_ph.idx))

        _SubElement(sp, 'p:spPr')

        placeholder_types_that_have_a_text_frame = (
            PH_TYPE_TITLE, PH_TYPE_CTRTITLE, PH_TYPE_SUBTITLE, PH_TYPE_BODY,
            PH_TYPE_OBJ)

        if ph_type in placeholder_types_that_have_a_text_frame:
            _SubElement(sp, 'p:txBody')
            _SubElement(sp.txBody, 'a:bodyPr')
            _SubElement(sp.txBody, 'a:lstStyle')
            _SubElement(sp.txBody, 'a:p')

        return sp

    def __next_ph_name(self, ph_type, id, orient):
        """
        Next unique placeholder name for placeholder shape of type *ph_type*,
        with id number *id* and orientation *orient*. Usually will be standard
        placeholder root name suffixed with id-1, e.g.
        __next_ph_name(PH_TYPE_TBL, 4, 'horz') ==> 'Table Placeholder 3'. The
        number is incremented as necessary to make the name unique within the
        collection. If *orient* is ``'vert'``, the placeholder name is
        prefixed with ``'Vertical '``.
        """
        basename = slide_ph_basenames[ph_type]
        # prefix rootname with 'Vertical ' if orient is 'vert'
        if orient == PH_ORIENT_VERT:
            basename = 'Vertical %s' % basename
        # increment numpart as necessary to make name unique
        numpart = id - 1
        names = self.__spTree.xpath('//p:cNvPr/@name', namespaces=_nsmap)
        while True:
            name = '%s %d' % (basename, numpart)
            if name not in names:
                break
            numpart += 1
        # log.debug("assigned placeholder name '%s'" % name)
        return name

    @property
    def __next_shape_id(self):
        """
        Next available drawing object id number in collection, starting from 1
        and making use of any gaps in numbering. In practice, the minimum id
        is 2 because the spTree element is always assigned id="1".
        """
        cNvPrs = self.__spTree.xpath('//p:cNvPr', namespaces=_nsmap)
        ids = [int(cNvPr.get('id')) for cNvPr in cNvPrs]
        ids.sort()
        # first gap in sequence wins, or falls off the end as max(ids)+1
        next_id = 1
        for id in ids:
            if id > next_id:
                break
            next_id += 1
        return next_id

    @property
    def __package(self):
        """
        Reference to |Package| this shape collection resides in.
        """
        return self.__slide._package

    def __pic(self, rId, file, x, y, cx=None, cy=None):
        """
        Return minimal ``<p:pic>`` element based on *rId* and *file*. *file* is
        either a path to the file (a string) or a file-like object.
        """
        id = self.__next_shape_id
        shapename = 'Picture %d' % (id-1)
        if isinstance(file, basestring):  # *file* is a path
            filename = os.path.split(file)[1]
        else:
            filename = None
            file.seek(0)

        # set cx and cy from image size if not specified
        cx_px, cy_px = PIL_Image.open(file).size
        cx = cx if cx is not None else Px(cx_px)
        cy = cy if cy is not None else Px(cy_px)

        # assemble XML hierarchy of pic element
        pic = _Element('p:pic', _nsmap)
        _SubElement(pic, 'p:nvPicPr')
        _SubElement(pic.nvPicPr, 'p:cNvPr')
        pic.nvPicPr.cNvPr.set('id', str(id))
        pic.nvPicPr.cNvPr.set('name', shapename)
        if filename:
            pic.nvPicPr.cNvPr.set('descr', filename)
        _SubElement(pic.nvPicPr, 'p:cNvPicPr')
        _SubElement(pic.nvPicPr, 'p:nvPr')

        _SubElement(pic, 'p:blipFill')
        _SubElement(pic.blipFill, 'a:blip')
        pic.blipFill[qn('a:blip')].set(qn('r:embed'), rId)
        _SubElement(pic.blipFill, 'a:stretch')
        _SubElement(pic.blipFill[qn('a:stretch')], 'a:fillRect')
        _SubElement(pic, 'p:spPr')
        _SubElement(pic.spPr, 'a:xfrm')
        _SubElement(pic.spPr[qn('a:xfrm')], 'a:off')
        pic.spPr[qn('a:xfrm')].off.set('x', str(x))
        pic.spPr[qn('a:xfrm')].off.set('y', str(y))
        _SubElement(pic.spPr[qn('a:xfrm')], 'a:ext')
        pic.spPr[qn('a:xfrm')].ext.set('cx', str(cx))
        pic.spPr[qn('a:xfrm')].ext.set('cy', str(cy))
        _SubElement(pic.spPr, 'a:prstGeom')
        pic.spPr[qn('a:prstGeom')].set('prst', 'rect')
        _SubElement(pic.spPr[qn('a:prstGeom')], 'a:avLst')
        return pic

    def __sp(self, sp_id, shapename, x, y, cx, cy, is_textbox=False):
        """Return new ``<p:sp>`` element based on parameters."""
        sp = _Element('p:sp', _nsmap)
        _SubElement(sp, 'p:nvSpPr')
        _SubElement(sp.nvSpPr, 'p:cNvPr')
        sp.nvSpPr.cNvPr.set('id', str(sp_id))
        sp.nvSpPr.cNvPr.set('name', shapename)
        _SubElement(sp.nvSpPr, 'p:cNvSpPr')
        if is_textbox:
            sp.nvSpPr.cNvSpPr.set('txBox', '1')
        _SubElement(sp.nvSpPr, 'p:nvPr')

        _SubElement(sp, 'p:spPr')
        _SubElement(sp.spPr, 'a:xfrm')
        _SubElement(sp.spPr[qn('a:xfrm')], 'a:off')
        sp.spPr[qn('a:xfrm')].off.set('x', str(x))
        sp.spPr[qn('a:xfrm')].off.set('y', str(y))
        _SubElement(sp.spPr[qn('a:xfrm')], 'a:ext')
        sp.spPr[qn('a:xfrm')].ext.set('cx', str(cx))
        sp.spPr[qn('a:xfrm')].ext.set('cy', str(cy))
        _SubElement(sp.spPr, 'a:prstGeom')
        sp.spPr[qn('a:prstGeom')].set('prst', 'rect')
        _SubElement(sp.spPr[qn('a:prstGeom')], 'a:avLst')
        _SubElement(sp.spPr, 'a:noFill')

        _SubElement(sp, 'p:txBody')
        _SubElement(sp.txBody, 'a:bodyPr')
        sp.txBody[qn('a:bodyPr')].set('wrap', 'none')
        _SubElement(sp.txBody[qn('a:bodyPr')], 'a:spAutoFit')
        _SubElement(sp.txBody, 'a:lstStyle')
        _SubElement(sp.txBody, 'a:p')

        return sp


class Placeholder(object):
    """
    Decorator (pattern) class for adding placeholder properties to a shape
    that contains a placeholder element, e.g. ``<p:ph>``.
    """
    def __new__(cls, shape):
        cls = type('PlaceholderDecorator', (Placeholder, shape.__class__), {})
        return object.__new__(cls)

    def __init__(self, shape):
        self.__decorated = shape
        xpath = './*[1]/p:nvPr/p:ph'
        self.__ph = self._element.xpath(xpath, namespaces=_nsmap)[0]

    def __getattr__(self, name):
        """
        Called when *name* is not found in ``self`` or in class tree. In this
        case, delegate attribute lookup to decorated (it's probably in its
        instance namespace).
        """
        return getattr(self.__decorated, name)

    @property
    def type(self):
        """Placeholder type, e.g. PH_TYPE_CTRTITLE"""
        return self.__ph.get('type', PH_TYPE_OBJ)

    @property
    def orient(self):
        """Placeholder 'orient' attribute, e.g. PH_ORIENT_HORZ"""
        return self.__ph.get('orient', PH_ORIENT_HORZ)

    @property
    def sz(self):
        """Placeholder 'sz' attribute, e.g. PH_SZ_FULL"""
        return self.__ph.get('sz', PH_SZ_FULL)

    @property
    def idx(self):
        """Placeholder 'idx' attribute, e.g. '0'"""
        return int(self.__ph.get('idx', 0))


class Picture(BaseShape):
    """
    A picture shape, one that places an image on a slide. Corresponds to the
    ``<p:pic>`` element.
    """
    def __init__(self, pic):
        super(Picture, self).__init__(pic)


class Shape(BaseShape):
    """
    A shape that can appear on a slide. Corresponds to the ``<p:sp>`` element
    that can appear in any of the slide-type parts (slide, slideLayout,
    slideMaster, notesPage, notesMaster, handoutMaster).
    """
    def __init__(self, shape_element):
        super(Shape, self).__init__(shape_element)


class Table(BaseShape):
    """
    A table shape. Corresponds to the ``<p:graphicFrame>`` element.
    """
    def __init__(self, graphicFrame):
        super(Table, self).__init__(graphicFrame)


# ============================================================================
# Text-related classes
# ============================================================================

class TextFrame(object):
    """
    The part of a shape that contains its text. Not all shapes have a text
    frame. Corresponds to the ``<p:txBody>`` element that can appear as a
    child element of ``<p:sp>``. Not intended to be constructed directly.
    """
    def __init__(self, txBody):
        super(TextFrame, self).__init__()
        self.__txBody = txBody

    @property
    def paragraphs(self):
        """
        Immutable sequence of :class:`Paragraph` instances corresponding to
        the paragraphs in this text frame. A text frame always contains at
        least one paragraph.
        """
        xpath = './a:p'
        p_elms = self.__txBody.xpath(xpath, namespaces=_nsmap)
        paragraphs = []
        for p in p_elms:
            paragraphs.append(Paragraph(p))
        return tuple(paragraphs)

    def _set_text(self, text):
        """Replace all text in text frame with single run containing *text*"""
        self.clear()
        self.paragraphs[0].text = _to_unicode(text)

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
        bodyPr = _get_or_add(self.__txBody, 'a:bodyPr')
        bodyPr.set('anchor', value_map[value])

    #: Write-only. Assignment to *vertical_anchor* sets the vertical
    #: alignment of the text frame to top, middle, or bottom. Valid values are
    #: ``MSO.ANCHOR_TOP``, ``MSO.ANCHOR_MIDDLE``, or ``MSO.ANCHOR_BOTTOM``.
    #: The ``MSO`` name is imported from ``pptx.constants``.
    vertical_anchor = property(None, _set_vertical_anchor)

    def add_paragraph(self):
        """
        Return new |Paragraph| instance appended to the sequence of paragraphs
        contained in this text frame.
        """
        # <a:p> elements are last in txBody, so can simply append new one
        p = _Element('a:p', _nsmap)
        self.__txBody.append(p)
        return Paragraph(p)

    def clear(self):
        """
        Remove all paragraphs except one empty one.
        """
        p_list = self.__txBody.xpath('./a:p', namespaces=_nsmap)
        for p in p_list[1:]:
            self.__txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()


class _Font(object):
    """
    Character properties object, prominent among those properties being font
    size, font name, bold, italic, etc. Corresponds to ``<a:rPr>`` child
    element of a run. Also appears as ``<a:defRPr>`` and ``<a:endParaRPr>``
    in paragraph and ``<a:defRPr>`` in list style elements. Not intended to be
    constructed directly.
    """
    def __init__(self, rPr):
        super(_Font, self).__init__()
        self.__rPr = rPr

    @property
    def bold(self):
        """
        Get or set boolean bold value of |Font|, e.g.
        ``paragraph.font.bold = True``.
        """
        b = self.__rPr.get('b')
        return True if b in ('true', '1') else False

    @bold.setter
    def bold(self, bool):
        if bool:
            self.__rPr.set('b', '1')
        elif 'b' in self.__rPr.attrib:
            del self.__rPr.attrib['b']

    @property
    def italic(self):
        """
        Get or set boolean italic value of |Font|, e.g.
        ``paragraph.font.italic = True``.
        """
        i = self.__rPr.get('i')
        return True if i in ('true', '1') else False

    @italic.setter
    def italic(self, bool):
        if bool:
            self.__rPr.set('i', '1')
        elif 'i' in self.__rPr.attrib:
            del self.__rPr.attrib['i']

    def _set_size(self, centipoints):
        # handle float centipoints value gracefully
        centipoints = int(centipoints)
        self.__rPr.set('sz', str(centipoints))

    #: Set the font size. In PresentationML, font size is expressed in
    #: hundredths of a point (centipoints). The :class:`pptx.util.Pt` class
    #: allows convenient conversion to centipoints from float or integer point
    #: values, e.g. ``Pt(12.5)``. I'm pretty sure I just made up the word
    #: *centipoint*, but it seems apt :).
    size = property(None, _set_size)


class Paragraph(object):
    """
    Paragraph object. Not intended to be constructed directly.
    """
    def __init__(self, p):
        super(Paragraph, self).__init__()
        self.__p = p

    @property
    def font(self):
        """
        :class:`_Font` object containing default character properties for the
        runs in this paragraph. These character properties override default
        properties inherited from parent objects such as the text frame the
        paragraph is contained in and they may be overridden by character
        properties set at the run level.
        """
        # A _Font instance is created on first access if it doesn't exist.
        # This can cause "litter" <a:pPr> and <a:defRPr> elements to be
        # included in the XML if the _Font element is referred to but not
        # populated with values.
        if not hasattr(self.__p, 'pPr'):
            pPr = _Element('a:pPr', _nsmap)
            self.__p.insert(0, pPr)
        if not hasattr(self.__p.pPr, 'defRPr'):
            _SubElement(self.__p.pPr, 'a:defRPr')
        return _Font(self.__p.pPr.defRPr)

    def _get_level(self):
        """
        Return integer indentation level of this paragraph.
        """
        if not hasattr(self.__p, 'pPr'):
            return 0
        return int(self.__p.pPr.get('lvl', 0))

    def _set_level(self, level):
        """
        Set indentation level of this paragraph to *level*, an integer value
        between 0 and 8 inclusive.
        """
        if not isinstance(level, int) or level < 0 or level > 8:
            msg = "paragraph level must be integer between 0 and 8 inclusive"
            raise ValueError(msg)
        if not hasattr(self.__p, 'pPr'):
            pPr = _Element('a:pPr', _nsmap)
            self.__p.insert(0, pPr)
        self.__p.pPr.set('lvl', str(level))

    #: Read-write integer indentation level of this paragraph. Range is 0-8.
    #: 0 represents a top-level paragraph and is the default value. Indentation
    #: level is most commonly encountered in a bulleted list, as is found on a
    #: word bullet slide.
    level = property(_get_level, _set_level)

    @property
    def runs(self):
        """
        Immutable sequence of :class:`Run` instances corresponding to the runs
        in this paragraph.
        """
        xpath = './a:r'
        r_elms = self.__p.xpath(xpath, namespaces=_nsmap)
        runs = []
        for r in r_elms:
            runs.append(Run(r))
        return tuple(runs)

    def _set_text(self, text):
        """Replace runs with single run containing *text*"""
        self.clear()
        r = self.add_run()
        r.text = _to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the paragraph. After assignment, the paragraph containins exactly
    #: one run containing the text value of the assigned expression. The
    #: assigned value can be a 7-bit ASCII string, a UTF-8 encoded 8-bit
    #: string, or unicode. String values are converted to unicode assuming
    #: UTF-8 encoding.
    text = property(None, _set_text)

    def add_run(self):
        """Return a new run appended to the runs in this paragraph."""
        r = _Element('a:r', _nsmap)
        _SubElement(r, 'a:t')
        # work out where to insert it, ahead of a:endParaRPr if there is one
        endParaRPr = _child(self.__p, 'a:endParaRPr')
        if endParaRPr is not None:
            endParaRPr.addprevious(r)
        else:
            self.__p.append(r)
        return Run(r)

    def clear(self):
        """Remove all runs from this paragraph."""
        # retain pPr if present
        pPr = _child(self.__p, 'a:pPr')
        self.__p.clear()
        if pPr is not None:
            self.__p.insert(0, pPr)


class Run(object):
    """
    Text run object. Corresponds to ``<a:r>`` child element in a paragraph.
    """
    def __init__(self, r):
        super(Run, self).__init__()
        self.__r = r

    @property
    def font(self):
        """
        :class:`_Font` object containing run-level character properties for the
        text in this run. Character properties can and perhaps most often are
        inherited from parent objects such as the paragraph and slide layout
        the run is contained in. Only those specifically assigned at the run
        level are contained in the :class:`_Font` object.
        """
        if not hasattr(self.__r, 'rPr'):
            self.__r.insert(0, _Element('a:rPr', _nsmap))
        return _Font(self.__r.rPr)

    @property
    def text(self):
        """
        Read/Write. Text contained in the run. A regular text run is required
        to contain exactly one ``<a:t>`` (text) element. Assignment to *text*
        replaces the text currently contained in the run. The assigned value
        can be a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode.
        String values are converted to unicode assuming UTF-8 encoding.
        """
        return self.__r.t.text

    @text.setter
    def text(self, str):
        """Set the text of this run to *str*."""
        self.__r.t._setText(_to_unicode(str))
