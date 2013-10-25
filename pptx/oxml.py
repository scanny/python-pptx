# encoding: utf-8

"""
Classes that directly manipulate Open XML and provide direct object-oriented
access to the XML elements. Classes are implemented as a wrapper around their
bit of the lxml graph that spans the entire Open XML package part, e.g. a
slide.
"""

import re

from datetime import datetime, timedelta

from lxml import etree, objectify

from pptx.spec import nsmap
from pptx.spec import (
    PH_ORIENT_HORZ, PH_SZ_FULL, PH_TYPE_BODY, PH_TYPE_CTRTITLE, PH_TYPE_OBJ,
    PH_TYPE_SUBTITLE, PH_TYPE_TITLE
)


# oxml-specific constants --------------
XSD_TRUE = '1'


# configure objectified XML parser
fallback_lookup = objectify.ObjectifyElementClassLookup()
element_class_lookup = etree.ElementNamespaceClassLookup(fallback_lookup)
oxml_parser = etree.XMLParser(remove_blank_text=True)
oxml_parser.set_element_class_lookup(element_class_lookup)


# ============================================================================
# API functions
# ============================================================================

class _NamespacePrefixedTag(str):
    """
    Value object that knows the semantics of an XML tag having a namespace
    prefix.
    """
    def __new__(cls, nstag, *args):
        return super(_NamespacePrefixedTag, cls).__new__(cls, nstag)

    def __init__(self, nstag, prefix_to_uri_map):
        self._pfx, self._local_part = nstag.split(':')
        self._ns_uri = prefix_to_uri_map[self._pfx]

    @property
    def clark_name(self):
        return '{%s}%s' % (self._ns_uri, self._local_part)

    @property
    def namespace_map(self):
        return {self._pfx: self._ns_uri}


def _Element(tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag, nsmap)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return oxml_parser.makeelement(tag_name, nsmap=tag_nsmap)


def _SubElement(parent, tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag, nsmap)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return objectify.SubElement(parent, tag_name, nsmap=tag_nsmap)


def new(tag, **extra):
    return objectify.Element(qn(tag), **extra)


def nsdecls(*prefixes):
    return ' '.join(['xmlns:%s="%s"' % (pfx, nsmap[pfx]) for pfx in prefixes])


def oxml_fromstring(text):
    """``etree.fromstring()`` replacement that uses oxml parser"""
    return objectify.fromstring(text, oxml_parser)


def oxml_parse(source):
    """``etree.parse()`` replacement that uses oxml parser"""
    return objectify.parse(source, oxml_parser)


def oxml_tostring(elm, encoding=None, pretty_print=False, standalone=None):
    # if xsi parameter is not set to False, PowerPoint won't load without a
    # repair step; deannotate removes some original xsi:type tags in core.xml
    # if this parameter is left out (or set to True)
    objectify.deannotate(elm, xsi=False, cleanup_namespaces=False)
    xml = etree.tostring(
        elm, encoding=encoding, pretty_print=pretty_print,
        standalone=standalone
    )
    return xml


def qn(namespace_prefixed_tag):
    """
    Return a Clark-notation qualified tag name corresponding to
    *namespace_prefixed_tag*, a string like 'p:body'. 'qn' stands for
    *qualified name*. As an example, ``qn('p:cSld')`` returns
    ``'{http://schemas.../main}cSld'``.
    """
    nsptag = _NamespacePrefixedTag(namespace_prefixed_tag, nsmap)
    return nsptag.clark_name


def sub_elm(parent, tag, **extra):
    return objectify.SubElement(parent, qn(tag), **extra)


# ============================================================================
# utility functions
# ============================================================================

def _child(element, child_tagname):
    """
    Return direct child of *element* having *child_tagname* or |None|
    if no such child element is present.
    """
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


def _child_list(element, child_tagname):
    """
    Return list containing the direct children of *element* having
    *child_tagname*.
    """
    xpath = './%s' % child_tagname
    return element.xpath(xpath, namespaces=nsmap)


def _get_or_add(start_elm, *path_tags):
    """
    Retrieve the element at the end of the branch starting at parent and
    traversing each of *path_tags* in order, creating any elements not found
    along the way. Not a good solution when sequence of added children is
    likely to be a concern.
    """
    parent = start_elm
    for tag in path_tags:
        child = _child(parent, tag)
        if child is None:
            child = _SubElement(parent, tag)
        parent = child
    return child


# ============================================================================
# Custom element classes
# ============================================================================

class CT_CoreProperties(objectify.ObjectifiedElement):
    """
    ``<cp:coreProperties>`` element, the root element of the Core Properties
    part stored as ``/docProps/core.xml``. Implements many of the Dublin Core
    document metadata elements. String elements resolve to an empty string
    ('') if the element is not present in the XML. String elements are
    limited in length to 255 unicode characters.
    """
    _date_tags = {
        'created':      'dcterms:created',
        'last_printed': 'cp:lastPrinted',
        'modified':     'dcterms:modified',
    }
    _str_tags = {
        'author':           'dc:creator',
        'category':         'cp:category',
        'comments':         'dc:description',
        'content_status':   'cp:contentStatus',
        'identifier':       'dc:identifier',
        'keywords':         'cp:keywords',
        'language':         'dc:language',
        'last_modified_by': 'cp:lastModifiedBy',
        'subject':          'dc:subject',
        'title':            'dc:title',
        'version':          'cp:version',
    }
    _coreProperties_tmpl = (
        '<cp:coreProperties %s/>\n' % nsdecls('cp', 'dc', 'dcterms')
    )

    @staticmethod
    def new_coreProperties():
        """Return a new ``<cp:coreProperties>`` element"""
        xml = CT_CoreProperties._coreProperties_tmpl
        coreProperties = oxml_fromstring(xml)
        return coreProperties

    def __getattribute__(self, name):
        """
        Intercept attribute access to generalize property getters.
        """
        if name in CT_CoreProperties._str_tags:
            return self.__get_str_prop(name)
        elif name in CT_CoreProperties._date_tags:
            return self.__get_date_prop(name)
        elif name == 'revision':
            return self.__get_revision()
        else:
            return super(CT_CoreProperties, self).__getattribute__(name)

    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in CT_CoreProperties._str_tags:
            self.__set_str_prop(name, value)
        elif name in CT_CoreProperties._date_tags:
            self.__set_date_prop(name, value)
        elif name == 'revision':
            self.__set_revision(value)
        else:
            super(CT_CoreProperties, self).__setattr__(name, value)

    def __get_str_prop(self, name):
        """Return string value of *name* property."""
        # explicit class reference avoids another pass through getattribute
        tag = qn(CT_CoreProperties._str_tags[name])
        if not hasattr(self, tag):
            return ''
        return getattr(self, tag).text

    def __get_date_prop(self, name):
        """Return datetime value of *name* property."""
        # explicit class reference avoids another pass through getattribute
        tag = qn(CT_CoreProperties._date_tags[name])
        # date properties return None when property element not present
        if not hasattr(self, tag):
            return None
        datetime_str = getattr(self, tag).text
        try:
            return self._parse_W3CDTF_to_datetime(datetime_str)
        except ValueError:
            # invalid datetime strings are ignored
            return None

    def __get_revision(self):
        """Return integer value of revision property."""
        tag = qn('cp:revision')
        # revision returns zero when element not present
        if not hasattr(self, tag):
            return 0
        revision_str = getattr(self, tag).text
        try:
            revision = int(revision_str)
        except ValueError:
            # non-integer revision strings also resolve to 0
            revision = 0
        # as do negative integers
        if revision < 0:
            revision = 0
        return revision

    def __set_str_prop(self, name, value):
        """Set string value of *name* property to *value*"""
        value = str(value)
        if len(value) > 255:
            tmpl = ("exceeded 255 char max length of property '%s', got:"
                    "\n\n'%s'")
            raise ValueError(tmpl % (name, value))
        tag = qn(CT_CoreProperties._str_tags[name])
        setattr(self, tag, value)
        # objectify will leave in a py: namespace without this cleanup
        elm = getattr(self, tag)
        objectify.deannotate(elm, cleanup_namespaces=True)

    def __set_date_prop(self, name, value):
        """Set datetime value of *name* property to *value*"""
        if not isinstance(value, datetime):
            tmpl = ("'%s' property requires <type 'datetime.datetime'> objec"
                    "t, got %s")
            raise ValueError(tmpl % (name, type(value)))
        tagname = CT_CoreProperties._date_tags[name]
        tag = qn(tagname)
        dt_str = value.strftime('%Y-%m-%dT%H:%M:%SZ')
        setattr(self, tag, dt_str)
        # objectify will leave in a py: namespace without this cleanup
        elm = getattr(self, tag)
        objectify.deannotate(elm, cleanup_namespaces=True)
        if name in ('created', 'modified'):
            # these two require an explicit 'xsi:type' attribute
            # first and last line are a hack required to add the xsi
            # namespace to the root element rather than each child element
            # in which it is referenced
            self.set(qn('xsi:foo'), 'bar')
            self[tag].set(qn('xsi:type'), 'dcterms:W3CDTF')
            del self.attrib[qn('xsi:foo')]

    def __set_revision(self, value):
        """Set integer value of revision property to *value*"""
        if not isinstance(value, int) or value < 1:
            tmpl = "revision property requires positive int, got '%s'"
            raise ValueError(tmpl % value)
        tag = qn('cp:revision')
        setattr(self, tag, str(value))
        # objectify will leave in a py: namespace without this cleanup
        elm = getattr(self, tag)
        objectify.deannotate(elm, cleanup_namespaces=True)

    _offset_pattern = re.compile('([+-])(\d\d):(\d\d)')

    @classmethod
    def _offset_dt(cls, dt, offset_str):
        """
        Return a |datetime| instance that is offset from datetime *dt* by
        the timezone offset specified in *offset_str*, a string like
        ``'-07:00'``.
        """
        match = cls._offset_pattern.match(offset_str)
        if match is None:
            raise ValueError("'%s' is not a valid offset string" % offset_str)
        sign, hours_str, minutes_str = match.groups()
        sign_factor = -1 if sign == '+' else 1
        hours = int(hours_str) * sign_factor
        minutes = int(minutes_str) * sign_factor
        td = timedelta(hours=hours, minutes=minutes)
        return dt + td

    @classmethod
    def _parse_W3CDTF_to_datetime(cls, w3cdtf_str):
        # valid W3CDTF date cases:
        # yyyy e.g. '2003'
        # yyyy-mm e.g. '2003-12'
        # yyyy-mm-dd e.g. '2003-12-31'
        # UTC timezone e.g. '2003-12-31T10:14:55Z'
        # numeric timezone e.g. '2003-12-31T10:14:55-08:00'
        templates = (
            '%Y-%m-%dT%H:%M:%S',
            '%Y-%m-%d',
            '%Y-%m',
            '%Y',
        )
        # strptime isn't smart enough to parse literal timezone offsets like
        # '-07:30', so we have to do it ourselves
        parseable_part = w3cdtf_str[:19]
        offset_str = w3cdtf_str[19:]
        dt = None
        for tmpl in templates:
            try:
                dt = datetime.strptime(parseable_part, tmpl)
            except ValueError:
                continue
        if dt is None:
            tmpl = "could not parse W3CDTF datetime string '%s'"
            raise ValueError(tmpl % w3cdtf_str)
        if len(offset_str) == 6:
            return cls._offset_dt(dt, offset_str)
        return dt


class CT_GraphicalObjectFrame(objectify.ObjectifiedElement):
    """
    ``<p:graphicFrame>`` element, which is a container for a table, a chart,
    or another graphical object.
    """
    DATATYPE_TABLE = 'http://schemas.openxmlformats.org/drawingml/2006/table'

    _graphicFrame_tmpl = (
        '<p:graphicFrame %s>\n'
        '  <p:nvGraphicFramePr>\n'
        '    <p:cNvPr id="%s" name="%s"/>\n'
        '    <p:cNvGraphicFramePr>\n'
        '      <a:graphicFrameLocks noGrp="1"/>\n'
        '    </p:cNvGraphicFramePr>\n'
        '    <p:nvPr/>\n'
        '  </p:nvGraphicFramePr>\n'
        '  <p:xfrm>\n'
        '    <a:off x="%s" y="%s"/>\n'
        '    <a:ext cx="%s" cy="%s"/>\n'
        '  </p:xfrm>\n'
        '  <a:graphic>\n'
        '    <a:graphicData/>\n'
        '  </a:graphic>\n'
        '</p:graphicFrame>' %
        (nsdecls('a', 'p'), '%d', '%s', '%d', '%d', '%d', '%d')
    )

    @property
    def has_table(self):
        """True if graphicFrame contains a table, False otherwise"""
        datatype = self[qn('a:graphic')].graphicData.get('uri')
        if datatype == CT_GraphicalObjectFrame.DATATYPE_TABLE:
            return True
        return False

    @staticmethod
    def new_graphicFrame(id_, name, left, top, width, height):
        """
        Return a new ``<p:graphicFrame>`` element tree suitable for containing
        a table or chart. Note that a graphicFrame element is not a valid
        shape until it contains a graphical object such as a table.
        """
        xml = CT_GraphicalObjectFrame._graphicFrame_tmpl % (
            id_, name, left, top, width, height)
        graphicFrame = oxml_fromstring(xml)

        objectify.deannotate(graphicFrame, cleanup_namespaces=True)
        return graphicFrame

    @staticmethod
    def new_table(id_, name, rows, cols, left, top, width, height):
        """
        Return a ``<p:graphicFrame>`` element tree populated with a table
        element.
        """
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, left, top, width, height)

        # set type of contained graphic to table
        graphicData = graphicFrame[qn('a:graphic')].graphicData
        graphicData.set('uri', CT_GraphicalObjectFrame.DATATYPE_TABLE)

        # add tbl element tree
        tbl = CT_Table.new_tbl(rows, cols, width, height)
        graphicData.append(tbl)

        objectify.deannotate(graphicFrame, cleanup_namespaces=True)
        return graphicFrame


class CT_Picture(objectify.ObjectifiedElement):
    """
    ``<p:pic>`` element, which represents a picture shape (an image placement
    on a slide).
    """
    _pic_tmpl = (
        '<p:pic %s>\n'
        '  <p:nvPicPr>\n'
        '    <p:cNvPr id="%s" name="%s" descr="%s"/>\n'
        '    <p:cNvPicPr>\n'
        '      <a:picLocks noChangeAspect="1"/>\n'
        '    </p:cNvPicPr>\n'
        '    <p:nvPr/>\n'
        '  </p:nvPicPr>\n'
        '  <p:blipFill>\n'
        '    <a:blip r:embed="%s"/>\n'
        '    <a:stretch>\n'
        '      <a:fillRect/>\n'
        '    </a:stretch>\n'
        '  </p:blipFill>\n'
        '  <p:spPr>\n'
        '    <a:xfrm>\n'
        '      <a:off x="%s" y="%s"/>\n'
        '      <a:ext cx="%s" cy="%s"/>\n'
        '    </a:xfrm>\n'
        '    <a:prstGeom prst="rect">\n'
        '      <a:avLst/>\n'
        '    </a:prstGeom>\n'
        '  </p:spPr>\n'
        '</p:pic>' % (nsdecls('a', 'p', 'r'), '%d', '%s', '%s', '%s',
                      '%d', '%d', '%d', '%d')
    )

    @staticmethod
    def new_pic(id_, name, desc, rId, left, top, width, height):
        """
        Return a new ``<p:pic>`` element tree configured with the supplied
        parameters.
        """
        xml = CT_Picture._pic_tmpl % (id_, name, desc, rId,
                                      left, top, width, height)
        pic = oxml_fromstring(xml)

        objectify.deannotate(pic, cleanup_namespaces=True)
        return pic


class CT_PresetGeometry2D(objectify.ObjectifiedElement):
    """<a:prstGeom> custom element class"""
    @property
    def gd(self):
        """
        Sequence containing the ``gd`` element children of ``<a:avLst>``
        child element, empty if none are present.
        """
        try:
            gd_elms = tuple([gd for gd in self.avLst.gd])
        except AttributeError:
            gd_elms = ()
        return gd_elms

    @property
    def prst(self):
        """Value of required ``prst`` attribute."""
        return self.get('prst')

    def rewrite_guides(self, guides):
        """
        Remove any ``<a:gd>`` element children of ``<a:avLst>`` and replace
        them with ones having (name, val) in *guides*.
        """
        try:
            avLst = self.avLst
        except AttributeError:
            avLst = _SubElement(self, 'a:avLst')
        if hasattr(self.avLst, 'gd'):
            for gd_elm in self.avLst.gd[:]:
                avLst.remove(gd_elm)
        for name, val in guides:
            gd = _SubElement(avLst, 'a:gd')
            gd.set('name', name)
            gd.set('fmla', 'val %d' % val)


class CT_Shape(objectify.ObjectifiedElement):
    """<p:sp> custom element class"""
    _autoshape_sp_tmpl = (
        '<p:sp %s>\n'
        '  <p:nvSpPr>\n'
        '    <p:cNvPr id="%s" name="%s"/>\n'
        '    <p:cNvSpPr/>\n'
        '    <p:nvPr/>\n'
        '  </p:nvSpPr>\n'
        '  <p:spPr>\n'
        '    <a:xfrm>\n'
        '      <a:off x="%s" y="%s"/>\n'
        '      <a:ext cx="%s" cy="%s"/>\n'
        '    </a:xfrm>\n'
        '    <a:prstGeom prst="%s">\n'
        '      <a:avLst/>\n'
        '    </a:prstGeom>\n'
        '  </p:spPr>\n'
        '  <p:style>\n'
        '    <a:lnRef idx="1">\n'
        '      <a:schemeClr val="accent1"/>\n'
        '    </a:lnRef>\n'
        '    <a:fillRef idx="3">\n'
        '      <a:schemeClr val="accent1"/>\n'
        '    </a:fillRef>\n'
        '    <a:effectRef idx="2">\n'
        '      <a:schemeClr val="accent1"/>\n'
        '    </a:effectRef>\n'
        '    <a:fontRef idx="minor">\n'
        '      <a:schemeClr val="lt1"/>\n'
        '    </a:fontRef>\n'
        '  </p:style>\n'
        '  <p:txBody>\n'
        '    <a:bodyPr rtlCol="0" anchor="ctr"/>\n'
        '    <a:lstStyle/>\n'
        '    <a:p>\n'
        '      <a:pPr algn="ctr"/>\n'
        '    </a:p>\n'
        '  </p:txBody>\n'
        '</p:sp>' %
        (nsdecls('a', 'p'), '%d', '%s', '%d', '%d', '%d', '%d', '%s')
    )

    _ph_sp_tmpl = (
        '<p:sp %s>\n'
        '  <p:nvSpPr>\n'
        '    <p:cNvPr id="%s" name="%s"/>\n'
        '    <p:cNvSpPr/>\n'
        '    <p:nvPr/>\n'
        '  </p:nvSpPr>\n'
        '  <p:spPr/>\n'
        '</p:sp>' % (nsdecls('a', 'p'), '%d', '%s')
    )

    _textbox_sp_tmpl = (
        '<p:sp %s>\n'
        '  <p:nvSpPr>\n'
        '    <p:cNvPr id="%s" name="%s"/>\n'
        '    <p:cNvSpPr txBox="1"/>\n'
        '    <p:nvPr/>\n'
        '  </p:nvSpPr>\n'
        '  <p:spPr>\n'
        '    <a:xfrm>\n'
        '      <a:off x="%s" y="%s"/>\n'
        '      <a:ext cx="%s" cy="%s"/>\n'
        '    </a:xfrm>\n'
        '    <a:prstGeom prst="rect">\n'
        '      <a:avLst/>\n'
        '    </a:prstGeom>\n'
        '    <a:noFill/>\n'
        '  </p:spPr>\n'
        '  <p:txBody>\n'
        '    <a:bodyPr wrap="none">\n'
        '      <a:spAutoFit/>\n'
        '    </a:bodyPr>\n'
        '    <a:lstStyle/>\n'
        '    <a:p/>\n'
        '  </p:txBody>\n'
        '</p:sp>' % (nsdecls('a', 'p'), '%d', '%s', '%d', '%d', '%d', '%d')
    )

    @property
    def is_autoshape(self):
        """
        True if this shape is an auto shape. A shape is an auto shape if it
        has a ``<a:prstGeom>`` element and does not have a txBox="1" attribute
        on cNvSpPr.
        """
        prstGeom = _child(self.spPr, 'a:prstGeom')
        if prstGeom is None:
            return False
        txBox = self.nvSpPr.cNvSpPr.get('txBox')
        if txBox in ('true', '1'):
            return False
        return True

    @property
    def is_textbox(self):
        """
        True if this shape is a text box. A shape is a text box if it has a
        txBox="1" attribute on cNvSpPr.
        """
        txBox = self.nvSpPr.cNvSpPr.get('txBox')
        if txBox in ('true', '1'):
            return True
        return False

    @staticmethod
    def new_autoshape_sp(id_, name, prst, left, top, width, height):
        """
        Return a new ``<p:sp>`` element tree configured as a base auto shape.
        """
        xml = CT_Shape._autoshape_sp_tmpl % (id_, name, left, top,
                                             width, height, prst)
        sp = oxml_fromstring(xml)
        objectify.deannotate(sp, cleanup_namespaces=True)
        return sp

    @staticmethod
    def new_placeholder_sp(id_, name, ph_type, orient, sz, idx):
        """
        Return a new ``<p:sp>`` element tree configured as a placeholder
        shape.
        """
        xml = CT_Shape._ph_sp_tmpl % (id_, name)
        sp = oxml_fromstring(xml)

        # placeholder shapes get a "no group" lock
        _SubElement(sp.nvSpPr.cNvSpPr, 'a:spLocks')
        sp.nvSpPr.cNvSpPr[qn('a:spLocks')].set('noGrp', '1')

        # placeholder (ph) element attributes values vary by type
        ph = _SubElement(sp.nvSpPr.nvPr, 'p:ph')
        if ph_type != PH_TYPE_OBJ:
            ph.set('type', ph_type)
        if orient != PH_ORIENT_HORZ:
            ph.set('orient', orient)
        if sz != PH_SZ_FULL:
            ph.set('sz', sz)
        if idx != 0:
            ph.set('idx', str(idx))

        placeholder_types_that_have_a_text_frame = (
            PH_TYPE_TITLE, PH_TYPE_CTRTITLE, PH_TYPE_SUBTITLE, PH_TYPE_BODY,
            PH_TYPE_OBJ)

        if ph_type in placeholder_types_that_have_a_text_frame:
            sp.append(CT_TextBody.new_txBody())

        objectify.deannotate(sp, cleanup_namespaces=True)
        return sp

    @staticmethod
    def new_textbox_sp(id_, name, left, top, width, height):
        """
        Return a new ``<p:sp>`` element tree configured as a base textbox
        shape.
        """
        xml = CT_Shape._textbox_sp_tmpl % (id_, name, left, top, width, height)
        sp = oxml_fromstring(xml)
        objectify.deannotate(sp, cleanup_namespaces=True)
        return sp

    @property
    def prst(self):
        """
        Value of ``prst`` attribute of ``<a:prstGeom>`` element or |None| if
        not present.
        """
        prstGeom = _child(self.spPr, 'a:prstGeom')
        if prstGeom is None:
            return None
        return prstGeom.get('prst')

    @property
    def prstGeom(self):
        """
        Reference to ``<a:prstGeom>`` child element or |None| if this shape
        doesn't have one, for example, if it's a placeholder shape.
        """
        return _child(self.spPr, 'a:prstGeom')


class CT_Table(objectify.ObjectifiedElement):
    """``<a:tbl>`` custom element class"""
    _tbl_tmpl = (
        '<a:tbl %s>\n'
        '  <a:tblPr firstRow="1" bandRow="1">\n'
        '    <a:tableStyleId>%s</a:tableStyleId>\n'
        '  </a:tblPr>\n'
        '  <a:tblGrid/>\n'
        '</a:tbl>' % (nsdecls('a'), '%s')
    )

    BOOLPROPS = (
        'bandCol', 'bandRow', 'firstCol', 'firstRow', 'lastCol', 'lastRow'
    )

    def __getattr__(self, attr):
        """
        Implement getter side of properties. Filters ``__getattr__`` messages
        to ObjectifiedElement base class to intercept messages intended for
        custom property getters.
        """
        if attr in CT_Table.BOOLPROPS:
            return self._get_boolean_property(attr)
        else:
            return super(CT_Table, self).__getattr__(attr)

    def __setattr__(self, attr, value):
        """
        Implement setter side of properties. Filters ``__setattr__`` messages
        to ObjectifiedElement base class to intercept messages intended for
        custom property setters.
        """
        if attr in CT_Table.BOOLPROPS:
            self._set_boolean_property(attr, value)
        else:
            super(CT_Table, self).__setattr__(attr, value)

    def _get_boolean_property(self, propname):
        """
        Generalized getter for the boolean properties on the ``<a:tblPr>``
        child element. Defaults to False if *propname* attribute is missing
        or ``<a:tblPr>`` element itself is not present.
        """
        if not self.has_tblPr:
            return False
        return self.tblPr.get(propname) in ('1', 'true')

    def _set_boolean_property(self, propname, value):
        """
        Generalized setter for boolean properties on the ``<a:tblPr>`` child
        element, setting *propname* attribute appropriately based on *value*.
        If *value* is truthy, the attribute is set to "1"; a tblPr child
        element is added if necessary. If *value* is falsey, the *propname*
        attribute is removed if present, allowing its default value of False
        to be its effective value.
        """
        if value:
            tblPr = self._get_or_insert_tblPr()
            tblPr.set(propname, XSD_TRUE)
        elif not self.has_tblPr:
            pass
        elif propname in self.tblPr.attrib:
            del self.tblPr.attrib[propname]

    @property
    def has_tblPr(self):
        """
        True if this ``<a:tbl>`` element has a ``<a:tblPr>`` child element,
        False otherwise.
        """
        try:
            self.tblPr
            return True
        except AttributeError:
            return False

    def _get_or_insert_tblPr(self):
        """Return tblPr child element, inserting a new one if not present"""
        if not self.has_tblPr:
            tblPr = _Element('a:tblPr')
            self.insert(0, tblPr)
        return self.tblPr

    @staticmethod
    def new_tbl(rows, cols, width, height, tableStyleId=None):
        """Return a new ``<p:tbl>`` element tree"""
        # working hypothesis is this is the default table style GUID
        if tableStyleId is None:
            tableStyleId = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

        xml = CT_Table._tbl_tmpl % (tableStyleId)
        tbl = oxml_fromstring(xml)

        # add specified number of rows and columns
        rowheight = height/rows
        colwidth = width/cols

        for col in range(cols):
            # adjust width of last col to absorb any div error
            if col == cols-1:
                colwidth = width - ((cols-1) * colwidth)
            sub_elm(tbl.tblGrid, 'a:gridCol', w=str(colwidth))

        for row in range(rows):
            # adjust height of last row to absorb any div error
            if row == rows-1:
                rowheight = height - ((rows-1) * rowheight)
            tr = sub_elm(tbl, 'a:tr', h=str(rowheight))
            for col in range(cols):
                tr.append(CT_TableCell.new_tc())

        objectify.deannotate(tbl, cleanup_namespaces=True)
        return tbl


class CT_TableCell(objectify.ObjectifiedElement):
    """``<a:tc>`` custom element class"""
    _tc_tmpl = (
        '<a:tc %s>\n'
        '  <a:txBody>\n'
        '    <a:bodyPr/>\n'
        '    <a:lstStyle/>\n'
        '    <a:p/>\n'
        '  </a:txBody>\n'
        '  <a:tcPr/>\n'
        '</a:tc>' % nsdecls('a')
    )

    @staticmethod
    def new_tc():
        """Return a new ``<a:tc>`` element tree"""
        xml = CT_TableCell._tc_tmpl
        tc = oxml_fromstring(xml)
        objectify.deannotate(tc, cleanup_namespaces=True)
        return tc

    @property
    def anchor(self):
        """
        String held in ``anchor`` attribute of ``<a:tcPr>`` child element of
        this ``<a:tc>`` element.
        """
        if not hasattr(self, 'tcPr'):
            return None
        return self.tcPr.get('anchor')

    def _set_anchor(self, anchor):
        """
        Set value of anchor attribute on ``<a:tcPr>`` child element
        """
        if anchor is None:
            return self._clear_anchor()
        if not hasattr(self, 'tcPr'):
            tcPr = _Element('a:tcPr')
            idx = 1 if hasattr(self, 'txBody') else 0
            self.insert(idx, tcPr)
        self.tcPr.set('anchor', anchor)

    def _clear_anchor(self):
        """
        Remove anchor attribute from ``<a:tcPr>`` if it exists and remove
        ``<a:tcPr>`` element if it then has no attributes.
        """
        if not hasattr(self, 'tcPr'):
            return
        if 'anchor' in self.tcPr.attrib:
            del self.tcPr.attrib['anchor']
        if len(self.tcPr.attrib) == 0:
            self.remove(self.tcPr)

    def __get_marX(self, attr_name, default):
        """generalized method to get margin values"""
        if not hasattr(self, 'tcPr'):
            return default
        return int(self.tcPr.get(attr_name, default))

    @property
    def marT(self):
        """
        Read/write integer top margin value represented in ``marT`` attribute
        of the ``<a:tcPr>`` child element of this ``<a:tc>`` element. If the
        attribute is not present, the default value ``45720`` (0.05 inches)
        is returned for top and bottom; ``91440`` (0.10 inches) is the
        default for left and right. Assigning |None| to any ``marX``
        property clears that attribute from the element, effectively setting
        it to the default value.
        """
        return self.__get_marX('marT', 45720)

    @property
    def marR(self):
        """right margin value represented in ``marR`` attribute"""
        return self.__get_marX('marR', 91440)

    @property
    def marB(self):
        """bottom margin value represented in ``marB`` attribute"""
        return self.__get_marX('marB', 45720)

    @property
    def marL(self):
        """left margin value represented in ``marL`` attribute"""
        return self.__get_marX('marL', 91440)

    def _set_marX(self, marX, value):
        """
        Set value of marX attribute on ``<a:tcPr>`` child element. If *marX*
        is |None|, the marX attribute is removed and the ``<a:tcPr>`` element
        is removed if it then has no attributes.
        """
        if value is None:
            return self.__clear_marX(marX)
        if not hasattr(self, 'tcPr'):
            tcPr = _Element('a:tcPr')
            idx = 1 if hasattr(self, 'txBody') else 0
            self.insert(idx, tcPr)
        self.tcPr.set(marX, str(value))

    def __clear_marX(self, marX):
        """
        Remove marX attribute from ``<a:tcPr>`` if it exists and remove
        ``<a:tcPr>`` element if it then has no attributes.
        """
        if not hasattr(self, 'tcPr'):
            return
        if marX in self.tcPr.attrib:
            del self.tcPr.attrib[marX]
        if len(self.tcPr.attrib) == 0:
            self.remove(self.tcPr)

    def __setattr__(self, attr, value):
        """
        This hack is needed to make setter side of properties work,
        overrides ``__setattr__`` defined in ObjectifiedElement super class
        just enough to route messages intended for custom property setters.
        """
        if attr == 'anchor':
            self._set_anchor(value)
        elif attr in ('marT', 'marR', 'marB', 'marL'):
            self._set_marX(attr, value)
        else:
            super(CT_TableCell, self).__setattr__(attr, value)


class CT_TextBody(objectify.ObjectifiedElement):
    """<p:txBody> custom element class"""
    _txBody_tmpl = (
        '<p:txBody %s>\n'
        '  <a:bodyPr/>\n'
        '  <a:lstStyle/>\n'
        '  <a:p/>\n'
        '</p:txBody>\n' % (nsdecls('a', 'p'))
    )

    @staticmethod
    def new_txBody():
        """Return a new ``<p:txBody>`` element tree"""
        xml = CT_TextBody._txBody_tmpl
        txBody = oxml_fromstring(xml)
        objectify.deannotate(txBody, cleanup_namespaces=True)
        return txBody


class CT_TextParagraph(objectify.ObjectifiedElement):
    """<a:p> custom element class"""
    def get_algn(self):
        """
        Paragraph horizontal alignment value, like ``TAT.CENTER``. Value of
        algn attribute on <a:pPr> child element
        """
        if not hasattr(self, 'pPr'):
            return None
        return self.pPr.get('algn')

    def set_algn(self, value):
        """
        Set value of algn attribute on <a:pPr> child element
        """
        if value is None:
            return self._clear_algn()
        if not hasattr(self, 'pPr'):
            pPr = _Element('a:pPr')
            self.insert(0, pPr)
        self.pPr.set('algn', value)

    def _clear_algn(self):
        """
        Remove algn attribute from ``<a:pPr>`` if it exists and remove
        ``<a:pPr>`` element if it then has no attributes.
        """
        if not hasattr(self, 'pPr'):
            return
        if 'algn' in self.pPr.attrib:
            del self.pPr.attrib['algn']
        if len(self.pPr.attrib) == 0:
            self.remove(self.pPr)

    # def __setattr__(self, attr, value):
    #     """
    #     This hack is needed to override ``__setattr__`` defined in
    #     ObjectifiedElement super class.
    #     """
    #     if attr == 'algn':
    #         self._set_algn(value)
    #     else:
    #         super(CT_TextParagraph, self).__setattr__(attr, value)


a_namespace = element_class_lookup.get_namespace(nsmap['a'])
a_namespace['p'] = CT_TextParagraph
a_namespace['prstGeom'] = CT_PresetGeometry2D
a_namespace['tbl'] = CT_Table
a_namespace['tc'] = CT_TableCell

a_namespace = element_class_lookup.get_namespace(nsmap['cp'])
a_namespace['coreProperties'] = CT_CoreProperties

p_namespace = element_class_lookup.get_namespace(nsmap['p'])
p_namespace['graphicFrame'] = CT_GraphicalObjectFrame
p_namespace['pic'] = CT_Picture
p_namespace['sp'] = CT_Shape
p_namespace['txBody'] = CT_TextBody
