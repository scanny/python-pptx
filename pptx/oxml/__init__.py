# encoding: utf-8

"""
Initialize lxml parser.
"""

from __future__ import absolute_import

from lxml import etree, objectify


# oxml-specific constants --------------
XSD_TRUE = '1'


# configure objectified XML parser
fallback_lookup = objectify.ObjectifyElementClassLookup()
element_class_lookup = etree.ElementNamespaceClassLookup(fallback_lookup)
oxml_parser = etree.XMLParser(remove_blank_text=True)
oxml_parser.set_element_class_lookup(element_class_lookup)


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


def child(element, child_tagname):
    """
    Return direct child of *element* having *child_tagname* or |None|
    if no such child element is present.
    """
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


def Element(tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag, nsmap)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return oxml_parser.makeelement(tag_name, nsmap=tag_nsmap)


def namespaces(*prefixes):
    """
    Return a dict containing the subset namespace prefix mappings specified by
    *prefixes*. Any number of namespace prefixes can be supplied, e.g.
    namespaces('a', 'r', 'p').
    """
    namespaces = {}
    for prefix in prefixes:
        namespaces[prefix] = nsmap[prefix]
    return namespaces


def nsdecls(*prefixes):
    return ' '.join(['xmlns:%s="%s"' % (pfx, nsmap[pfx]) for pfx in prefixes])


def oxml_fromstring(text):
    """``etree.fromstring()`` replacement that uses oxml parser"""
    return objectify.fromstring(text, oxml_parser)


def oxml_parse(source):
    """``etree.parse()`` replacement that uses oxml parser"""
    return objectify.parse(source, oxml_parser)


def qn(namespace_prefixed_tag):
    """
    Return a Clark-notation qualified tag name corresponding to
    *namespace_prefixed_tag*, a string like 'p:body'. 'qn' stands for
    *qualified name*. As an example, ``qn('p:cSld')`` returns
    ``'{http://schemas.../main}cSld'``.
    """
    nsptag = _NamespacePrefixedTag(namespace_prefixed_tag, nsmap)
    return nsptag.clark_name


def SubElement(parent, tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag, nsmap)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return objectify.SubElement(parent, tag_name, nsmap=tag_nsmap)


def sub_elm(parent, tag, **extra):
    return objectify.SubElement(parent, qn(tag), **extra)


# ============================================================================
# nsmap
# ============================================================================
# namespace prefix to namespace name map
# ============================================================================

nsmap = {
    'a':   ('http://schemas.openxmlformats.org/drawingml/2006/main'),
    'cp':  ('http://schemas.openxmlformats.org/package/2006/metadata/core-pro'
            'perties'),
    'ct':  ('http://schemas.openxmlformats.org/package/2006/content-types'),
    'dc':  ('http://purl.org/dc/elements/1.1/'),
    'dcmitype': ('http://purl.org/dc/dcmitype/'),
    'dcterms':  ('http://purl.org/dc/terms/'),
    'ep':  ('http://schemas.openxmlformats.org/officeDocument/2006/extended-p'
            'roperties'),
    'i':   ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
            'ips/image'),
    'm':   ('http://schemas.openxmlformats.org/officeDocument/2006/math'),
    'mo':  ('http://schemas.microsoft.com/office/mac/office/2008/main'),
    'mv':  ('urn:schemas-microsoft-com:mac:vml'),
    'o':   ('urn:schemas-microsoft-com:office:office'),
    'p':   ('http://schemas.openxmlformats.org/presentationml/2006/main'),
    'pd':  ('http://schemas.openxmlformats.org/drawingml/2006/presentationDra'
            'wing'),
    'pic': ('http://schemas.openxmlformats.org/drawingml/2006/picture'),
    'pr':  ('http://schemas.openxmlformats.org/package/2006/relationships'),
    'r':   ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
            'ips'),
    'sl':  ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
            'ips/slideLayout'),
    'v':   ('urn:schemas-microsoft-com:vml'),
    've':  ('http://schemas.openxmlformats.org/markup-compatibility/2006'),
    'w':   ('http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
    'w10': ('urn:schemas-microsoft-com:office:word'),
    'wne': ('http://schemas.microsoft.com/office/word/2006/wordml'),
    'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingD'
            'rawing'),
    'xsi': ('http://www.w3.org/2001/XMLSchema-instance')
}
