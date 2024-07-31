# encoding: utf-8

"""
Namespace related objects.
"""

from __future__ import absolute_import


#: Maps namespace prefix to namespace name for all known PowerPoint XML
#: namespaces.
_nsmap = {
    "a": ("http://schemas.openxmlformats.org/drawingml/2006/main"),
    "c": ("http://schemas.openxmlformats.org/drawingml/2006/chart"),
    "cp": (
        "http://schemas.openxmlformats.org/package/2006/metadata/core-pro" "perties"
    ),
    "ct": ("http://schemas.openxmlformats.org/package/2006/content-types"),
    "dc": ("http://purl.org/dc/elements/1.1/"),
    "dcmitype": ("http://purl.org/dc/dcmitype/"),
    "dcterms": ("http://purl.org/dc/terms/"),
    "ep": (
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-p" "roperties"
    ),
    "i": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationsh" "ips/image"
    ),
    "m": ("http://schemas.openxmlformats.org/officeDocument/2006/math"),
    "mo": ("http://schemas.microsoft.com/office/mac/office/2008/main"),
    "mv": ("urn:schemas-microsoft-com:mac:vml"),
    "o": ("urn:schemas-microsoft-com:office:office"),
    "p": ("http://schemas.openxmlformats.org/presentationml/2006/main"),
    "pd": ("http://schemas.openxmlformats.org/drawingml/2006/presentationDra" "wing"),
    "pic": ("http://schemas.openxmlformats.org/drawingml/2006/picture"),
    "pr": ("http://schemas.openxmlformats.org/package/2006/relationships"),
    "r": ("http://schemas.openxmlformats.org/officeDocument/2006/relationsh" "ips"),
    "sl": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationsh"
        "ips/slideLayout"
    ),
    "v": ("urn:schemas-microsoft-com:vml"),
    "ve": ("http://schemas.openxmlformats.org/markup-compatibility/2006"),
    "w": ("http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
    "w10": ("urn:schemas-microsoft-com:office:word"),
    "wne": ("http://schemas.microsoft.com/office/word/2006/wordml"),
    "wp": ("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingD" "rawing"),
    "xsi": ("http://www.w3.org/2001/XMLSchema-instance"),
}


class NamespacePrefixedTag(str):
    """
    Value object that knows the semantics of an XML tag having a namespace
    prefix.
    """

    def __new__(cls, nstag, *args):
        return super(NamespacePrefixedTag, cls).__new__(cls, nstag)

    def __init__(self, nstag):
        self._pfx, self._local_part = nstag.split(":")
        self._ns_uri = _nsmap[self._pfx]

    @property
    def clark_name(self):
        return "{%s}%s" % (self._ns_uri, self._local_part)

    @property
    def local_part(self):
        """
        Return the local part of the tag as a string. E.g. 'foobar' is
        returned for tag 'f:foobar'.
        """
        return self._local_part

    @property
    def nsmap(self):
        """
        Return a dict having a single member, mapping the namespace prefix of
        this tag to it's namespace name (e.g. {'f': 'http://foo/bar'}). This
        is handy for passing to xpath calls and other uses.
        """
        return {self._pfx: self._ns_uri}

    @property
    def nspfx(self):
        """
        Return the string namespace prefix for the tag, e.g. 'f' is returned
        for tag 'f:foobar'.
        """
        return self._pfx

    @property
    def nsuri(self):
        """
        Return the namespace URI for the tag, e.g. 'http://foo/bar' would be
        returned for tag 'f:foobar' if the 'f' prefix maps to
        'http://foo/bar' in _nsmap.
        """
        return self._ns_uri


def namespaces(*prefixes):
    """
    Return a dict containing the subset namespace prefix mappings specified by
    *prefixes*. Any number of namespace prefixes can be supplied, e.g.
    namespaces('a', 'r', 'p').
    """
    namespaces = {}
    for prefix in prefixes:
        namespaces[prefix] = _nsmap[prefix]
    return namespaces


nsmap = namespaces  # alias for more compact use with Element()


def nsdecls(*prefixes):
    return " ".join(['xmlns:%s="%s"' % (pfx, _nsmap[pfx]) for pfx in prefixes])


def nsuri(nspfx):
    """
    Return the namespace URI corresponding to *nspfx*. For example, it would
    return 'http://foo/bar' for an *nspfx* of 'f' if the 'f' prefix maps to
    'http://foo/bar' in _nsmap.
    """
    return _nsmap[nspfx]


def qn(namespace_prefixed_tag):
    """
    Return a Clark-notation qualified tag name corresponding to
    *namespace_prefixed_tag*, a string like 'p:body'. 'qn' stands for
    *qualified name*. As an example, ``qn('p:cSld')`` returns
    ``'{http://schemas.../main}cSld'``.
    """
    nsptag = NamespacePrefixedTag(namespace_prefixed_tag)
    return nsptag.clark_name
