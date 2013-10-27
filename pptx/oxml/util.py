# encoding: utf-8

"""
Classes that directly manipulate Open XML and provide direct object-oriented
access to the XML elements. Classes are implemented as a wrapper around their
bit of the lxml graph that spans the entire Open XML package part, e.g. a
slide.
"""

from __future__ import absolute_import


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
