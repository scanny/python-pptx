# encoding: utf-8

"""
General purpose functions that raise the abstraction level of interacting with
lxml.objectify elements.
"""

from __future__ import absolute_import

import itertools
import re

from lxml import etree, objectify

from . import oxml_parser
from .ns import NamespacePrefixedTag, qn


def child(element, child_tag_str):
    """
    Return the first direct child of *element* having tag matching
    *child_tag_str* or |None| if no such child element is present.
    """
    nsptag = NamespacePrefixedTag(child_tag_str)
    xpath = './%s' % child_tag_str
    matching_children = element.xpath(xpath, namespaces=nsptag.nsmap)
    return matching_children[0] if len(matching_children) else None


def Element(nsptag_str, nsmap=None):
    """
    Return a 'loose' lxml element having the tag specified by *nsptag_str*.
    *nsptag_str* must contain the standard namespace prefix, e.g. 'a:tbl'.
    The resulting element is an instance of the custom element class for this
    tag name if one is defined.
    """
    nsptag = NamespacePrefixedTag(nsptag_str)
    nsmap = nsmap if nsmap is not None else nsptag.nsmap
    return oxml_parser.makeelement(nsptag.clark_name, nsmap=nsmap)


def get_or_add(parent, nsptag_str):
    """
    Return the first direct child element of *parent* with tag matching
    *nsptag_str*. If no such child is found, a new one is created and
    returned.
    """
    _child = child(parent, nsptag_str)
    if _child is None:
        _child = SubElement(parent, nsptag_str)
    return _child


def serialize_for_reading(element):
    """
    Serialize *element* to human-readable XML suitable for tests. No XML
    declaration.
    """
    xml = etree.tostring(element, encoding='unicode', pretty_print=True)
    return XmlString(xml)


def serialize_part_xml(part_elm):
    # if xsi parameter is not set to False, PowerPoint won't load without a
    # repair step; deannotate removes some original xsi:type tags in core.xml
    # if this parameter is left out (or set to True)
    objectify.deannotate(part_elm, xsi=False, cleanup_namespaces=False)
    xml = etree.tostring(part_elm, encoding='UTF-8', standalone=True)
    return xml


def SubElement(parent, nsptag_str, **extra):
    """
    Return an lxml element having *nsptag_str*, newly added as a direct child
    of *parent*. The new element is appended to the sequence of children, so
    this method is not suitable if the child element must be inserted at a
    different position in the sequence. The class of the returned element is
    the custom element class for its tag, if one is defined. Additional
    named parameters defined on lxml ``makeelement()`` are accepted, such as
    attrib=attr_dct and e.g. ``visible='1'``.
    """
    nsptag = NamespacePrefixedTag(nsptag_str)
    return objectify.SubElement(
        parent, nsptag.clark_name, nsmap=nsptag.nsmap, **extra
    )


class XmlString(unicode):
    """
    Provides string comparison override suitable for serialized XML that is
    useful for tests.
    """

    # '    <w:xyz xmlns:a="http://ns/decl/a" attr_name="val">text</w:xyz>'
    # |          |                                          ||           |
    # +----------+------------------------------------------++-----------+
    #  front      attrs                                     | text
    #                                                     close

    _xml_elm_line_patt = re.compile(
        '( *</?[\w:]+)(.*?)(/?>)([^<]*</\w+>)?'
    )

    def __eq__(self, other):
        lines = self.splitlines()
        lines_other = other.splitlines()
        if len(lines) != len(lines_other):
            return False
        for line, line_other in zip(lines, lines_other):
            if not self._eq_elm_strs(line, line_other):
                return False
        return True

    def __ne__(self, other):
        return not self.__eq__(other)

    def _attr_seq(self, attrs):
        """
        Return a sequence of attribute strings parsed from *attrs*. Each
        attribute string is stripped of whitespace on both ends.
        """
        attrs = attrs.strip()
        attr_lst = attrs.split()
        return sorted(attr_lst)

    def _eq_elm_strs(self, line, line_2):
        """
        Return True if the element in *line_2* is XML equivalent to the
        element in *line*.
        """
        front, attrs, close, text = self._parse_line(line)
        front_2, attrs_2, close_2, text_2 = self._parse_line(line_2)
        if front != front_2:
            return False
        if self._attr_seq(attrs) != self._attr_seq(attrs_2):
            return False
        if close != close_2:
            return False
        if text != text_2:
            return False
        return True

    def _parse_line(self, line):
        """
        Return front, attrs, close, text 4-tuple result of parsing XML element
        string *line*.
        """
        match = self._xml_elm_line_patt.match(line)
        front, attrs, close, text = [match.group(n) for n in range(1, 5)]
        return front, attrs, close, text


class BaseOxmlElement(objectify.ObjectifiedElement):
    """
    Provides common behavior for oxml element classes
    """
    @classmethod
    def child_tagnames_after(cls, tagname):
        """
        Return a sequence containing the namespace prefixed child tagnames,
        e.g. 'a:prstGeom', that occur after *tagname* in this element.
        """
        return cls.child_tagnames.tagnames_after(tagname)

    def first_child_found_in(self, *tagnames):
        """
        Return the first child found with tag in *tagnames*, or None if
        not found.
        """
        for tagname in tagnames:
            child = self.find(qn(tagname))
            if child is not None:
                return child
        return None

    def insert_element_before(self, elm, *tagnames):
        successor = self.first_child_found_in(*tagnames)
        if successor is not None:
            successor.addprevious(elm)
        else:
            self.append(elm)
        return elm

    def remove_if_present(self, *tagnames):
        """
        Remove all child elements having tagname in *tagnames*.
        """
        for tagname in tagnames:
            element = self.find(qn(tagname))
            if element is not None:
                self.remove(element)

    @property
    def xml(self):
        """
        Return XML string for this element, suitable for testing purposes.
        Pretty printed for readability and without an XML declaration at the
        top.
        """
        return serialize_for_reading(self)


class _Tagname(object):
    """
    A leaf node in a |_ChildTagnames| tree, containing an individual tagname.
    """
    def __init__(self, tagname):
        super(_Tagname, self).__init__()
        self._tagname = tagname

    def __contains__(self, tagname):
        return tagname == self._tagname

    @property
    def tagnames(self):
        """
        A sequence containing the tagname for this instance
        """
        return (self._tagname,)


class ChildTagnames(object):
    """
    Sequenced tree structure of namespace prefixed tagnames occuring in an
    XML element. An element group is represented by a child node of this same
    class. An element name is represented by an instance of _MemberName.
    """
    def __init__(self, children):
        super(ChildTagnames, self).__init__()
        self._children = tuple(children)

    def __contains__(self, tagname):
        """
        Return |True| if *tagname* belongs to this set of tagnames, |False|
        otherwise. Implements ``tagname in member_names`` functionality.
        """
        for child in self._children:
            if tagname in child:
                return True
        return False

    def __iter__(self):
        return iter(self._children)

    @classmethod
    def from_nested_sequence(cls, *nested_sequence):
        """
        Return an instance of this class constructed from a sequence of
        tagnames and tagname sequences representing the child elements and
        element groups of an XML element.
        """
        children = []
        for item in nested_sequence:
            if isinstance(item, basestring):
                member_name = _Tagname(item)
                children.append(member_name)
                continue
            subtree = ChildTagnames.from_nested_sequence(*item)
            children.append(subtree)
        return cls(children)

    @property
    def tagnames(self):
        """
        A sequence containing the tagnames in this subgraph, in depth-first
        order.
        """
        tagname_lists = [child.tagnames for child in self._children]
        tagnames = itertools.chain(*tagname_lists)
        return tuple(tagnames)

    def tagnames_after(self, tagname):
        """
        Return a sequence containing the tagnames in this subtree that occur
        in children that follow the child containing *tagname*.
        """
        if tagname not in self:
            raise ValueError("tagname '%s' not in element member names")
        # pass over child nodes before and within which item occurs
        found = False
        tagnames_after = []
        for subtree in self:
            if found:
                tagnames_after.extend(subtree.tagnames)
                continue
            found = tagname in subtree
        return tuple(tagnames_after)
