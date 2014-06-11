# encoding: utf-8

"""
General purpose functions that raise the abstraction level of interacting with
lxml elements.
"""

from __future__ import absolute_import

from lxml import etree


def serialize_part_xml(part_elm):
    xml = etree.tostring(part_elm, encoding='UTF-8', standalone=True)
    return xml


# class _Tagname(object):
#     """
#     A leaf node in a |_ChildTagnames| tree, containing an individual tagname.
#     """
#     def __init__(self, tagname):
#         super(_Tagname, self).__init__()
#         self._tagname = tagname

#     def __contains__(self, tagname):
#         return tagname == self._tagname

#     @property
#     def tagnames(self):
#         """
#         A sequence containing the tagname for this instance
#         """
#         return (self._tagname,)


# import itertools
# class ChildTagnames(object):
#     """
#     Sequenced tree structure of namespace prefixed tagnames occuring in an
#     XML element. An element group is represented by a child node of this same
#     class. An element name is represented by an instance of _MemberName.
#     """
#     def __init__(self, children):
#         super(ChildTagnames, self).__init__()
#         self._children = tuple(children)

#     def __contains__(self, tagname):
#         """
#         Return |True| if *tagname* belongs to this set of tagnames, |False|
#         otherwise. Implements ``tagname in member_names`` functionality.
#         """
#         for child in self._children:
#             if tagname in child:
#                 return True
#         return False

#     def __iter__(self):
#         return iter(self._children)

#     @classmethod
#     def from_nested_sequence(cls, *nested_sequence):
#         """
#         Return an instance of this class constructed from a sequence of
#         tagnames and tagname sequences representing the child elements and
#         element groups of an XML element.
#         """
#         children = []
#         for item in nested_sequence:
#             if isinstance(item, basestring):
#                 member_name = _Tagname(item)
#                 children.append(member_name)
#                 continue
#             subtree = ChildTagnames.from_nested_sequence(*item)
#             children.append(subtree)
#         return cls(children)

#     @property
#     def tagnames(self):
#         """
#         A sequence containing the tagnames in this subgraph, in depth-first
#         order.
#         """
#         tagname_lists = [child.tagnames for child in self._children]
#         tagnames = itertools.chain(*tagname_lists)
#         return tuple(tagnames)

#     def tagnames_after(self, tagname):
#         """
#         Return a sequence containing the tagnames in this subtree that occur
#         in children that follow the child containing *tagname*.
#         """
#         if tagname not in self:
#             raise ValueError("tagname '%s' not in element member names")
#         # pass over child nodes before and within which item occurs
#         found = False
#         tagnames_after = []
#         for subtree in self:
#             if found:
#                 tagnames_after.extend(subtree.tagnames)
#                 continue
#             found = tagname in subtree
#         return tuple(tagnames_after)
