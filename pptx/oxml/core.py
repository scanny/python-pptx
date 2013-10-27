# encoding: utf-8

"""
General purpose functions that raise the abstraction level of interacting with
objectify elements.
"""

from __future__ import absolute_import

from pptx.oxml import child, SubElement


def get_or_add(start_elm, *path_tags):
    """
    Retrieve the element at the end of the branch starting at parent and
    traversing each of *path_tags* in order, creating any elements not found
    along the way. Not a good solution when sequence of added children is
    likely to be a concern.
    """
    parent = start_elm
    for tag in path_tags:
        child_ = child(parent, tag)
        if child_ is None:
            child_ = SubElement(parent, tag)
        parent = child_
    return child_
