# encoding: utf-8

"""
Test data for Part-related tests.
"""

from __future__ import absolute_import

from pptx.parts.part import BasePart


class PartBuilder(object):
    """Builder class for test Parts"""
    def __init__(self):
        self.partname = '/ppt/slides/slide1.xml'

    def with_partname(self, partname):
        self.partname = partname
        return self

    def build(self):
        p = BasePart()
        p.partname = self.partname
        return p


def a_Part():
    """
    Return a PartBuilder instance.
    """
    return PartBuilder()
