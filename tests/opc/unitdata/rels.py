# encoding: utf-8

"""
Test data for relationship-related unit tests.
"""

from __future__ import absolute_import

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.rels import Relationship, RelationshipCollection

from ...parts.unitdata.part import a_part


class RelationshipCollectionBuilder(object):
    """Builder class for test RelationshipCollections"""
    partname_tmpls = {RT.SLIDE_MASTER: '/ppt/slideMasters/slideMaster%d.xml',
                      RT.SLIDE: '/ppt/slides/slide%d.xml'}

    def __init__(self):
        self.relationships = []
        self.next_rel_num = 1
        self.next_partnums = {}
        self.reltype_ordering = None

    def with_ordering(self, *reltypes):
        self.reltype_ordering = tuple(reltypes)
        return self

    def with_tuple_targets(self, count, reltype):
        for i in range(count):
            rId = self._next_rId
            partname = self._next_tuple_partname(reltype)
            target = a_part().with_partname(partname).build()
            rel = Relationship(rId, reltype, target)
            self.relationships.append(rel)
        return self

    def _next_partnum(self, reltype):
        if reltype not in self.next_partnums:
            self.next_partnums[reltype] = 1
        partnum = self.next_partnums[reltype]
        self.next_partnums[reltype] = partnum + 1
        return partnum

    @property
    def _next_rId(self):
        rId = 'rId%d' % self.next_rel_num
        self.next_rel_num += 1
        return rId

    def _next_tuple_partname(self, reltype):
        partname_tmpl = self.partname_tmpls[reltype]
        partnum = self._next_partnum(reltype)
        return partname_tmpl % partnum

    def build(self):
        rels = RelationshipCollection()
        for rel in self.relationships:
            rels._additem(rel)
        if self.reltype_ordering:
            rels._reltype_ordering = self.reltype_ordering
        return rels


def a_rels():
    """
    Return a PartBuilder instance.
    """
    return RelationshipCollectionBuilder()
