# encoding: utf-8

"""Test suite for pptx.part module."""

from __future__ import absolute_import

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.rels import Relationship, RelationshipCollection
from pptx.parts.slides import SlideMaster

from ..unitutil import instance_mock


class DescribeRelationship(object):

    def it_remembers_attrs_from_construction(self, rel, slide_master):
        assert rel.rId == 'rId2'
        assert rel.reltype == RT.SLIDE_MASTER
        assert rel.target == slide_master

    def it_allows_its_rId_to_be_changed(self, rel):
        assert rel.rId == 'rId2'
        rel.rId = 'rId9'
        assert rel.rId == 'rId9'


class DescribeRelationshipCollection(object):

    def it_can_find_a_related_part_by_reltype(self, rels):
        part = rels.related_part(RT.SLIDE_MASTER)
        assert isinstance(part, SlideMaster)

    def test_it_raises_on_related_part_not_found(self, rels):
        with pytest.raises(KeyError):
            rels.related_part('foobar')

    def it_raises_on_add_rel_with_duplicate_rId(self, rels, rel):
        with pytest.raises(ValueError):
            rels.add_rel(rel)

    def it_knows_which_rels_match_a_specified_reltype(self, rels):
        rels_to_slides = rels.rels_of_reltype(RT.SLIDE)
        assert [r.rId for r in rels_to_slides] == ['rId1', 'rId3']

    def it_fills_first_rId_gap_when_adding_rel(self, rels_with_rId_gap):
        rels, expected_next_rId = rels_with_rId_gap
        next_rId = rels._next_rId
        assert next_rId == expected_next_rId

    # fixtures ---------------------------------------------

    @pytest.fixture
    def rels(self, slide_master):
        """
        General-purpose RelationshipCollection fixture
        """
        rels = RelationshipCollection()
        rels.add_rel(Relationship('rId1', RT.SLIDE, None))
        rels.add_rel(Relationship('rId2', RT.SLIDE_MASTER, slide_master))
        rels.add_rel(Relationship('rId3', RT.SLIDE, None))
        return rels

    @pytest.fixture(
        params=[
            (('rId2', 'rId3', 'rId4'), 'rId1'),
            (('rId1', 'rId3', 'rId4'), 'rId2'),
            (('rId1', 'rId2', 'rId4'), 'rId3'),
            (('rId1', 'rId2', 'rId3'), 'rId4'),
        ]
    )
    def rels_with_rId_gap(self, request):
        """
        Return RelationshipCollection with a set of rels having a gap in the
        rId sequence. Return value is 2-tuple (rels, expected_next_rId).
        """
        rels = RelationshipCollection()
        rIds, expected_next_rId = request.param
        rels.add_rel(Relationship(rIds[0], None, None))
        rels.add_rel(Relationship(rIds[1], None, None))
        rels.add_rel(Relationship(rIds[2], None, None))
        return (rels, expected_next_rId)


# ===========================================================================
# fixtures
# ===========================================================================

@pytest.fixture
def rel(slide_master):
    return Relationship('rId2', RT.SLIDE_MASTER, slide_master)


@pytest.fixture
def slide_master(request):
    return instance_mock(request, SlideMaster)
