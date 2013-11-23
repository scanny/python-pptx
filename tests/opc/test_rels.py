# encoding: utf-8

"""Test suite for pptx.part module."""

from __future__ import absolute_import

import pytest

from mock import call, Mock, patch, PropertyMock

from pptx.opc.oxml import CT_Relationships
from pptx.opc.packuri import PackURI
from pptx.opc.rels import _Relationship, RelationshipCollection

from ..unitutil import class_mock


class Describe_Relationship(object):

    def it_remembers_construction_values(self):
        # test data --------------------
        rId = 'rId9'
        reltype = 'reltype'
        target = Mock(name='target_part')
        external = False
        # exercise ---------------------
        rel = _Relationship(rId, reltype, target, None, external)
        # verify -----------------------
        assert rel.rId == rId
        assert rel.reltype == reltype
        assert rel.target_part == target
        assert rel.is_external == external

    def it_should_raise_on_target_part_access_on_external_rel(self):
        rel = _Relationship(None, None, None, None, external=True)
        with pytest.raises(ValueError):
            rel.target_part

    def it_should_have_target_ref_for_external_rel(self):
        rel = _Relationship(None, None, 'target', None, external=True)
        assert rel.target_ref == 'target'

    def it_should_have_relative_ref_for_internal_rel(self):
        """
        Internal relationships (TargetMode == 'Internal' in the XML) should
        have a relative ref, e.g. '../slideLayouts/slideLayout1.xml', for
        the target_ref attribute.
        """
        part = Mock(name='part', partname=PackURI('/ppt/media/image1.png'))
        baseURI = '/ppt/slides'
        rel = _Relationship(None, None, part, baseURI)  # external=False
        assert rel.target_ref == '../media/image1.png'


class DescribeRelationshipCollection(object):

    def it_has_a_len(self):
        rels = RelationshipCollection(None)
        assert len(rels) == 0

    def it_supports_indexed_access(self):
        rels = RelationshipCollection(None)
        try:
            rels[0]
        except TypeError:
            msg = 'RelationshipCollection does not support indexed access'
            pytest.fail(msg)
        except IndexError:
            pass

    def it_has_dict_style_lookup_of_rel_by_rId(self):
        rel = Mock(name='rel', rId='foobar')
        rels = RelationshipCollection(None)
        rels._rels.append(rel)
        assert rels['foobar'] == rel

    def it_should_raise_on_failed_lookup_by_rId(self):
        rel = Mock(name='rel', rId='foobar')
        rels = RelationshipCollection(None)
        rels._rels.append(rel)
        with pytest.raises(KeyError):
            rels['barfoo']

    def it_can_add_a_relationship(self, _Relationship_):
        baseURI, rId, reltype, target, is_external = (
            'baseURI', 'rId9', 'reltype', 'target', False
        )
        rels = RelationshipCollection(baseURI)
        rel = rels.add_relationship(reltype, target, rId, is_external)
        _Relationship_.assert_called_once_with(rId, reltype, target, baseURI,
                                               is_external)
        assert rels[0] == rel
        assert rel == _Relationship_.return_value

#     def it_raises_on_add_rel_with_duplicate_rId(self, rels, rel):
#         with pytest.raises(ValueError):
#             rels.add_rel(rel)

#     def it_fills_first_rId_gap_when_adding_rel(self, rels_with_rId_gap):
#         rels, expected_next_rId = rels_with_rId_gap
#         next_rId = rels.next_rId
#         assert next_rId == expected_next_rId

#     def it_raises_on_no_next_rId_found(self, _rIds):
#         rels = RelationshipCollection()
#         with pytest.raises(AssertionError):
#             rels.next_rId

#     def it_can_find_a_related_part_by_rId(self, rels, slide_master):
#         part = rels.part_with_rId('rId2')
#         assert part is slide_master

#     def it_raises_KeyError_on_part_with_rId_not_found(self, rels):
#         with pytest.raises(KeyError):
#             rels.part_with_rId('rId666')

#     def it_can_find_a_related_part_by_reltype(self, rels):
#         part = rels.related_part(RT.SLIDE_MASTER)
#         assert isinstance(part, SlideMaster)

#     def it_raises_on_related_part_not_found(self, rels):
#         with pytest.raises(KeyError):
#             rels.related_part('foobar')

#     def it_knows_which_rels_match_a_specified_reltype(self, rels):
#         rels_to_slides = rels.rels_of_reltype(RT.SLIDE)
#         assert [r.rId for r in rels_to_slides] == ['rId1', 'rId3']

    def it_can_compose_rels_xml(self, rels, rels_elm):
        # exercise ---------------------
        rels.xml
        # trace ------------------------
        print('Actual calls:\n%s' % rels_elm.mock_calls)
        # verify -----------------------
        expected_rels_elm_calls = [
            call.add_rel('rId1', 'http://rt-hyperlink', 'http://some/link',
                         True),
            call.add_rel('rId2', 'http://rt-image', '../media/image1.png',
                         False),
            call.xml()
        ]
        assert rels_elm.mock_calls == expected_rels_elm_calls

    # fixtures ---------------------------------------------

    @pytest.fixture
    def _Relationship_(self, request):
        return class_mock(request, 'pptx.opc.rels._Relationship')

    @pytest.fixture
    def rels(self):
        """
        Populated RelationshipCollection instance that will exercise the
        rels.xml property.
        """
        rels = RelationshipCollection('/baseURI')
        rels.add_relationship(
            reltype='http://rt-hyperlink', target='http://some/link',
            rId='rId1', is_external=True
        )
        part = Mock(name='part')
        part.partname.relative_ref.return_value = '../media/image1.png'
        rels.add_relationship(reltype='http://rt-image', target=part,
                              rId='rId2')
        return rels

    @pytest.fixture
    def rels_elm(self, request):
        """
        Return a rels_elm mock that will be returned from
        CT_Relationships.new()
        """
        # create rels_elm mock with a .xml property
        rels_elm = Mock(name='rels_elm')
        xml = PropertyMock(name='xml')
        type(rels_elm).xml = xml
        rels_elm.attach_mock(xml, 'xml')
        rels_elm.reset_mock()  # to clear attach_mock call
        # patch CT_Relationships to return that rels_elm
        patch_ = patch.object(CT_Relationships, 'new', return_value=rels_elm)
        patch_.start()
        request.addfinalizer(patch_.stop)
        return rels_elm

#     @pytest.fixture
#     def _rIds(self, request):
#         mock_rId_lst = Mock(name='mock_rId_lst')
#         mock_rId_lst.__contains__ = Mock(return_value=True)
#         _rIds = property_mock(
#             request, 'pptx.opc.rels.RelationshipCollection._rIds'
#         )
#         _rIds.return_value = mock_rId_lst
#         return _rIds

#     @pytest.fixture
#     def rels(self, slide_master):
#         """
#         General-purpose RelationshipCollection fixture
#         """
#         rels = RelationshipCollection()
#         rels.add_rel(Relationship('rId1', RT.SLIDE, None))
#         rels.add_rel(Relationship('rId2', RT.SLIDE_MASTER, slide_master))
#         rels.add_rel(Relationship('rId3', RT.SLIDE, None))
#         return rels

#     @pytest.fixture(
#         params=[
#             (('rId2', 'rId3', 'rId4'), 'rId1'),
#             (('rId1', 'rId3', 'rId4'), 'rId2'),
#             (('rId1', 'rId2', 'rId4'), 'rId3'),
#             (('rId1', 'rId2', 'rId3'), 'rId4'),
#         ]
#     )

#     def rels_with_rId_gap(self, request):
#         """
#         Return RelationshipCollection with a set of rels having a gap in the
#         rId sequence. Return value is 2-tuple (rels, expected_next_rId).
#         """
#         rels = RelationshipCollection()
#         rIds, expected_next_rId = request.param
#         rels.add_rel(Relationship(rIds[0], None, None))
#         rels.add_rel(Relationship(rIds[1], None, None))
#         rels.add_rel(Relationship(rIds[2], None, None))
#         return (rels, expected_next_rId)

#     @pytest.fixture
#     def rel(slide_master):
#         return Relationship('rId2', RT.SLIDE_MASTER, slide_master)

#     @pytest.fixture
#     def slide_master(request):
#         return instance_mock(request, SlideMaster)
