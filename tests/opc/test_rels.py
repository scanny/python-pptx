# encoding: utf-8

"""Test suite for pptx.part module."""

from __future__ import absolute_import

import pytest

from mock import call, Mock, patch, PropertyMock

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.oxml import CT_Relationships
from pptx.opc.packuri import PackURI
from pptx.opc.rels import _Relationship, RelationshipCollection

from ..unitutil import class_mock, loose_mock


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

    def it_supports_indexed_access(self, rels):
        try:
            rels[0]
        except TypeError:
            msg = 'RelationshipCollection does not support indexed access'
            pytest.fail(msg)
        except IndexError:
            pass

    def it_also_has_dict_style_get_rel_by_rId(self, rels_with_known_rel):
        rels, rId, known_rel = rels_with_known_rel
        assert rels[rId] == known_rel

    def it_should_raise_on_failed_lookup_by_rId(self, rels):
        with pytest.raises(KeyError):
            rels['rId666']

    def it_has_a_len(self, rels):
        assert len(rels) == 0

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

    def it_can_find_a_related_part_by_reltype(
            self, rels_with_target_known_by_reltype):
        rels, reltype, known_target_part = rels_with_target_known_by_reltype
        part = rels.part_with_reltype(reltype)
        assert part is known_target_part

    def it_can_find_a_related_part_by_rId(self, rels_with_known_target_part):
        rels, rId, known_target_part = rels_with_known_target_part
        part = rels.part_with_rId(rId)
        assert part is known_target_part

    def it_raises_KeyError_on_part_with_rId_not_found(self, rels):
        with pytest.raises(KeyError):
            rels.part_with_rId('rId666')

    # def it_knows_the_next_available_rId(self, rels_with_rId_gap):
    #     rels, expected_next_rId = rels_with_rId_gap
    #     next_rId = rels.next_rId
    #     assert next_rId == expected_next_rId

#     def it_raises_on_no_next_rId_found(self, _rIds):
#         rels = RelationshipCollection()
#         with pytest.raises(AssertionError):
#             rels.next_rId

#     def it_raises_on_add_rel_with_duplicate_rId(self, rels, rel):
#         with pytest.raises(ValueError):
#             rels.add_rel(rel)

#     def it_knows_which_rels_match_a_specified_reltype(self, rels):
#         rels_to_slides = rels.rels_of_reltype(RT.SLIDE)
#         assert [r.rId for r in rels_to_slides] == ['rId1', 'rId3']

    def it_can_compose_rels_xml(self, rels_with_known_rels, rels_elm):
        # exercise ---------------------
        rels_with_known_rels.xml
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
    def rel(self, _rId, _reltype, _target_part, _baseURI):
        return _Relationship(_rId, _reltype, _target_part, _baseURI)

    @pytest.fixture
    def rels(self, _baseURI):
        return RelationshipCollection(_baseURI)

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

    @pytest.fixture
    def rels_with_known_rel(self, rels, _rId, rel):
        rels._rels.append(rel)
        return rels, _rId, rel

    @pytest.fixture
    def rels_with_known_rels(self):
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
    def rels_with_target_known_by_reltype(
            self, rels, _rel_with_target_known_by_reltype):
        rel, reltype, target_part = _rel_with_target_known_by_reltype
        rels._rels.append(rel)
        return rels, reltype, target_part

    @pytest.fixture
    def rels_with_known_target_part(self, rels, _rel_with_known_target_part):
        rel, rId, target_part = _rel_with_known_target_part
        rels._rels.append(rel)
        return rels, rId, target_part

    @pytest.fixture
    def _baseURI(self):
        return '/baseURI'

    @pytest.fixture
    def _rel_with_known_target_part(
            self, _rId, _reltype, _target_part, _baseURI):
        rel = _Relationship(_rId, _reltype, _target_part, _baseURI)
        return rel, _rId, _target_part

    @pytest.fixture
    def _rel_with_target_known_by_reltype(
            self, _rId, _reltype, _target_part, _baseURI):
        rel = _Relationship(_rId, _reltype, _target_part, _baseURI)
        return rel, _reltype, _target_part

    @pytest.fixture
    def _reltype(self):
        return RT.SLIDE

    @pytest.fixture
    def _rId(self):
        return 'rId6'

    @pytest.fixture
    def _target_part(self, request):
        return loose_mock(request)
