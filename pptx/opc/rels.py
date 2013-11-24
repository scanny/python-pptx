# encoding: utf-8

"""
Relationship-related objects.
"""

from __future__ import absolute_import

from .oxml import CT_Relationships


class _Relationship(object):
    """
    Value object for relationship to part.
    """
    def __init__(self, rId, reltype, target, baseURI, external=False):
        super(_Relationship, self).__init__()
        self._rId = rId
        self._reltype = reltype
        self._target = target
        self._baseURI = baseURI
        self._is_external = bool(external)

    @property
    def is_external(self):
        return self._is_external

    @property
    def reltype(self):
        return self._reltype

    @property
    def rId(self):
        return self._rId

    @property
    def target_part(self):
        if self._is_external:
            raise ValueError(
                "target_part property on _Relationship is undefined when tar"
                "get mode is External"
            )
        return self._target

    @property
    def target_ref(self):
        if self._is_external:
            return self._target
        else:
            return self._target.partname.relative_ref(self._baseURI)


class RelationshipCollection(object):
    """
    Collection object for |_Relationship| instances, having list semantics.
    """
    def __init__(self, baseURI):
        super(RelationshipCollection, self).__init__()
        self._baseURI = baseURI
        self._rels = []

    def __getitem__(self, key):
        """
        Implements access by subscript, e.g. ``rels[9]``. It also implements
        dict-style lookup of a relationship by rId, e.g. ``rels['rId1']``.
        """
        if isinstance(key, basestring):
            for rel in self._rels:
                if rel.rId == key:
                    return rel
            raise KeyError("no rId '%s' in RelationshipCollection" % key)
        else:
            return self._rels.__getitem__(key)

    def __iter__(self):
        """
        Supports iteration (e.g. 'for rel in rels:')
        """
        return self._rels.__iter__()

    def __len__(self):
        """Implements len() built-in on this object"""
        return self._rels.__len__()

    def add_relationship(self, reltype, target, rId, is_external=False):
        """
        Return a |_Relationship| instance of *reltype* to *target_part*, an
        existing rel if a matching one is present, otherwise a newly created
        one.
        """
        # if relationship.rId in self._rIds:
        #     tmpl = "cannot add relationship with duplicate rId '%s'"
        #     raise ValueError(tmpl % relationship.rId)
        rel = _Relationship(rId, reltype, target, self._baseURI, is_external)
        self._rels.append(rel)
        return rel

    def get_or_add(self, reltype, target_part):
        """
        Return relationship of *reltype* to *target_part*, newly added if not
        already present in collection.
        """
        rel = self._get_matching(reltype, target_part)
        if rel is None:
            rId = self.next_rId
            rel = self.add_relationship(reltype, target_part, rId)
        return rel

    def get_rel_of_type(self, reltype):
        """
        Return single relationship of type *reltype* from the collection.
        Raises |KeyError| if no matching relationship is found. Raises
        |ValueError| if more than one matching relationship is found.
        """
        matching = [rel for rel in self._rels if rel.reltype == reltype]
        if len(matching) == 0:
            tmpl = "no relationship of type '%s' in collection"
            raise KeyError(tmpl % reltype)
        if len(matching) > 1:
            tmpl = "multiple relationships of type '%s' in collection"
            raise ValueError(tmpl % reltype)
        return matching[0]

#     def get_rels_of_type(self, reltype):
#         """
#         Return a :class:`list` containing the subset of relationships in this
#         collection of type *reltype*. The returned list is ordered by rId.
#         Returns an empty list if there are no relationships of type *reltype*
#         in the collection.
#         """
#         return [rel for rel in self if rel.reltype == reltype]

    @property
    def next_rId(self):
        """
        Next available rId in collection, starting from 'rId1' and making use
        of any gaps in numbering, e.g. 'rId2' for rIds ['rId1', 'rId3'].
        """
        for n in range(1, len(self)+2):
            rId_candidate = 'rId%d' % n  # like 'rId19'
            if rId_candidate not in self._rIds:
                return rId_candidate
        assert False, 'programming error in RelationshipCollection.next_rId'

    def part_with_reltype(self, reltype):
        """
        Return target part of rel with matching *reltype*, raising |KeyError|
        if not found and |ValueError| if more than one matching relationship
        is found.
        """
        rel = self.get_rel_of_type(reltype)
        return rel.target_part

    def part_with_rId(self, rId):
        """
        Return target part with matching *rId*, raising |KeyError| if not
        found.
        """
        for rel in self:
            if rel.rId == rId:
                return rel.target_part
        raise KeyError("no relationship with rId '%s'" % rId)

    @property
    def xml(self):
        """
        Serialize this relationship collection into XML suitable for storage
        as a .rels file in an OPC package.
        """
        rels_elm = CT_Relationships.new()
        for rel in self._rels:
            rels_elm.add_rel(rel.rId, rel.reltype, rel.target_ref,
                             rel.is_external)
        return rels_elm.xml

    def _get_matching(self, reltype, target_part):
        """
        Return relationship of matching *reltype* to *target_part* from
        collection, or none if no such relationship is present.
        """
        for rel in self._rels:
            if rel.target_part == target_part and rel.reltype == reltype:
                return rel
        return None

    @property
    def _rIds(self):
        return [rel.rId for rel in self]
