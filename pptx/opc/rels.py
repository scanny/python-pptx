# encoding: utf-8

"""
Relationship-related objects.
"""


class Relationship(object):
    """
    Relationship to a part from a package or part. *rId* must be unique in any
    |RelationshipCollection| this relationship is added to; use
    :attr:`RelationshipCollection.next_rId` to get a unique rId.
    """
    def __init__(self, rId, reltype, target):
        super(Relationship, self).__init__()
        self._rId = rId
        self._reltype = reltype
        self._target = target

    @property
    def rId(self):
        """
        Relationship id for this relationship. Must be of the form
        ``rId[1-9][0-9]*``.
        """
        return self._rId

    @property
    def reltype(self):
        """
        Relationship type URI for this relationship. Corresponds roughly to
        the part type of the target part.
        """
        return self._reltype

    @property
    def target(self):
        """
        Target part of this relationship. Relationships are directed, from a
        source and a target. The target is always a part.
        """
        return self._target

    @rId.setter
    def rId(self, value):
        self._rId = value


class RelationshipCollection(object):
    """
    Sequence of |Relationship| instances.
    """
    def __init__(self):
        super(RelationshipCollection, self).__init__()
        self._rels = []

    def __getitem__(self, key):
        """Provides indexed access, (e.g. 'collection[0]')."""
        return self._rels.__getitem__(key)

    def __iter__(self):
        """Supports iteration (e.g. 'for x in collection:')."""
        return self._rels.__iter__()

    def __len__(self):
        """Supports len() function (e.g. 'len(collection) == 1')."""
        return len(self._rels)

    def add_rel(self, relationship):
        """
        Insert *relationship* into the appropriate position in this ordered
        collection.
        """
        if relationship.rId in self._rIds:
            tmpl = "cannot add relationship with duplicate rId '%s'"
            raise ValueError(tmpl % relationship.rId)
        self._rels.append(relationship)

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

    def part_with_rId(self, rId):
        """
        Return target part with matching *rId*, raising |KeyError| if not
        found.
        """
        for rel in self:
            if rel.rId == rId:
                return rel.target
        raise KeyError("no relationship with rId '%s'" % rId)

    def related_part(self, reltype):
        """
        Return first part in collection having relationship type *reltype* or
        raise |KeyError| if not found.
        """
        for rel in self:
            if rel.reltype == reltype:
                return rel.target
        tmpl = "no related part with relationship type '%s'"
        raise KeyError(tmpl % reltype)

    def rels_of_reltype(self, reltype):
        """
        Return a :class:`list` containing the subset of relationships in this
        collection of type *reltype*. The returned list is ordered by rId.
        Returns an empty list if there are no relationships of type *reltype*
        in the collection.
        """
        return [rel for rel in self if rel.reltype == reltype]

    @property
    def _rIds(self):
        return [rel.rId for rel in self]
