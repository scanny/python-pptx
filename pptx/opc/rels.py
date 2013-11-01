# encoding: utf-8

"""
Relationship-related objects.
"""

from pptx.util import Collection


class Relationship(object):
    """
    Relationship to a part from a package or part. *rId* must be unique in any
    |RelationshipCollection| this relationship is added to; use
    :attr:`RelationshipCollection._next_rId` to get a unique rId.
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


class RelationshipCollection(Collection):
    """
    Sequence of |Relationship| instances.
    """
    def __init__(self):
        super(RelationshipCollection, self).__init__()

    def add_rel(self, relationship):
        """
        Insert *relationship* into the appropriate position in this ordered
        collection.
        """
        rIds = [rel.rId for rel in self]
        if relationship.rId in rIds:
            tmpl = "cannot add relationship with duplicate rId '%s'"
            raise ValueError(tmpl % relationship.rId)
        self._values.append(relationship)

    @property
    def _next_rId(self):
        """
        Next available rId in collection, starting from 'rId1' and making use
        of any gaps in numbering, e.g. 'rId2' for rIds ['rId1', 'rId3'].
        """
        rIds = [rel.rId for rel in self]
        tmpl = 'rId%d'
        for n in range(1, 999999):
            rId_candidate = tmpl % n  # like 'rId19'
            if rId_candidate not in rIds:
                return rId_candidate
        raise ValueError('implausible relationship count in collection')

    def related_part(self, reltype):
        """
        Return first part in collection having relationship type *reltype* or
        raise |KeyError| if not found.
        """
        for relationship in self._values:
            if relationship.reltype == reltype:
                return relationship.target
        tmpl = "no related part with relationship type '%s'"
        raise KeyError(tmpl % reltype)

    def rels_of_reltype(self, reltype):
        """
        Return a :class:`list` containing the subset of relationships in this
        collection of type *reltype*. The returned list is ordered by rId.
        Returns an empty list if there are no relationships of type *reltype*
        in the collection.
        """
        return [rel for rel in self._values if rel.reltype == reltype]
