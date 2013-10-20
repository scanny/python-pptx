# encoding: utf-8

"""
Relationship-related objects.
"""

from pptx.util import Collection, Partname


class _Relationship(object):
    """
    Relationship to a part from a package or part. *rId* must be unique in any
    |_RelationshipCollection| this relationship is added to; use
    :attr:`_RelationshipCollection._next_rId` to get a unique rId.
    """
    def __init__(self, rId, reltype, target):
        super(_Relationship, self).__init__()
        # can't get _num right if rId is non-standard form
        assert rId.startswith('rId'), "rId in non-standard form: '%s'" % rId
        self.__rId = rId
        self.__reltype = reltype
        self.__target = target

    @property
    def _rId(self):
        """
        Relationship id for this relationship. Must be of the form
        ``rId[1-9][0-9]*``.
        """
        return self.__rId

    @property
    def _reltype(self):
        """
        Relationship type URI for this relationship. Corresponds roughly to
        the part type of the target part.
        """
        return self.__reltype

    @property
    def _target(self):
        """
        Target part of this relationship. Relationships are directed, from a
        source and a target. The target is always a part.
        """
        return self.__target

    @property
    def _num(self):
        """
        The numeric portion of the rId of this |_Relationship|, expressed as
        an :class:`int`. For example, :attr:`_num` for a relationship with an
        rId of ``'rId12'`` would be ``12``.
        """
        try:
            num = int(self.__rId[3:])
        except ValueError:
            num = 9999
        return num

    @_rId.setter
    def _rId(self, value):
        self.__rId = value


class _RelationshipCollection(Collection):
    """
    Sequence of relationships maintained in rId order. Maintaining the
    relationships in sorted order makes the .rels files both repeatable and
    more readable, which is very helpful for debugging.
    |_RelationshipCollection| has an attribute *_reltype_ordering* which is a
    sequence (tuple) of reltypes. If *_reltype_ordering* contains one or more
    reltype, the collection is maintained in reltype + partname.idx order and
    relationship ids (rIds) are renumbered to match that sequence and any
    numbering gaps are filled in.
    """
    def __init__(self):
        super(_RelationshipCollection, self).__init__()
        self.__reltype_ordering = ()

    def _additem(self, relationship):
        """
        Insert *relationship* into the appropriate position in this ordered
        collection.
        """
        rIds = [rel._rId for rel in self]
        if relationship._rId in rIds:
            tmpl = "cannot add relationship with duplicate rId '%s'"
            raise ValueError(tmpl % relationship._rId)
        self._values.append(relationship)
        self.__resequence()
        # register as observer of partname changes
        relationship._target.add_observer(self)

    @property
    def _next_rId(self):
        """
        Next available rId in collection, starting from 'rId1' and making use
        of any gaps in numbering, e.g. 'rId2' for rIds ['rId1', 'rId3'].
        """
        tmpl = 'rId%d'
        next_rId_num = 1
        for relationship in self._values:
            if relationship._num > next_rId_num:
                return tmpl % next_rId_num
            next_rId_num += 1
        return tmpl % next_rId_num

    def related_part(self, reltype):
        """
        Return first part in collection having relationship type *reltype* or
        raise |KeyError| if not found.
        """
        for relationship in self._values:
            if relationship._reltype == reltype:
                return relationship._target
        tmpl = "no related part with relationship type '%s'"
        raise KeyError(tmpl % reltype)

    @property
    def _reltype_ordering(self):
        """
        Tuple of relationship types, e.g. ``(RT.SLIDE, RT.SLIDE_LAYOUT)``. If
        present, relationships of those types are grouped, and those groups
        are ordered in the same sequence they appear in the tuple. In
        addition, relationships of the same type are sequenced in order of
        partname number (e.g. 16 for /ppt/slides/slide16.xml). If empty, the
        collection is maintained in existing rId order; rIds are not
        renumbered and any gaps in numbering are left to remain. Specifying
        this value for a collection with members causes it to be immediately
        reordered. The ordering is maintained as relationships are added or
        removed, renumbering rIds whenever necessary to also maintain the
        sequence in rId order.
        """
        return self.__reltype_ordering

    @_reltype_ordering.setter
    def _reltype_ordering(self, ordering):
        self.__reltype_ordering = tuple(ordering)
        self.__resequence()

    def rels_of_reltype(self, reltype):
        """
        Return a :class:`list` containing the subset of relationships in this
        collection of type *reltype*. The returned list is ordered by rId.
        Returns an empty list if there are no relationships of type *reltype*
        in the collection.
        """
        return [rel for rel in self._values if rel._reltype == reltype]

    def notify(self, subject, name, value):
        """RelationshipCollection implements the Observer interface"""
        if name == 'partname':
            self.__resequence()

    def __resequence(self):
        """
        Sort relationships and renumber if necessary to maintain values in rId
        order.
        """
        if self.__reltype_ordering:
            def reltype_key(rel):
                reltype = rel._reltype
                if reltype in self.__reltype_ordering:
                    return self.__reltype_ordering.index(reltype)
                return len(self.__reltype_ordering)

            def partname_idx_key(rel):
                partname = Partname(rel._target.partname)
                if partname.idx is None:
                    return 0
                return partname.idx
            self._values.sort(key=lambda rel: partname_idx_key(rel))
            self._values.sort(key=lambda rel: reltype_key(rel))
            # renumber consistent with new sort order
            for idx, relationship in enumerate(self._values):
                relationship._rId = 'rId%d' % (idx+1)
        else:
            self._values.sort(key=lambda rel: rel._num)
