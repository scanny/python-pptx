# encoding: utf-8

"""
Relationship-related objects.
"""


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
