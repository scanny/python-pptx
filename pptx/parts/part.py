# encoding: utf-8

"""
Part objects, including BasePart.
"""

from pptx.opc.package import RelationshipCollection
from pptx.opc.packuri import PackURI
from pptx.util import Collection


class BasePart(object):
    """
    Base class for presentation model parts. Provides common code to all parts
    and is the class we instantiate for parts we don't unmarshal or manipulate
    yet.

    .. attribute:: _element

       ElementTree element for XML parts. ``None`` for binary parts.

    .. attribute:: _load_blob

       Contents of part as a byte string extracted from the package file. May
       be set to ``None`` by subclasses that override ._blob after content is
       unmarshaled, to free up memory.

    .. attribute:: _relationships

       |RelationshipCollection| instance containing the relationships for this
       part.

    """
    def __init__(self, partname, content_type, blob=None):
        """
        ... re-document me
        """
        super(BasePart, self).__init__()
        self._partname = partname
        self._content_type_ = content_type
        self._blob = blob

    def _add_relationship(self, reltype, target, rId, external=False):
        """
        Return newly added |_Relationship| instance of *reltype* between this
        part and *target* with key *rId*. Target mode is set to
        ``RTM.EXTERNAL`` if *external* is |True|. If *reltype* and *target*
        match an existing relationship, that relationship is returned rather
        than creating a new one.
        """
        return self._relationships.add_relationship(
            reltype, target, rId, external
        )

    def after_unmarshal(self):
        """
        Entry point for any logic to be executed once the part's
        relationships have all be unmarshaled. Just a ``pass`` for
        |BasePart|, expected to be overridden by subclasses that need the
        entry point.
        """
        pass

    def before_marshal(self):
        """
        Entry point for pre-serialization processing, for example to finalize
        part naming if necessary. May be overridden by subclasses without
        forwarding call to super.
        """
        # don't place any code here, just catch call if not overridden by
        # subclass
        pass

    @property
    def blob(self):
        """
        Return binary value of this part. Intended to be overridden by
        subclasses. Default is to return load blob.
        """
        return self._blob

    @property
    def content_type(self):
        """
        REMOVE ME, SHOULD BE ABLE TO GET FROM opc.package.Part superclass
        """
        return self._content_type_

    @classmethod
    def load(cls, partname, content_type, blob):
        return cls(partname, content_type, blob)

    @property
    def partname(self):
        """
        |PackURI| instance holding partname of this part, e.g.
        '/ppt/slides/slide1.xml'
        """
        return self._partname

    @partname.setter
    def partname(self, partname):
        self._partname = partname

    @property
    def _content_type(self):
        """
        Content type of this part, e.g.
        'application/vnd.openxmlformats-officedocument.theme+xml'.
        """
        return self._content_type_

    @_content_type.setter
    def _content_type(self, content_type):
        self._content_type_ = content_type

    @property
    def _relationships(self):
        """
        |RelationshipCollection| instance holding the relationships for this
        part.
        """
        if not hasattr(self, '_rels'):
            self._rels = RelationshipCollection(self._partname.baseURI)
        return self._rels


class PartCollection(Collection):
    """
    Sequence of parts. Sensitive to partname index when ordering parts added
    via _loadpart(), e.g. ``/ppt/slide/slide2.xml`` appears before
    ``/ppt/slide/slide10.xml`` rather than after it as it does in a
    lexicographical sort.
    """
    def __init__(self):
        super(PartCollection, self).__init__()

    def add_part(self, part):
        """
        Insert a new part into the collection such that list remains sorted
        in logical partname order (e.g. slide10.xml comes after slide9.xml).
        """
        new_partidx = part.partname.idx
        for idx, seq_part in enumerate(self._values):
            partidx = PackURI(seq_part.partname).idx
            if partidx > new_partidx:
                self._values.insert(idx, part)
                return
        self._values.append(part)
