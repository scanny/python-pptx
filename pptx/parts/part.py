# encoding: utf-8

"""
Part objects, including BasePart.
"""

from pptx.opc.rels import Relationship, RelationshipCollection
from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import serialize_part_xml
from pptx.util import Collection, Partname


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
    def __init__(self, content_type=None, partname=None):
        """
        Needs content_type parameter so newly created parts (not loaded from
        package) can register their content type.
        """
        super(BasePart, self).__init__()
        self._content_type_ = content_type
        self._partname = partname
        self._element = None
        self._load_blob = None
        self._relationships = RelationshipCollection()

    @property
    def _blob(self):
        """
        Default is to return unchanged _load_blob. Dynamic parts will override.
        Raises |ValueError| if _load_blob is None.
        """
        if self.partname.endswith('.xml'):
            assert self._element is not None, (
                'BasePart._blob is undefined for xml parts when part._elemen'
                't is None'
            )
            xml = serialize_part_xml(self._element)
            return xml
        # default for binary parts is to return _load_blob unchanged
        assert self._load_blob, (
            'BasePart._blob called on part with no _load_blob; perhaps _blob'
            ' not overridden by sub-class?'
        )
        return self._load_blob

    @property
    def _content_type(self):
        """
        Content type of this part, e.g.
        'application/vnd.openxmlformats-officedocument.theme+xml'.
        """
        assert self._content_type_, (
            'BasePart._content_type accessed before assigned'
        )
        return self._content_type_

    @_content_type.setter
    def _content_type(self, content_type):
        self._content_type_ = content_type

    @property
    def partname(self):
        """Part name of this part, e.g. '/ppt/slides/slide1.xml'."""
        assert self._partname, "BasePart.partname referenced before assigned"
        return self._partname

    @partname.setter
    def partname(self, partname):
        self._partname = partname

    def _add_relationship(self, reltype, target_part):
        """
        Return new relationship of *reltype* to *target_part* after adding it
        to the relationship collection of this part.
        """
        # reuse existing relationship if there's a match
        for rel in self._relationships:
            if rel._target == target_part and rel._reltype == reltype:
                return rel
        # otherwise construct a new one
        rId = self._relationships._next_rId
        rel = Relationship(rId, reltype, target_part)
        self._relationships._additem(rel)
        return rel

    def _load(self, pkgpart, part_dict):
        """
        Load part and relationships from package part, and propagate load
        process down the relationship graph. *pkgpart* is an instance of
        :class:`pptx.packaging.Part` containing the part contents read from
        the on-disk package. *part_dict* is a dictionary of already-loaded
        parts, keyed by partname.
        """
        # set attributes from package part
        self._content_type_ = pkgpart.content_type
        self._partname = pkgpart.partname
        if pkgpart.partname.endswith('.xml'):
            self._element = parse_xml_bytes(pkgpart.blob)
        else:
            self._load_blob = pkgpart.blob

        # discard any previously loaded relationships
        self._relationships = RelationshipCollection()

        # load relationships and propagate load for related parts
        for pkgrel in pkgpart.relationships:
            # unpack working values for part to be loaded
            reltype = pkgrel.reltype
            target_pkgpart = pkgrel.target
            partname = target_pkgpart.partname
            content_type = target_pkgpart.content_type

            # create target part
            if partname in part_dict:
                part = part_dict[partname]
            else:
                # !!!~!!!~!!!~!!!~!!!~!!!~!!!~
                # get rid of this as soon as possible
                from pptx.presentation import Part
                # !!!~!!!~!!!~!!!~!!!~!!!~!!!~
                part = Part(reltype, content_type)
                part_dict[partname] = part
                part._load(target_pkgpart, part_dict)

            # create model-side package relationship
            model_rel = Relationship(pkgrel.rId, reltype, part)
            self._relationships._additem(model_rel)
        return self


class PartCollection(Collection):
    """
    Sequence of parts. Sensitive to partname index when ordering parts added
    via _loadpart(), e.g. ``/ppt/slide/slide2.xml`` appears before
    ``/ppt/slide/slide10.xml`` rather than after it as it does in a
    lexicographical sort.
    """
    def __init__(self):
        super(PartCollection, self).__init__()

    def _loadpart(self, part):
        """
        Insert a new part loaded from a package, such that list remains
        sorted in logical partname order (e.g. slide10.xml comes after
        slide9.xml).
        """
        new_partidx = Partname(part.partname).idx
        for idx, seq_part in enumerate(self._values):
            partidx = Partname(seq_part.partname).idx
            if partidx > new_partidx:
                self._values.insert(idx, part)
                return
        self._values.append(part)
