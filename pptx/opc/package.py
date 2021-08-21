# encoding: utf-8

"""Fundamental Open Packaging Convention (OPC) objects.

The :mod:`pptx.packaging` module coheres around the concerns of reading and writing
presentations to and from a .pptx file.
"""

from pptx.compat import is_string
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.oxml import CT_Relationships, serialize_part_xml
from pptx.opc.packuri import PACKAGE_URI, PackURI
from pptx.opc.serialized import PackageReader, PackageWriter
from pptx.oxml import parse_xml
from pptx.util import lazyproperty


class OpcPackage(object):
    """Main API class for |python-opc|.

    A new instance is constructed by calling the :meth:`open` classmethod with a path
    to a package file or file-like object containing a package (.pptx file).
    """

    def __init__(self):
        super(OpcPackage, self).__init__()

    def after_unmarshal(self):
        """
        Called by loading code after all parts and relationships have been
        loaded, to afford the opportunity for any required post-processing.
        This one does nothing other than catch the call if a subclass
        doesn't.
        """
        pass

    def iter_parts(self):
        """Generate exactly one reference to each part in the package."""

        def walk_parts(source, visited=list()):
            for rel in source.rels.values():
                if rel.is_external:
                    continue
                part = rel.target_part
                if part in visited:
                    continue
                visited.append(part)
                yield part
                new_source = part
                for part in walk_parts(new_source, visited):
                    yield part

        for part in walk_parts(self):
            yield part

    def iter_rels(self):
        """Generate exactly one reference to each relationship in package.

        Performs a depth-first traversal of the rels graph.
        """

        def walk_rels(source, visited=None):
            visited = [] if visited is None else visited
            for rel in source.rels.values():
                yield rel
                # --- external items can have no relationships ---
                if rel.is_external:
                    continue
                # --- all relationships other than those for the package belong to a
                # --- part. Once that part has been processed, processing it again
                # --- would lead to the same relationships appearing more than once.
                part = rel.target_part
                if part in visited:
                    continue
                visited.append(part)
                new_source = part
                # --- recurse into relationships of each unvisited target-part ---
                for rel in walk_rels(new_source, visited):
                    yield rel

        for rel in walk_rels(self):
            yield rel

    def load_rel(self, reltype, target, rId, is_external=False):
        """
        Return newly added |_Relationship| instance of *reltype* between this
        part and *target* with key *rId*. Target mode is set to
        ``RTM.EXTERNAL`` if *is_external* is |True|. Intended for use during
        load from a serialized package, where the rId is well known. Other
        methods exist for adding a new relationship to the package during
        processing.
        """
        return self.rels.add_relationship(reltype, target, rId, is_external)

    @property
    def main_document_part(self):
        """Return |Part| subtype serving as the main document part for this package.

        In this case it will be a |Presentation| part.
        """
        return self.part_related_by(RT.OFFICE_DOCUMENT)

    def next_partname(self, tmpl):
        """Return |PackURI| next available partname matching `tmpl`.

        `tmpl` is a printf (%)-style template string containing a single replacement
        item, a '%d' to be used to insert the integer portion of the partname.
        Example: '/ppt/slides/slide%d.xml'
        """
        partnames = [part.partname for part in self.iter_parts()]
        for n in range(1, len(partnames) + 2):
            candidate_partname = tmpl % n
            if candidate_partname not in partnames:
                return PackURI(candidate_partname)
        raise Exception("ProgrammingError: ran out of candidate_partnames")

    @classmethod
    def open(cls, pkg_file):
        """Return an |OpcPackage| instance loaded with the contents of `pkg_file`."""
        pkg_reader = PackageReader.from_file(pkg_file)
        package = cls()
        Unmarshaller.unmarshal(pkg_reader, package, PartFactory)
        return package

    def part_related_by(self, reltype):
        """Return (single) part having relationship to this package of `reltype`.

        Raises |KeyError| if no such relationship is found and |ValueError| if more than
        one such relationship is found.
        """
        return self.rels.part_with_reltype(reltype)

    @property
    def parts(self):
        """
        Return a list containing a reference to each of the parts in this
        package.
        """
        return [part for part in self.iter_parts()]

    def relate_to(self, part, reltype):
        """Return rId key of relationship of `reltype` to `target`.

        If such a relationship already exists, its rId is returned. Otherwise the
        relationship is added and its new rId returned.
        """
        rel = self.rels.get_or_add(reltype, part)
        return rel.rId

    @lazyproperty
    def rels(self):
        """
        Return a reference to the |RelationshipCollection| holding the
        relationships for this package.
        """
        return RelationshipCollection(PACKAGE_URI.baseURI)

    def save(self, pkg_file):
        """Save this package to `pkg_file`.

        `pkg_file` can be either a path to a file (a string) or a file-like object.
        """
        for part in self.parts:
            part.before_marshal()
        PackageWriter.write(pkg_file, self.rels, self.parts)


class Part(object):
    """Base class for package parts.

    Provides common properties and methods, but intended to be subclassed in client code
    to implement specific part behaviors. Also serves as the default class for parts
    that are not yet given specific behaviors.
    """

    def __init__(self, partname, content_type, blob=None, package=None):
        super(Part, self).__init__()
        self._partname = partname
        self._content_type = content_type
        self._blob = blob
        self._package = package

    # load/save interface to OpcPackage ------------------------------

    def after_unmarshal(self):
        """
        Entry point for post-unmarshaling processing, for example to parse
        the part XML. May be overridden by subclasses without forwarding call
        to super.
        """
        # don't place any code here, just catch call if not overridden by
        # subclass
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
        """Contents of this package part as a sequence of bytes.

        May be text (XML generally) or binary. Intended to be overridden by subclasses.
        Default behavior is to return the blob initial loaded during `Package.open()`
        operation.
        """
        return self._blob

    @blob.setter
    def blob(self, bytes_):
        """Note that not all subclasses use the part blob as their blob source.

        In particular, the |XmlPart| subclass uses its `self._element` to serialize a
        blob on demand. This works fine for binary parts though.
        """
        self._blob = bytes_

    @property
    def content_type(self):
        """Content-type (MIME-type) of this part."""
        return self._content_type

    @classmethod
    def load(cls, partname, content_type, blob, package):
        """Return `cls` instance loaded from arguments.

        This one is a straight pass-through, but subtypes may do some pre-processing,
        see XmlPart for an example.
        """
        return cls(partname, content_type, blob, package)

    def load_rel(self, reltype, target, rId, is_external=False):
        """
        Return newly added |_Relationship| instance of *reltype* between this
        part and *target* with key *rId*. Target mode is set to
        ``RTM.EXTERNAL`` if *is_external* is |True|. Intended for use during
        load from a serialized package, where the rId is well known. Other
        methods exist for adding a new relationship to a part when
        manipulating a part.
        """
        return self.rels.add_relationship(reltype, target, rId, is_external)

    @property
    def package(self):
        """|OpcPackage| instance this part belongs to."""
        return self._package

    @property
    def partname(self):
        """|PackURI| partname for this part, e.g. "/ppt/slides/slide1.xml"."""
        return self._partname

    @partname.setter
    def partname(self, partname):
        if not isinstance(partname, PackURI):
            tmpl = "partname must be instance of PackURI, got '%s'"
            raise TypeError(tmpl % type(partname).__name__)
        self._partname = partname

    # relationship management interface for child objects ------------

    def drop_rel(self, rId):
        """Remove relationship identified by `rId` if its reference count is under 2.

        Relationships with a reference count of 0 are implicit relationships. Note that
        only XML parts can drop relationships.
        """
        if self._rel_ref_count(rId) < 2:
            del self.rels[rId]

    def part_related_by(self, reltype):
        """Return (single) part having relationship to this part of `reltype`.

        Raises |KeyError| if no such relationship is found and |ValueError| if more than
        one such relationship is found.
        """
        return self.rels.part_with_reltype(reltype)

    def relate_to(self, target, reltype, is_external=False):
        """Return rId key of relationship of `reltype` to `target`.

        If such a relationship already exists, its rId is returned. Otherwise the
        relationship is added and its new rId returned.
        """
        if is_external:
            return self.rels.get_or_add_ext_rel(reltype, target)
        else:
            rel = self.rels.get_or_add(reltype, target)
            return rel.rId

    @property
    def related_parts(self):
        """
        Dictionary mapping related parts by rId, so child objects can resolve
        explicit relationships present in the part XML, e.g. sldIdLst to a
        specific |Slide| instance.
        """
        return self.rels.related_parts

    @lazyproperty
    def rels(self):
        """|Relationships| object containing relationships from this part to others."""
        return RelationshipCollection(self._partname.baseURI)

    def target_ref(self, rId):
        """Return URL contained in target ref of relationship identified by `rId`."""
        rel = self.rels[rId]
        return rel.target_ref

    def _blob_from_file(self, file):
        """Return bytes of `file`, which is either a str path or a file-like object."""
        # --- a str `file` is assumed to be a path ---
        if is_string(file):
            with open(file, "rb") as f:
                return f.read()

        # --- otherwise, assume `file` is a file-like object
        # --- reposition file cursor if it has one
        if callable(getattr(file, "seek")):
            file.seek(0)
        return file.read()

    def _rel_ref_count(self, rId):
        """Return int count of references in this part's XML to `rId`."""
        rIds = self._element.xpath("//@r:id")
        return len([_rId for _rId in rIds if _rId == rId])


class XmlPart(Part):
    """Base class for package parts containing an XML payload, which is most of them.

    Provides additional methods to the |Part| base class that take care of parsing and
    reserializing the XML payload and managing relationships to other parts.
    """

    def __init__(self, partname, content_type, element, package=None):
        super(XmlPart, self).__init__(partname, content_type, package=package)
        self._element = element

    @property
    def blob(self):
        """bytes XML serialization of this part."""
        return serialize_part_xml(self._element)

    @classmethod
    def load(cls, partname, content_type, blob, package):
        """Return instance of `cls` loaded with parsed XML from `blob`."""
        element = parse_xml(blob)
        return cls(partname, content_type, element, package)

    @property
    def part(self):
        """This part.

        This is part of the parent protocol, "children" of the document will not know
        the part that contains them so must ask their parent object. That chain of
        delegation ends here for child objects.
        """
        return self


class PartFactory(object):
    """Constructs a registered subtype of |Part|.

    Client code can register a subclass of |Part| to be used for a package blob based on
    its content type.
    """

    part_type_for = {}
    default_part_type = Part

    def __new__(cls, partname, content_type, blob, package):
        PartClass = cls._part_cls_for(content_type)
        return PartClass.load(partname, content_type, blob, package)

    @classmethod
    def _part_cls_for(cls, content_type):
        """Return the custom part class registered for `content_type`.

        Returns |Part| if no custom class is registered for `content_type`.
        """
        if content_type in cls.part_type_for:
            return cls.part_type_for[content_type]
        return cls.default_part_type


class RelationshipCollection(dict):
    """Collection of |_Relationship| instances, largely having dict semantics.

    Relationships are keyed by their rId, but may also be found in other ways, such as
    by their relationship type. `rels` is a dict of |Relationship| objects keyed by
    their rId.

    Note that iterating this collection generates |Relationship| references (values),
    not rIds (keys) as it would for a dict.
    """

    def __init__(self, baseURI):
        super(RelationshipCollection, self).__init__()
        self._baseURI = baseURI
        self._target_parts_by_rId = {}

    def add_relationship(self, reltype, target, rId, is_external=False):
        """
        Return a newly added |_Relationship| instance.
        """
        rel = _Relationship(rId, reltype, target, self._baseURI, is_external)
        self[rId] = rel
        if not is_external:
            self._target_parts_by_rId[rId] = target
        return rel

    def get_or_add(self, reltype, target_part):
        """
        Return relationship of *reltype* to *target_part*, newly added if not
        already present in collection.
        """
        rel = self._get_matching(reltype, target_part)
        if rel is None:
            rId = self._next_rId
            rel = self.add_relationship(reltype, target_part, rId)
        return rel

    def get_or_add_ext_rel(self, reltype, target_ref):
        """
        Return rId of external relationship of *reltype* to *target_ref*,
        newly added if not already present in collection.
        """
        rel = self._get_matching(reltype, target_ref, is_external=True)
        if rel is None:
            rId = self._next_rId
            rel = self.add_relationship(reltype, target_ref, rId, is_external=True)
        return rel.rId

    def part_with_reltype(self, reltype):
        """Return target part of relationship with matching `reltype`.

        Raises |KeyError| if not found and |ValueError| if more than one matching
        relationship is found.
        """
        rel = self._get_rel_of_type(reltype)
        return rel.target_part

    @property
    def related_parts(self):
        """
        dict mapping rIds to target parts for all the internal relationships
        in the collection.
        """
        return self._target_parts_by_rId

    @property
    def xml(self):
        """bytes XML serialization of this relationship collection.

        This value is suitable for storage as a .rels file in an OPC package. Includes
        a `<?xml` header with encoding as UTF-8.
        """
        rels_elm = CT_Relationships.new()
        for rel in self.values():
            rels_elm.add_rel(rel.rId, rel.reltype, rel.target_ref, rel.is_external)
        return rels_elm.xml

    def _get_matching(self, reltype, target, is_external=False):
        """
        Return relationship of matching *reltype*, *target*, and
        *is_external* from collection, or None if not found.
        """

        def matches(rel, reltype, target, is_external):
            if rel.reltype != reltype:
                return False
            if rel.is_external != is_external:
                return False
            rel_target = rel.target_ref if rel.is_external else rel.target_part
            if rel_target != target:
                return False
            return True

        for rel in self.values():
            if matches(rel, reltype, target, is_external):
                return rel
        return None

    def _get_rel_of_type(self, reltype):
        """
        Return single relationship of type *reltype* from the collection.
        Raises |KeyError| if no matching relationship is found. Raises
        |ValueError| if more than one matching relationship is found.
        """
        matching = [rel for rel in self.values() if rel.reltype == reltype]
        if len(matching) == 0:
            tmpl = "no relationship of type '%s' in collection"
            raise KeyError(tmpl % reltype)
        if len(matching) > 1:
            tmpl = "multiple relationships of type '%s' in collection"
            raise ValueError(tmpl % reltype)
        return matching[0]

    @property
    def _next_rId(self):
        """Next str rId available in collection.

        The next rId is the first unused key starting from "rId1" and making use of any
        gaps in numbering, e.g. 'rId2' for rIds ['rId1', 'rId3'].
        """
        for n in range(1, len(self) + 2):
            rId_candidate = "rId%d" % n  # like 'rId19'
            if rId_candidate not in self:
                return rId_candidate


class Unmarshaller(object):
    """
    Hosts static methods for unmarshalling a package from a |PackageReader|
    instance.
    """

    @staticmethod
    def unmarshal(pkg_reader, package, part_factory):
        """
        Construct graph of parts and realized relationships based on the
        contents of *pkg_reader*, delegating construction of each part to
        *part_factory*. Package relationships are added to *pkg*.
        """
        parts = Unmarshaller._unmarshal_parts(pkg_reader, package, part_factory)
        Unmarshaller._unmarshal_relationships(pkg_reader, package, parts)
        for part in parts.values():
            part.after_unmarshal()
        package.after_unmarshal()

    @staticmethod
    def _unmarshal_parts(pkg_reader, package, part_factory):
        """
        Return a dictionary of |Part| instances unmarshalled from
        *pkg_reader*, keyed by partname. Side-effect is that each part in
        *pkg_reader* is constructed using *part_factory*.
        """
        parts = {}
        for partname, content_type, blob in pkg_reader.iter_sparts():
            parts[partname] = part_factory(partname, content_type, blob, package)
        return parts

    @staticmethod
    def _unmarshal_relationships(pkg_reader, package, parts):
        """
        Add a relationship to the source object corresponding to each of the
        relationships in *pkg_reader* with its target_part set to the actual
        target part in *parts*.
        """
        for source_uri, srel in pkg_reader.iter_srels():
            source = package if source_uri == "/" else parts[source_uri]
            target = (
                srel.target_ref if srel.is_external else parts[srel.target_partname]
            )
            source.load_rel(srel.reltype, target, srel.rId, srel.is_external)


class _Relationship(object):
    """Value object describing link from a part or package to another part."""

    def __init__(self, rId, reltype, target, baseURI, external=False):
        super(_Relationship, self).__init__()
        self._rId = rId
        self._reltype = reltype
        self._target = target
        self._baseURI = baseURI
        self._is_external = bool(external)

    @property
    def is_external(self):
        """True if target_mode is `RTM.EXTERNAL`.

        An external relationship is a link to a resource outside the package, such as
        a web-resource (URL).
        """
        return self._is_external

    @property
    def reltype(self):
        """Member of RELATIONSHIP_TYPE describing relationship of target to source."""
        return self._reltype

    @property
    def rId(self):
        """str relationship-id, like 'rId9'.

        Corresponds to the `Id` attribute on the `CT_Relationship` element and
        uniquely identifies this relationship within its peers for the source-part or
        package.
        """
        return self._rId

    @property
    def target_part(self):
        """|Part| or subtype referred to by this relationship."""
        if self._is_external:
            raise ValueError(
                "target_part property on _Relationship is undef"
                "ined when target mode is External"
            )
        return self._target

    @property
    def target_ref(self):
        """str reference to relationship target.

        For internal relationships this is the relative partname, suitable for
        serialization purposes. For an external relationship it is typically a URL.
        """
        if self._is_external:
            return self._target
        else:
            return self._target.partname.relative_ref(self._baseURI)
