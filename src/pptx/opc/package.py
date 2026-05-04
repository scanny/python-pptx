"""Fundamental Open Packaging Convention (OPC) objects.

The :mod:`pptx.packaging` module coheres around the concerns of reading and writing
presentations to and from a .pptx file.
"""

from __future__ import annotations

import collections
import copy
from typing import IO, TYPE_CHECKING, DefaultDict, Iterator, Mapping, Set, cast

from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.oxml import CT_Relationships, serialize_part_xml
from pptx.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI, PackURI
from pptx.opc.serialized import PackageReader, PackageWriter
from pptx.opc.shared import CaseInsensitiveDict
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from typing_extensions import Self

    from pptx.opc.oxml import CT_Relationship, CT_Types
    from pptx.opc.serialized import ZipDateTime
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.package import Package
    from pptx.parts.presentation import PresentationPart


# -- rId-bearing attributes recognized by :class:`PartRelationshipCloner`. These correspond
# -- to the three relationship-reference attribute names used throughout OOXML: `r:id`
# -- (standard target reference, e.g. on `p:sldId`, `c:chart`, `a:hlinkClick`); `r:embed`
# -- (embedded-image reference on `a:blip`); and `r:link` (linked resource reference, e.g.
# -- on `a:videoFile`, `p:audioFile`, or `a:blip` for a linked image).
_R_ID_ATTRS: tuple[str, ...] = (
    qn("r:id"),
    qn("r:embed"),
    qn("r:link"),
)


class _RelatableMixin:
    """Provide relationship methods required by both the package and each part."""

    def part_related_by(self, reltype: str) -> Part:
        """Return (single) part having relationship to this package of `reltype`.

        Raises |KeyError| if no such relationship is found and |ValueError| if more than one such
        relationship is found.
        """
        return self._rels.part_with_reltype(reltype)

    def relate_to(self, target: Part | str, reltype: str, is_external: bool = False) -> str:
        """Return rId key of relationship of `reltype` to `target`.

        If such a relationship already exists, its rId is returned. Otherwise the relationship is
        added and its new rId returned.
        """
        if isinstance(target, str):
            assert is_external
            return self._rels.get_or_add_ext_rel(reltype, target)

        return self._rels.get_or_add(reltype, target)

    def related_part(self, rId: str) -> Part:
        """Return related |Part| subtype identified by `rId`."""
        return self._rels[rId].target_part

    def target_ref(self, rId: str) -> str:
        """Return URL contained in target ref of relationship identified by `rId`."""
        return self._rels[rId].target_ref

    @lazyproperty
    def _rels(self) -> _Relationships:
        """|_Relationships| object containing relationships from this part to others."""
        raise NotImplementedError(  # pragma: no cover
            "`%s` must implement `.rels`" % type(self).__name__
        )


class OpcPackage(_RelatableMixin):
    """Main API class for |python-opc|.

    A new instance is constructed by calling the :meth:`open` classmethod with a path to a package
    file or file-like object containing a package (.pptx file).
    """

    def __init__(self, pkg_file: str | IO[bytes], password: str | None = None):
        self._pkg_file = pkg_file
        self._password = password
        self._orphan_parts: tuple[Part, ...] = ()

    @classmethod
    def open(cls, pkg_file: str | IO[bytes], password: str | None = None) -> Self:
        """Return an |OpcPackage| instance loaded with the contents of `pkg_file`.

        When `password` is provided and the package is encrypted, the package is decrypted
        before loading. Decryption requires the optional ``msoffcrypto-tool`` dependency.
        """
        return cls(pkg_file, password)._load()

    def drop_rel(self, rId: str) -> None:
        """Remove relationship identified by `rId`."""
        self._rels.pop(rId)

    def iter_parts(self) -> Iterator[Part]:
        """Generate exactly one reference to each part in the package.

        Includes "orphan" parts — parts present in the loaded package that are declared
        in ``[Content_Types].xml`` but not reachable via the relationships graph (e.g.
        the traditional-chart fallback referenced from an ``mc:AlternateContent``
        fallback element rather than via a rel). Preserving orphans on round-trip keeps
        the output byte-level compatible with PowerPoint's expectations.
        """
        visited: Set[Part] = set()
        for rel in self.iter_rels():
            if rel.is_external:
                continue
            part = rel.target_part
            if part in visited:
                continue
            yield part
            visited.add(part)
        # -- orphan parts (and anything reachable from their rels) --
        for part in self._iter_orphan_parts(visited):
            yield part

    def iter_rels(self) -> Iterator[_Relationship]:
        """Generate exactly one reference to each relationship in package.

        Performs a depth-first traversal of the rels graph.

        Also yields relationships that hang off "orphan" parts — parts present in the
        loaded package but not reachable from the package-root rels graph (e.g. a
        fallback chart referenced via ``mc:AlternateContent`` rather than a rel). This
        ensures, for example, that an orphan chart part's embedded xlsx is preserved
        on round-trip.
        """
        visited: Set[Part] = set()

        def walk_rels(rels: _Relationships) -> Iterator[_Relationship]:
            for rel in rels.values():
                yield rel
                # --- external items can have no relationships ---
                if rel.is_external:
                    continue
                # -- all relationships other than those for the package belong to a part. Once
                # -- that part has been processed, processing it again would lead to the same
                # -- relationships appearing more than once.
                part = rel.target_part
                if part in visited:
                    continue
                visited.add(part)
                # --- recurse into relationships of each unvisited target-part ---
                yield from walk_rels(part.rels)

        yield from walk_rels(self._rels)

        # -- also walk rels of each orphan part so their dependents (e.g. an
        # -- embedded xlsx referenced from a fallback chart) are reachable too.
        for orphan in self._orphan_parts:
            if orphan in visited:
                continue
            visited.add(orphan)
            yield from walk_rels(orphan.rels)

    def _iter_orphan_parts(self, visited: Set[Part]) -> Iterator[Part]:
        """Yield each orphan part (and any descendant reachable only through it).

        `visited` is updated in place with each yielded part so the caller can use
        it as a de-duplication set across the combined rel-graph + orphan walk.
        """
        stack: list[Part] = [p for p in self._orphan_parts if p not in visited]
        while stack:
            part = stack.pop()
            if part in visited:
                continue
            visited.add(part)
            yield part
            # -- queue descendants reached through this orphan's rels --
            for rel in part.rels.values():
                if rel.is_external:
                    continue
                target = rel.target_part
                if target not in visited:
                    stack.append(target)

    @property
    def main_document_part(self) -> PresentationPart:
        """Return |Part| subtype serving as the main document part for this package.

        In this case it will be a |Presentation| part.
        """
        return cast("PresentationPart", self.part_related_by(RT.OFFICE_DOCUMENT))

    def next_partname(self, tmpl: str) -> PackURI:
        """Return |PackURI| next available partname matching `tmpl`.

        `tmpl` is a printf (%)-style template string containing a single replacement item, a '%d'
        to be used to insert the integer portion of the partname. Example:
        '/ppt/slides/slide%d.xml'
        """
        # --- The allocated partnames for `tmpl` are cached the first time `tmpl` is
        # --- encountered. Subsequent calls reuse that cache, avoiding the O(N)
        # --- `iter_parts()` walk on each allocation that would otherwise produce
        # --- O(N**2) behavior when constructing a large presentation (#644). The
        # --- newly allocated partname is added to the cache before return so the
        # --- cache remains authoritative for subsequent calls.
        allocated = self._partnames_by_tmpl.get(tmpl)
        if allocated is None:
            prefix = tmpl[: (tmpl % 42).find("42")]
            allocated = {p.partname for p in self.iter_parts() if p.partname.startswith(prefix)}
            self._partnames_by_tmpl[tmpl] = allocated
        # --- The common case is an unbroken sequence starting at 1, so
        # --- `len(allocated) + 1` is the expected answer. Scan downward from there to
        # --- fill any gaps while still short-circuiting on the first candidate in the
        # --- common case.
        for n in range(len(allocated) + 1, 0, -1):
            candidate_partname = tmpl % n
            if candidate_partname not in allocated:
                partname = PackURI(candidate_partname)
                allocated.add(partname)
                return partname
        raise Exception("ProgrammingError: ran out of candidate_partnames")  # pragma: no cover

    def save(
        self,
        pkg_file: str | IO[bytes],
        zip_date_time: ZipDateTime | None = None,
        password: str | None = None,
    ) -> None:
        """Save this package to `pkg_file`.

        `file` can be either a path to a file (a string) or a file-like object.

        When `zip_date_time` is provided, every zip-member in the saved package is stamped
        with that fixed last-modified timestamp, yielding byte-identical output for repeated
        saves of identical content (reproducible-build / source-control friendly).

        When `password` is provided, the resulting .pptx is password-protected using
        ECMA-376 Agile Encryption; this requires the optional ``msoffcrypto-tool``
        dependency. The two keywords are orthogonal: `zip_date_time` stamps the inner
        (plaintext) zip members before the package is wrapped in the encryption
        container.
        """
        PackageWriter.write(
            pkg_file,
            self._rels,
            tuple(self.iter_parts()),
            zip_date_time=zip_date_time,
            password=password,
        )

    def save_flat_xml(self, pkg_file: str | IO[bytes]) -> None:
        """Save this package to `pkg_file` as a Flat OPC (single-XML) document.

        Flat OPC is the "XML Presentation" format defined in ECMA-376 Part 4:
        the entire package is serialized as a single XML document with one
        ``<pkg:part>`` child per package item. XML parts are embedded inline,
        binary parts (images, embedded fonts, OLE, media) are base64-encoded.

        `pkg_file` is either a filesystem path (``str``) or a file-like object
        opened for binary writing.
        """
        # -- imported lazily to keep the hot-path save unaffected --
        from pptx.opc.flat_opc import FlatOpcWriter

        FlatOpcWriter.write(pkg_file, self._rels, tuple(self.iter_parts()))


    def _load(self) -> Self:
        """Return the package after loading all parts and relationships."""
        pkg_xml_rels, parts, orphan_partnames = _PackageLoader.load(
            self._pkg_file, cast("Package", self), self._password
        )
        self._rels.load_from_xml(PACKAGE_URI, pkg_xml_rels, parts)
        # -- record orphan parts (declared in [Content_Types].xml but not reachable
        # -- from the package-root rels graph) so they survive round-trip. Any parts
        # -- already reachable via the rels graph are filtered out to preserve the
        # -- rel-walk as the source of truth for everything else.
        reachable: Set[Part] = {p for p in self.iter_parts()}
        self._orphan_parts = tuple(
            parts[pn] for pn in orphan_partnames if parts[pn] not in reachable
        )
        return self

    @lazyproperty
    def _partnames_by_tmpl(self) -> dict[str, set[PackURI]]:
        """Cached {tmpl: {allocated_partnames}} used by :meth:`next_partname`.

        Populated lazily on first use of a given `tmpl` by scanning the current part graph
        for partnames sharing its prefix. Each subsequent allocation for that `tmpl` is
        added to the cache, so the expensive `iter_parts()` walk happens at most once per
        template per package instance. See issue #644.
        """
        return {}

    @lazyproperty
    def _rels(self) -> _Relationships:
        """|Relationships| object containing relationships of this package."""
        return _Relationships(PACKAGE_URI.baseURI)


class _PackageLoader:
    """Function-object that loads a package from disk (or other store)."""

    def __init__(self, pkg_file: str | IO[bytes], package: Package, password: str | None = None):
        self._pkg_file = pkg_file
        self._package = package
        self._password = password
        # -- populated by :attr:`_xml_rels` with partnames declared in
        # -- [Content_Types].xml but not reached via the rels graph.
        self._orphan_partnames: tuple[PackURI, ...] = ()
        # -- populated by :meth:`_xml_rels_for` with each partname whose ``.rels`` file
        # -- was physically present in the input package (even if empty). Consumed by
        # -- :meth:`_load` to flag the corresponding part for round-trip preservation.
        self._partnames_with_rels_file: Set[PackURI] = set()

    @classmethod
    def load(
        cls,
        pkg_file: str | IO[bytes],
        package: Package,
        password: str | None = None,
    ) -> tuple[CT_Relationships, dict[PackURI, Part], tuple[PackURI, ...]]:
        """Return (pkg_xml_rels, parts, orphan_partnames) triple from loading `pkg_file`.

        The returned `parts` value is a {partname: part} mapping with each part in the package
        included and constructed complete with its relationships to other parts in the package.

        The returned `pkg_xml_rels` value is a `CT_Relationships` object containing the parsed
        package relationships. It is the caller's responsibility (the package object) to load
        those relationships into its |_Relationships| object.

        `orphan_partnames` lists partnames that were declared in ``[Content_Types].xml``
        (as ``Override`` entries) but are not reachable via the package-root rels graph.
        Preserving these on round-trip keeps parts such as an ``mc:AlternateContent``
        fallback chart (and its embedded workbook) present in the output package.
        """
        return cls(pkg_file, package, password)._load()

    def _load(
        self,
    ) -> tuple[CT_Relationships, dict[PackURI, Part], tuple[PackURI, ...]]:
        """Return (pkg_xml_rels, parts, orphan_partnames) triple from loading pkg_file."""
        parts, xml_rels = self._parts, self._xml_rels

        for partname, part in parts.items():
            part.load_rels_from_xml(xml_rels[partname], parts)
            # -- propagate whether a physical ``.rels`` file was present on disk for this
            # -- part, so the writer can preserve an empty rels file on round-trip when
            # -- one was present in the input package.
            if partname in self._partnames_with_rels_file:
                part._rels_file_present_on_load = True  # pyright: ignore[reportPrivateUsage]

        return xml_rels[PACKAGE_URI], parts, self._orphan_partnames

    @lazyproperty
    def _content_types(self) -> _ContentTypeMap:
        """|_ContentTypeMap| object providing content-types for items of this package.

        Provides a content-type (MIME-type) for any given partname.
        """
        return _ContentTypeMap.from_xml(self._package_reader[CONTENT_TYPES_URI])

    @lazyproperty
    def _package_reader(self) -> PackageReader:
        """|PackageReader| object providing access to package-items in pkg_file."""
        return PackageReader(self._pkg_file, self._password)

    @lazyproperty
    def _parts(self) -> dict[PackURI, Part]:
        """dict {partname: Part} populated with parts loading from package.

        Among other duties, this collection is passed to each relationships collection so each
        relationship can resolve a reference to its target part when required. This reference can
        only be reliably carried out once the all parts have been loaded.
        """
        content_types = self._content_types
        package = self._package
        package_reader = self._package_reader

        return {
            partname: PartFactory(
                partname,
                content_types[partname],
                package,
                blob=package_reader[partname],
            )
            for partname in (p for p in self._xml_rels if p != "/")
            # -- invalid partnames can arise in some packages; ignore those rather than raise an
            # -- exception.
            if partname in package_reader
        }

    @lazyproperty
    def _xml_rels(self) -> dict[PackURI, CT_Relationships]:
        """dict {partname: xml_rels} for package and all package parts.

        This is used as the basis for other loading operations such as loading parts and
        populating their relationships.

        In addition to the parts reachable from the package-root rels graph, also includes
        any "orphan" parts — those declared in ``[Content_Types].xml`` as ``Override``
        entries but not reachable via rels. These can arise from ``mc:AlternateContent``
        fallback elements (e.g. a traditional chart associated with an extended
        ``chartEx`` part). Their rels are loaded too, so descendants (for example the
        fallback chart's embedded xlsx) come along for the round-trip.
        """
        xml_rels: dict[PackURI, CT_Relationships] = {}
        visited_partnames: Set[PackURI] = set()

        def load_rels(source_partname: PackURI, rels: CT_Relationships):
            """Populate `xml_rels` dict by traversing relationships depth-first."""
            xml_rels[source_partname] = rels
            visited_partnames.add(source_partname)
            base_uri = source_partname.baseURI

            # --- recursion stops when there are no unvisited partnames in rels ---
            for rel in rels.relationship_lst:
                if rel.targetMode == RTM.EXTERNAL:
                    continue
                target_partname = PackURI.from_rel_ref(base_uri, rel.target_ref)
                if target_partname in visited_partnames:
                    continue
                load_rels(target_partname, self._xml_rels_for(target_partname))

        load_rels(PACKAGE_URI, self._xml_rels_for(PACKAGE_URI))

        # -- Include orphans: parts declared in [Content_Types].xml Overrides but not
        # -- reached via the rels graph. We walk each orphan's rels too, so anything
        # -- reachable solely through an orphan (e.g. the xlsx behind a fallback
        # -- chart) is also materialised during load.
        package_reader = self._package_reader
        orphans: list[PackURI] = []
        for partname in self._content_types.override_partnames:
            if partname in visited_partnames:
                continue
            if partname not in package_reader:
                continue
            orphans.append(partname)
            load_rels(partname, self._xml_rels_for(partname))
        # -- stash the orphan partname list for the package to consume after
        # -- `_parts` has been built.
        self._orphan_partnames = tuple(orphans)
        return xml_rels

    def _xml_rels_for(self, partname: PackURI) -> CT_Relationships:
        """Return CT_Relationships object formed by parsing rels XML for `partname`.

        A CT_Relationships object is returned in all cases. A part that has no relationships
        receives an "empty" CT_Relationships object, i.e. containing no `CT_Relationship` objects.

        Records `partname` in :attr:`_partnames_with_rels_file` when a physical rels file
        was found on disk (even an empty one). The writer uses this set to preserve an
        empty rels file on round-trip when one was present in the input package.
        """
        rels_xml = self._package_reader.rels_xml_for(partname)
        if rels_xml is None:
            return CT_Relationships.new()
        self._partnames_with_rels_file.add(partname)
        return cast(CT_Relationships, parse_xml(rels_xml))


class Part(_RelatableMixin):
    """Base class for package parts.

    Provides common properties and methods, but intended to be subclassed in client code to
    implement specific part behaviors. Also serves as the default class for parts that are not yet
    given specific behaviors.
    """

    def __init__(
        self, partname: PackURI, content_type: str, package: Package, blob: bytes | None = None
    ):
        # --- XmlPart subtypes, don't store a blob (the original XML) ---
        self._partname = partname
        self._content_type = content_type
        self._package = package
        self._blob = blob
        # -- True when this part was loaded from a package that had a ``.rels`` file for
        # -- it, even if that file was empty. Used by the package writer to preserve an
        # -- empty rels file on round-trip when one was physically present in the input.
        self._rels_file_present_on_load = False

    @classmethod
    def load(cls, partname: PackURI, content_type: str, package: Package, blob: bytes) -> Self:
        """Return `cls` instance loaded from arguments.

        This one is a straight pass-through, but subtypes may do some pre-processing, see XmlPart
        for an example.
        """
        return cls(partname, content_type, package, blob)

    @property
    def blob(self) -> bytes:
        """Contents of this package part as a sequence of bytes.

        Intended to be overridden by subclasses. Default behavior is to return the blob initial
        loaded during `Package.open()` operation.
        """
        return self._blob or b""

    @blob.setter
    def blob(self, blob: bytes):
        """Note that not all subclasses use the part blob as their blob source.

        In particular, the |XmlPart| subclass uses its `self._element` to serialize a blob on
        demand. This works fine for binary parts though.
        """
        self._blob = blob

    @lazyproperty
    def content_type(self) -> str:
        """Content-type (MIME-type) of this part."""
        return self._content_type

    def load_rels_from_xml(self, xml_rels: CT_Relationships, parts: dict[PackURI, Part]) -> None:
        """load _Relationships for this part from `xml_rels`.

        Part references are resolved using the `parts` dict that maps each partname to the loaded
        part with that partname. These relationships are loaded from a serialized package and so
        already have assigned rIds. This method is only used during package loading.
        """
        self._rels.load_from_xml(self._partname.baseURI, xml_rels, parts)

    @lazyproperty
    def package(self) -> Package:
        """Package this part belongs to."""
        return self._package

    @property
    def partname(self) -> PackURI:
        """|PackURI| partname for this part, e.g. "/ppt/slides/slide1.xml"."""
        return self._partname

    @partname.setter
    def partname(self, partname: PackURI):
        if not isinstance(partname, PackURI):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(  # pragma: no cover
                "partname must be instance of PackURI, got '%s'" % type(partname).__name__
            )
        self._partname = partname

    @lazyproperty
    def rels(self) -> _Relationships:
        """Collection of relationships from this part to other parts."""
        # --- this must be public to allow the part graph to be traversed ---
        return self._rels

    def _blob_from_file(self, file: str | IO[bytes]) -> bytes:
        """Return bytes of `file`, which is either a str path or a file-like object."""
        # --- a str `file` is assumed to be a path ---
        if isinstance(file, str):
            with open(file, "rb") as f:
                return f.read()

        # --- otherwise, assume `file` is a file-like object
        # --- reposition file cursor if it has one
        if callable(getattr(file, "seek")):
            file.seek(0)
        return file.read()

    def _new_rId(self) -> str:
        """Return str rId that is not yet used in this part's relationships.

        This is a thin, package-public helper around the same logic used by
        :meth:`_Relationships.get_or_add`. Primarily consumed by
        :class:`PartRelationshipCloner` to pre-compute rId remappings before relationships
        are added.
        """
        return self._rels._next_rId  # pyright: ignore[reportPrivateUsage]

    def _next_partname(self, tmpl: str) -> PackURI:
        """Return |PackURI| next available partname matching `tmpl` in the package.

        `tmpl` is a printf (%)-style template string containing a single `%d` replacement,
        e.g. `"/ppt/slides/slide%d.xml"`. The returned partname is guaranteed not to
        collide with any part already present in the package.

        A convenience wrapper over :meth:`OpcPackage.next_partname` so callers that only
        hold a reference to a `Part` don't have to reach back into the package.
        """
        return self._package.next_partname(tmpl)

    @lazyproperty
    def _rels(self) -> _Relationships:
        """Relationships from this part to others."""
        return _Relationships(self._partname.baseURI)


class XmlPart(Part):
    """Base class for package parts containing an XML payload, which is most of them.

    Provides additional methods to the |Part| base class that take care of parsing and
    reserializing the XML payload and managing relationships to other parts.
    """

    def __init__(
        self, partname: PackURI, content_type: str, package: Package, element: BaseOxmlElement
    ):
        super(XmlPart, self).__init__(partname, content_type, package)
        self._element = element

    @classmethod
    def load(cls, partname: PackURI, content_type: str, package: Package, blob: bytes):
        """Return instance of `cls` loaded with parsed XML from `blob`."""
        return cls(
            partname, content_type, package, element=cast("BaseOxmlElement", parse_xml(blob))
        )

    @property
    def blob(self) -> bytes:  # pyright: ignore[reportIncompatibleMethodOverride]
        """bytes XML serialization of this part."""
        return serialize_part_xml(self._element)

    # -- XmlPart cannot set its blob, which is why pyright complains --

    def drop_rel(self, rId: str) -> None:
        """Remove relationship identified by `rId` if its reference count is under 2.

        Relationships with a reference count of 0 are implicit relationships. Note that only XML
        parts can drop relationships.
        """
        if self._rel_ref_count(rId) < 2:
            self._rels.pop(rId)

    @property
    def part(self):
        """This part.

        This is part of the parent protocol, "children" of the document will not know the part
        that contains them so must ask their parent object. That chain of delegation ends here for
        child objects.
        """
        return self

    def _rel_ref_count(self, rId: str) -> int:
        """Return int count of references in this part's XML to `rId`."""
        return len([r for r in cast("list[str]", self._element.xpath("//@r:id")) if r == rId])


class PartFactory:
    """Constructs a registered subtype of |Part|.

    Client code can register a subclass of |Part| to be used for a package blob based on its
    content type.
    """

    part_type_for: dict[str, type[Part]] = {}

    def __new__(cls, partname: PackURI, content_type: str, package: Package, blob: bytes) -> Part:
        PartClass = cls._part_cls_for(content_type)
        return PartClass.load(partname, content_type, package, blob)

    @classmethod
    def _part_cls_for(cls, content_type: str) -> type[Part]:
        """Return the custom part class registered for `content_type`.

        Returns |Part| if no custom class is registered for `content_type`.

        Lookup is case-insensitive. Per RFC 2046 §4.1 MIME type tokens are
        case-insensitive, and real-world `.pptx` files occasionally carry
        content-type strings whose casing differs from python-pptx's canonical
        registration (e.g. ``Image/Tiff`` vs ``image/tiff``). Without this
        normalization such parts fall through to the generic |Part|, producing
        ``AttributeError: 'Part' object has no attribute 'image'`` at access
        time — see issue #1084.
        """
        if content_type in cls.part_type_for:
            return cls.part_type_for[content_type]
        # -- fall back to case-insensitive match against the registered keys --
        lowered = content_type.lower()
        if lowered != content_type and lowered in cls.part_type_for:
            return cls.part_type_for[lowered]
        return Part


class _ContentTypeMap:
    """Value type providing dict semantics for looking up content type by partname."""

    def __init__(self, overrides: dict[str, str], defaults: dict[str, str]):
        self._overrides = overrides
        self._defaults = defaults

    def __getitem__(self, partname: PackURI) -> str:
        """Return content-type (MIME-type) for part identified by *partname*."""
        if not isinstance(partname, PackURI):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(
                "_ContentTypeMap key must be <type 'PackURI'>, got %s" % type(partname).__name__
            )

        if partname in self._overrides:
            return self._overrides[partname]

        if partname.ext in self._defaults:
            return self._defaults[partname.ext]

        raise KeyError("no content-type for partname '%s' in [Content_Types].xml" % partname)

    @classmethod
    def from_xml(cls, content_types_xml: bytes) -> _ContentTypeMap:
        """Return |_ContentTypeMap| instance populated from `content_types_xml`."""
        types_elm = cast("CT_Types", parse_xml(content_types_xml))
        # -- note all partnames in [Content_Types].xml are absolute --
        overrides = CaseInsensitiveDict(
            (o.partName.lower(), o.contentType) for o in types_elm.override_lst
        )
        defaults = CaseInsensitiveDict(
            (d.extension.lower(), d.contentType) for d in types_elm.default_lst
        )
        return cls(overrides, defaults)

    @property
    def override_partnames(self) -> tuple[PackURI, ...]:
        """Tuple of |PackURI| partnames with an explicit ``Override`` content-type.

        Useful for discovering parts that are present in the package but not reachable
        via the rels graph, so they can still be preserved on round-trip.
        """
        return tuple(PackURI(pn) for pn in self._overrides)


class _Relationships(Mapping[str, "_Relationship"]):
    """Collection of |_Relationship| instances having `dict` semantics.

    Relationships are keyed by their rId, but may also be found in other ways, such as by their
    relationship type. |Relationship| objects are keyed by their rId.

    Iterating this collection has normal mapping semantics, generating the keys (rIds) of the
    mapping. `rels.keys()`, `rels.values()`, and `rels.items() can be used as they would be for a
    `dict`.
    """

    def __init__(self, base_uri: str):
        self._base_uri = base_uri

    def __contains__(self, rId: object) -> bool:
        """Implement 'in' operation, like `"rId7" in relationships`."""
        return rId in self._rels

    def __getitem__(self, rId: str) -> _Relationship:
        """Implement relationship lookup by rId using indexed access, like rels[rId]."""
        try:
            return self._rels[rId]
        except KeyError:
            raise KeyError("no relationship with key '%s'" % rId)

    def __iter__(self) -> Iterator[str]:
        """Implement iteration of rIds (iterating a mapping produces its keys)."""
        return iter(self._rels)

    def __len__(self) -> int:
        """Return count of relationships in collection."""
        return len(self._rels)

    def get_or_add(self, reltype: str, target_part: Part) -> str:
        """Return str rId of `reltype` to `target_part`.

        The rId of an existing matching relationship is used if present. Otherwise, a new
        relationship is added and that rId is returned.
        """
        existing_rId = self._get_matching(reltype, target_part)
        return (
            self._add_relationship(reltype, target_part) if existing_rId is None else existing_rId
        )

    def get_or_add_ext_rel(self, reltype: str, target_ref: str) -> str:
        """Return str rId of external relationship of `reltype` to `target_ref`.

        The rId of an existing matching relationship is used if present. Otherwise, a new
        relationship is added and that rId is returned.
        """
        existing_rId = self._get_matching(reltype, target_ref, is_external=True)
        return (
            self._add_relationship(reltype, target_ref, is_external=True)
            if existing_rId is None
            else existing_rId
        )

    def load_from_xml(
        self, base_uri: str, xml_rels: CT_Relationships, parts: dict[PackURI, Part]
    ) -> None:
        """Replace any relationships in this collection with those from `xml_rels`."""

        def iter_valid_rels():
            """Filter out broken relationships such as those pointing to NULL."""
            for rel_elm in xml_rels.relationship_lst:
                # --- Occasionally a PowerPoint plugin or other client will "remove"
                # --- a relationship simply by "voiding" its Target value, like making
                # --- it "/ppt/slides/NULL". Skip any relationships linking to a
                # --- partname that is not present in the package.
                if rel_elm.targetMode == RTM.INTERNAL:
                    partname = PackURI.from_rel_ref(base_uri, rel_elm.target_ref)
                    if partname not in parts:
                        continue
                yield _Relationship.from_xml(base_uri, rel_elm, parts)

        self._rels.clear()
        self._rels.update((rel.rId, rel) for rel in iter_valid_rels())

    def part_with_reltype(self, reltype: str) -> Part:
        """Return target part of relationship with matching `reltype`.

        Raises |KeyError| if not found and |ValueError| if more than one matching relationship is
        found.
        """
        rels_of_reltype = self._rels_by_reltype[reltype]

        if len(rels_of_reltype) == 0:
            raise KeyError("no relationship of type '%s' in collection" % reltype)

        if len(rels_of_reltype) > 1:
            raise ValueError("multiple relationships of type '%s' in collection" % reltype)

        return rels_of_reltype[0].target_part

    def pop(self, rId: str) -> _Relationship:
        """Return |_Relationship| identified by `rId` after removing it from collection.

        The caller is responsible for ensuring it is no longer required.
        """
        return self._rels.pop(rId)

    @property
    def xml(self):
        """bytes XML serialization of this relationship collection.

        This value is suitable for storage as a .rels file in an OPC package. Includes a `<?xml..`
        declaration header with encoding as UTF-8.
        """
        rels_elm = CT_Relationships.new()

        # -- Sequence <Relationship> elements deterministically (in numerical order) to
        # -- simplify testing and manual inspection.
        def iter_rels_in_numerical_order():
            sorted_num_rId_pairs = sorted(
                (
                    int(rId[3:]) if rId.startswith("rId") and rId[3:].isdigit() else 0,
                    rId,
                )
                for rId in self.keys()
            )
            return (self[rId] for _, rId in sorted_num_rId_pairs)

        for rel in iter_rels_in_numerical_order():
            rels_elm.add_rel(rel.rId, rel.reltype, rel.target_ref, rel.is_external)

        return rels_elm.xml_file_bytes

    def _add_relationship(self, reltype: str, target: Part | str, is_external: bool = False) -> str:
        """Return str rId of |_Relationship| newly added to spec."""
        rId = self._next_rId
        self._rels[rId] = _Relationship(
            self._base_uri,
            rId,
            reltype,
            target_mode=RTM.EXTERNAL if is_external else RTM.INTERNAL,
            target=target,
        )
        return rId

    def _get_matching(
        self, reltype: str, target: Part | str, is_external: bool = False
    ) -> str | None:
        """Return optional str rId of rel of `reltype`, `target`, and `is_external`.

        Returns `None` on no matching relationship
        """
        for rel in self._rels_by_reltype[reltype]:
            if rel.is_external != is_external:
                continue
            rel_target = rel.target_ref if rel.is_external else rel.target_part
            if rel_target == target:
                return rel.rId

        return None

    @property
    def _next_rId(self) -> str:
        """Next str rId available in collection.

        The next rId is the first unused key starting from "rId1" and making use of any gaps in
        numbering, e.g. 'rId2' for rIds ['rId1', 'rId3'].
        """
        # --- The common case is where all sequential numbers starting at "rId1" are
        # --- used and the next available rId is "rId%d" % (len(rels)+1). So we start
        # --- there and count down to produce the best performance.
        for n in range(len(self) + 1, 0, -1):
            rId_candidate = "rId%d" % n  # like 'rId19'
            if rId_candidate not in self._rels:
                return rId_candidate
        raise Exception(
            "ProgrammingError: Impossible to have more distinct rIds than relationships"
        )

    @lazyproperty
    def _rels(self) -> dict[str, _Relationship]:
        """dict {rId: _Relationship} containing relationships of this collection."""
        return {}

    @property
    def _rels_by_reltype(self) -> dict[str, list[_Relationship]]:
        """defaultdict {reltype: [rels]} for all relationships in collection."""
        D: DefaultDict[str, list[_Relationship]] = collections.defaultdict(list)
        for rel in self.values():
            D[rel.reltype].append(rel)
        return D


class _Relationship:
    """Value object describing link from a part or package to another part."""

    def __init__(self, base_uri: str, rId: str, reltype: str, target_mode: str, target: Part | str):
        self._base_uri = base_uri
        self._rId = rId
        self._reltype = reltype
        self._target_mode = target_mode
        self._target = target

    @classmethod
    def from_xml(
        cls, base_uri: str, rel: CT_Relationship, parts: dict[PackURI, Part]
    ) -> _Relationship:
        """Return |_Relationship| object based on CT_Relationship element `rel`."""
        target = (
            rel.target_ref
            if rel.targetMode == RTM.EXTERNAL
            else parts[PackURI.from_rel_ref(base_uri, rel.target_ref)]
        )
        return cls(base_uri, rel.rId, rel.reltype, rel.targetMode, target)

    @lazyproperty
    def is_external(self) -> bool:
        """True if target_mode is `RTM.EXTERNAL`.

        An external relationship is a link to a resource outside the package, such as a
        web-resource (URL).
        """
        return self._target_mode == RTM.EXTERNAL

    @lazyproperty
    def reltype(self) -> str:
        """Member of RELATIONSHIP_TYPE describing relationship of target to source."""
        return self._reltype

    @lazyproperty
    def rId(self) -> str:
        """str relationship-id, like 'rId9'.

        Corresponds to the `Id` attribute on the `CT_Relationship` element and uniquely identifies
        this relationship within its peers for the source-part or package.
        """
        return self._rId

    @lazyproperty
    def target_part(self) -> Part:
        """|Part| or subtype referred to by this relationship."""
        if self.is_external:
            raise ValueError(
                "`.target_part` property on _Relationship is undefined when "
                "target-mode is external"
            )
        assert isinstance(self._target, Part)
        return self._target

    @lazyproperty
    def target_partname(self) -> PackURI:
        """|PackURI| instance containing partname targeted by this relationship.

        Raises `ValueError` on reference if target_mode is external. Use :attr:`target_mode` to
        check before referencing.
        """
        if self.is_external:
            raise ValueError(
                "`.target_partname` property on _Relationship is undefined when "
                "target-mode is external"
            )
        assert isinstance(self._target, Part)
        return self._target.partname

    @lazyproperty
    def target_ref(self) -> str:
        """str reference to relationship target.

        For internal relationships this is the relative partname, suitable for serialization
        purposes. For an external relationship it is typically a URL.
        """
        if self.is_external:
            assert isinstance(self._target, str)
            return self._target

        return self.target_partname.relative_ref(self._base_uri)


class PartRelationshipCloner:
    """Clones an XML element plus every Part relationship it references.

    Given a `src_element` belonging to `src_part` and a `tgt_part` to receive the clone,
    walks every `r:id` / `r:embed` / `r:link` attribute in the source element (and its
    descendants), clones each referenced part into the target package (reusing an existing
    part when the source already belongs to the target package; materialising a
    non-colliding clone otherwise), adds a corresponding relationship on `tgt_part`, and
    rewrites the rId values in the returned cloned element so they match the new
    relationships on `tgt_part`.

    Intended use-cases:

    * Slide duplication (copy a slide within, or across, a presentation).
    * Shape duplication (copy a shape from one slide to another; brings its picture /
      chart / OLE / media parts along).
    * Chart-copy (clone a chart plus its embedded workbook).
    * Any other operation that needs to move "an XML element plus every file-backed thing
      it points at" between OPC parts.

    External (`TargetMode="External"`) relationships such as hyperlinks are cloned as
    plain string references — no new part is materialised.

    Note: this helper deliberately does *not* recurse into the relationships *of* a cloned
    target part (e.g. a chart part that points at an embedded xlsx). Cloning those "deep"
    relationships is the responsibility of a content-aware helper layered on top of this
    one. Keeping the scope narrow keeps the behaviour predictable and bounded.
    """

    def __init__(
        self,
        src_part: Part,
        tgt_part: Part,
        src_element: BaseOxmlElement,
    ):
        self._src_part = src_part
        self._tgt_part = tgt_part
        self._src_element = src_element

    @classmethod
    def clone(
        cls,
        src_part: Part,
        tgt_part: Part,
        src_element: BaseOxmlElement,
    ) -> BaseOxmlElement:
        """Return a deep-copy of `src_element` with its rIds remapped against `tgt_part`.

        Every Part referenced from `src_element` via `r:id`, `r:embed`, or `r:link` is
        ensured to exist in `tgt_part`'s package (cloned if necessary), a relationship is
        established from `tgt_part` to each such part, and the rId values in the returned
        element are rewritten to the newly-assigned rIds. `src_element` itself is not
        modified.
        """
        return cls(src_part, tgt_part, src_element)._clone()

    def _clone(self) -> BaseOxmlElement:
        """Implementation of :meth:`clone`."""
        # -- deep-copy up-front so the caller's source element is never mutated; all rId
        # -- rewrites happen on the clone --
        new_element: BaseOxmlElement = copy.deepcopy(self._src_element)

        # -- build a {old_rId -> new_rId} mapping once, then apply it to every rId-bearing
        # -- attribute in the clone. Each distinct old rId produces at most one target
        # -- relationship (and at most one cloned part). --
        rId_map = self._build_rId_map(new_element)
        if rId_map:
            self._remap_rIds(new_element, rId_map)
        return new_element

    # -- helpers ---------------------------------------------

    def _build_rId_map(self, clone: BaseOxmlElement) -> dict[str, str]:
        """Return `{old_rId: new_rId}` mapping for every distinct rId in `clone`."""
        rId_map: dict[str, str] = {}
        for old_rId in self._iter_rIds(clone):
            if old_rId in rId_map:
                continue
            rId_map[old_rId] = self._clone_relationship(old_rId)
        return rId_map

    def _iter_rId_bearing_elements(self, element: BaseOxmlElement) -> Iterator[BaseOxmlElement]:
        """Generate every element in `element`'s subtree (including itself) that carries an rId.

        Preserves document order so callers can rely on stable rId allocation when the
        mapping is built.
        """
        # -- the "descendant-or-self" axis picks up `element` and every descendant --
        xpath_expr = "descendant-or-self::*[@r:id or @r:embed or @r:link]"
        results = cast("list[BaseOxmlElement]", element.xpath(xpath_expr))
        yield from results

    def _clone_relationship(self, old_rId: str) -> str:
        """Return new rId on `tgt_part` for the relationship keyed by `old_rId` on `src_part`.

        Handles external relationships (e.g. hyperlinks) and internal relationships
        (referenced parts) uniformly. The returned rId is freshly allocated on the target
        part's relationship collection.
        """
        src_rel = self._src_part.rels[old_rId]

        if src_rel.is_external:
            # -- external references (hyperlinks, linked media) carry a URI rather than a
            # -- part; duplicate the URI verbatim. `relate_to` returns an existing rId if
            # -- the same URI is already linked, which is the desired behaviour. --
            return self._tgt_part.relate_to(src_rel.target_ref, src_rel.reltype, is_external=True)

        # -- internal: make sure the target package has a Part with the same content --
        src_target_part = src_rel.target_part
        tgt_target_part = self._get_or_clone_part(src_target_part)
        return self._tgt_part.relate_to(tgt_target_part, src_rel.reltype)

    def _get_or_clone_part(self, src_target_part: Part) -> Part:
        """Return Part in tgt package that mirrors `src_target_part`.

        If the source part already belongs to the target package (same-package clone),
        reuse it directly — that's both safe and efficient. Otherwise, materialise a
        shallow duplicate: a new Part with a non-colliding partname in the target
        package, the same content-type, and the source part's binary blob.

        "Shallow" here means relationships *within* `src_target_part` are *not* recursed
        into. Callers that need deep graph cloning (e.g. a chart part that references an
        embedded xlsx) must compose that logic on top of this helper.
        """
        tgt_package = self._tgt_part.package
        if src_target_part.package is tgt_package:
            return src_target_part

        # -- derive a non-colliding partname in the target package using the source
        # -- part's partname template (e.g. "/ppt/media/image%d.png"). We can't reuse the
        # -- literal source partname because the target package may already have a part at
        # -- that URI. --
        partname_tmpl = _partname_template_for(src_target_part.partname)
        new_partname = tgt_package.next_partname(partname_tmpl)
        blob = src_target_part.blob
        cloned = type(src_target_part)(
            new_partname, src_target_part.content_type, tgt_package, blob
        )
        return cloned

    def _iter_rIds(self, element: BaseOxmlElement) -> Iterator[str]:
        """Generate rId values found in `element` and its descendants, in document order."""
        for el in self._iter_rId_bearing_elements(element):
            for attr in _R_ID_ATTRS:
                value = el.get(attr)
                if value is not None:
                    yield value

    def _remap_rIds(self, element: BaseOxmlElement, rId_map: Mapping[str, str]) -> None:
        """Rewrite every rId attribute on `element` (or descendants) using `rId_map`."""
        for el in self._iter_rId_bearing_elements(element):
            for attr in _R_ID_ATTRS:
                value = el.get(attr)
                if value is None:
                    continue
                new_value = rId_map.get(value)
                if new_value is None:  # pragma: no cover -- defensive: unknown rId
                    continue
                el.set(attr, new_value)


def _partname_template_for(partname: PackURI) -> str:
    """Return a printf-style template derived from `partname`.

    The returned template contains a single `%d` replacement in the trailing numeric
    position. For `/ppt/media/image3.png` this returns `/ppt/media/image%d.png`. For a
    partname with no numeric suffix (e.g. `/ppt/theme/theme.xml`) the template is
    returned with `%d` inserted before the extension: `/ppt/theme/theme%d.xml`.
    """
    s = str(partname)
    # -- split off extension (everything after the last '.') --
    dot = s.rfind(".")
    if dot == -1:
        stem, ext = s, ""
    else:
        stem, ext = s[:dot], s[dot:]

    # -- strip trailing digits from stem to find the numeric suffix position --
    i = len(stem)
    while i > 0 and stem[i - 1].isdigit():
        i -= 1
    prefix = stem[:i]
    return f"{prefix}%d{ext}"
