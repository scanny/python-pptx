# encoding: utf-8

"""API for reading/writing serialized Open Packaging Convention (OPC) package."""

import os
import zipfile

from pptx.compat import Container, is_string
from pptx.exceptions import PackageNotFoundError
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.oxml import CT_Types, parse_xml, serialize_part_xml
from pptx.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI, PackURI
from pptx.opc.shared import CaseInsensitiveDict
from pptx.opc.spec import default_content_types
from pptx.util import lazyproperty


class PackageReader(Container):
    """Provides access to package-parts of an OPC package with dict semantics.

    The package may be in zip-format (a .pptx file) or expanded into a directory
    structure, perhaps by unzipping a .pptx file.
    """

    def __init__(self, pkg_file):
        self._pkg_file = pkg_file

    def __contains__(self, pack_uri):
        """Return True when part identified by `pack_uri` is present in package."""
        return pack_uri in self._blob_reader

    def iter_sparts(self):
        """
        Generate a 3-tuple `(partname, content_type, blob)` for each of the
        serialized parts in the package.
        """
        for spart in self._sparts:
            yield (spart.partname, spart.content_type, spart.blob)

    def iter_srels(self):
        """
        Generate a 2-tuple `(source_uri, srel)` for each of the relationships
        in the package.
        """
        for srel in self._pkg_srels:
            yield (PACKAGE_URI, srel)
        for spart in self._sparts:
            for srel in spart.srels:
                yield (spart.partname, srel)

    def rels_xml_for(self, pkg_file, partname):
        """Return optional rels item XML for `partname`.

        Returns `None` if no rels item is present for `partname`. `partname` is a
        |PackURI| instance.
        """
        # --- ugly temporary hack to make this interim `._rels_xml_for()` method
        # --- produce the same result as the one that's coming a few commits later.
        phys_reader = _PhysPkgReader(pkg_file)
        rels_xml = phys_reader.rels_xml_for(partname)
        phys_reader.close()
        return rels_xml

    @lazyproperty
    def _content_types(self):
        """Filty temporary hack during refactoring."""
        phys_reader = _PhysPkgReader(self._pkg_file)
        return _ContentTypeMap.from_xml(phys_reader.content_types_xml)
        phys_reader.close()

    @lazyproperty
    def _pkg_srels(self):
        """Filty temporary hack during refactoring."""
        phys_reader = _PhysPkgReader(self._pkg_file)
        pkg_srels = self._srels_for(phys_reader, PACKAGE_URI)
        phys_reader.close()
        return pkg_srels

    @lazyproperty
    def _sparts(self):
        """Filty temporary hack during refactoring.

        Return a list of |_SerializedPart| instances corresponding to the
        parts in *phys_reader* accessible by walking the relationship graph
        starting with `pkg_srels`.
        """
        phys_reader = _PhysPkgReader(self._pkg_file)
        sparts = []
        part_walker = self._walk_phys_parts(phys_reader, self._pkg_srels)
        for partname, blob, srels in part_walker:
            content_type = self._content_types[partname]
            spart = _SerializedPart(partname, content_type, blob, srels)
            sparts.append(spart)
        phys_reader.close()
        return tuple(sparts)

    def _srels_for(self, phys_reader, source_uri):
        """
        Return |_SerializedRelationshipCollection| instance populated with
        relationships for source identified by *source_uri*.
        """
        rels_xml = phys_reader.rels_xml_for(source_uri)
        return _SerializedRelationshipCollection.load_from_xml(
            source_uri.baseURI, rels_xml
        )

    def _walk_phys_parts(self, phys_reader, srels, visited_partnames=None):
        """
        Generate a 3-tuple `(partname, blob, srels)` for each of the parts in
        *phys_reader* by walking the relationship graph rooted at srels.
        """
        if visited_partnames is None:
            visited_partnames = []
        for srel in srels:
            if srel.is_external:
                continue
            partname = srel.target_partname
            if partname in visited_partnames:
                continue
            visited_partnames.append(partname)
            part_srels = self._srels_for(phys_reader, partname)
            blob = phys_reader.blob_for(partname)
            yield (partname, blob, part_srels)
            for partname, blob, srels in self._walk_phys_parts(
                phys_reader, part_srels, visited_partnames
            ):
                yield (partname, blob, srels)


class _ContentTypeMap(object):
    """
    Value type providing dictionary semantics for looking up content type by
    part name, e.g. ``content_type = cti['/ppt/presentation.xml']``.
    """

    def __init__(self):
        super(_ContentTypeMap, self).__init__()
        self._overrides = CaseInsensitiveDict()
        self._defaults = CaseInsensitiveDict()

    def __getitem__(self, partname):
        """
        Return content type for part identified by *partname*.
        """
        if not isinstance(partname, PackURI):
            tmpl = "_ContentTypeMap key must be <type 'PackURI'>, got %s"
            raise KeyError(tmpl % type(partname))
        if partname in self._overrides:
            return self._overrides[partname]
        if partname.ext in self._defaults:
            return self._defaults[partname.ext]
        tmpl = "no content type for partname '%s' in [Content_Types].xml"
        raise KeyError(tmpl % partname)

    @staticmethod
    def from_xml(content_types_xml):
        """
        Return a new |_ContentTypeMap| instance populated with the contents
        of *content_types_xml*.
        """
        types_elm = parse_xml(content_types_xml)
        ct_map = _ContentTypeMap()
        for o in types_elm.override_lst:
            ct_map._add_override(o.partName, o.contentType)
        for d in types_elm.default_lst:
            ct_map._add_default(d.extension, d.contentType)
        return ct_map

    def _add_default(self, extension, content_type):
        """
        Add the default mapping of *extension* to *content_type* to this
        content type mapping. *extension* does not include the leading
        period.
        """
        self._defaults[extension] = content_type

    def _add_override(self, partname, content_type):
        """
        Add the default mapping of *partname* to *content_type* to this
        content type mapping.
        """
        self._overrides[partname] = content_type


class PackageWriter(object):
    """
    Writes a zip-format OPC package to *pkg_file*, where *pkg_file* can be
    either a path to a zip file (a string) or a file-like object. Its single
    API method, :meth:`write`, is static, so this class is not intended to
    be instantiated.
    """

    @staticmethod
    def write(pkg_file, pkg_rels, parts):
        """
        Write a physical package (.pptx file) to *pkg_file* containing
        *pkg_rels* and *parts* and a content types stream based on the
        content types of the parts.
        """
        phys_writer = _PhysPkgWriter(pkg_file)
        PackageWriter._write_content_types_stream(phys_writer, parts)
        PackageWriter._write_pkg_rels(phys_writer, pkg_rels)
        PackageWriter._write_parts(phys_writer, parts)
        phys_writer.close()

    @staticmethod
    def _write_content_types_stream(phys_writer, parts):
        """
        Write ``[Content_Types].xml`` part to the physical package with an
        appropriate content type lookup target for each part in *parts*.
        """
        content_types_blob = serialize_part_xml(_ContentTypesItem.xml_for(parts))
        phys_writer.write(CONTENT_TYPES_URI, content_types_blob)

    @staticmethod
    def _write_parts(phys_writer, parts):
        """
        Write the blob of each part in *parts* to the package, along with a
        rels item for its relationships if and only if it has any.
        """
        for part in parts:
            phys_writer.write(part.partname, part.blob)
            if len(part._rels):
                phys_writer.write(part.partname.rels_uri, part._rels.xml)

    @staticmethod
    def _write_pkg_rels(phys_writer, pkg_rels):
        """
        Write the XML rels item for *pkg_rels* ('/_rels/.rels') to the
        package.
        """
        phys_writer.write(PACKAGE_URI.rels_uri, pkg_rels.xml)


class _PhysPkgReader(object):
    """
    Factory for physical package reader objects.
    """

    def __new__(cls, pkg_file):
        # if *pkg_file* is a string, treat it as a path
        if is_string(pkg_file):
            if os.path.isdir(pkg_file):
                reader_cls = _DirPkgReader
            elif zipfile.is_zipfile(pkg_file):
                reader_cls = _ZipPkgReader
            else:
                raise PackageNotFoundError("Package not found at '%s'" % pkg_file)
        else:  # assume it's a stream and pass it to Zip reader to sort out
            reader_cls = _ZipPkgReader

        return super(_PhysPkgReader, cls).__new__(reader_cls)


class _DirPkgReader(_PhysPkgReader):
    """
    Implements |PhysPkgReader| interface for an OPC package extracted into a
    directory.
    """

    def __init__(self, path):
        """
        *path* is the path to a directory containing an expanded package.
        """
        super(_DirPkgReader, self).__init__()
        self._path = os.path.abspath(path)

    def blob_for(self, pack_uri):
        """
        Return contents of file corresponding to *pack_uri* in package
        directory.
        """
        path = os.path.join(self._path, pack_uri.membername)
        with open(path, "rb") as f:
            blob = f.read()
        return blob

    def close(self):
        """
        Provides interface consistency with |ZipFileSystem|, but does
        nothing, a directory file system doesn't need closing.
        """
        pass

    @property
    def content_types_xml(self):
        """
        Return the `[Content_Types].xml` blob from the package.
        """
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """
        Return rels item XML for source with *source_uri*, or None if the
        item has no rels item.
        """
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except IOError:
            rels_xml = None
        return rels_xml


class _ZipPkgReader(_PhysPkgReader):
    """
    Implements |PhysPkgReader| interface for a zip file OPC package.
    """

    def __init__(self, pkg_file):
        super(_ZipPkgReader, self).__init__()
        self._zipf = zipfile.ZipFile(pkg_file, "r")

    def blob_for(self, pack_uri):
        """
        Return blob corresponding to *pack_uri*. Raises |ValueError| if no
        matching member is present in zip archive.
        """
        return self._zipf.read(pack_uri.membername)

    def close(self):
        """
        Close the zip archive, releasing any resources it is using.
        """
        self._zipf.close()

    @property
    def content_types_xml(self):
        """
        Return the `[Content_Types].xml` blob from the zip package.
        """
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """
        Return rels item XML for source with *source_uri* or None if no rels
        item is present.
        """
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except KeyError:
            rels_xml = None
        return rels_xml


class _PhysPkgWriter(object):
    """
    Factory for physical package writer objects.
    """

    def __new__(cls, pkg_file):
        return super(_PhysPkgWriter, cls).__new__(_ZipPkgWriter)


class _ZipPkgWriter(_PhysPkgWriter):
    """
    Implements |PhysPkgWriter| interface for a zip file OPC package.
    """

    def __init__(self, pkg_file):
        super(_ZipPkgWriter, self).__init__()
        self._zipf = zipfile.ZipFile(pkg_file, "w", compression=zipfile.ZIP_DEFLATED)

    def close(self):
        """
        Close the zip archive, flushing any pending physical writes and
        releasing any resources it's using.
        """
        self._zipf.close()

    def write(self, pack_uri, blob):
        """
        Write *blob* to this zip package with the membername corresponding to
        *pack_uri*.
        """
        self._zipf.writestr(pack_uri.membername, blob)


class _SerializedPart(object):
    """
    Value object for an OPC package part. Provides access to the partname,
    content type, blob, and serialized relationships for the part.
    """

    def __init__(self, partname, content_type, blob, srels):
        super(_SerializedPart, self).__init__()
        self._partname = partname
        self._content_type = content_type
        self._blob = blob
        self._srels = srels

    @property
    def partname(self):
        return self._partname

    @property
    def content_type(self):
        return self._content_type

    @property
    def blob(self):
        return self._blob

    @property
    def srels(self):
        return self._srels


class _SerializedRelationship(object):
    """
    Value object representing a serialized relationship in an OPC package.
    Serialized, in this case, means any target part is referred to via its
    partname rather than a direct link to an in-memory |Part| object.
    """

    def __init__(self, baseURI, rel_elm):
        super(_SerializedRelationship, self).__init__()
        self._baseURI = baseURI
        self._rId = rel_elm.rId
        self._reltype = rel_elm.reltype
        self._target_mode = rel_elm.targetMode
        self._target_ref = rel_elm.target_ref

    @property
    def is_external(self):
        """
        True if target_mode is ``RTM.EXTERNAL``
        """
        return self._target_mode == RTM.EXTERNAL

    @property
    def reltype(self):
        """Relationship type, like ``RT.OFFICE_DOCUMENT``"""
        return self._reltype

    @property
    def rId(self):
        """
        Relationship id, like 'rId9', corresponds to the ``Id`` attribute on
        the ``CT_Relationship`` element.
        """
        return self._rId

    @property
    def target_mode(self):
        """
        String in ``TargetMode`` attribute of ``CT_Relationship`` element,
        one of ``RTM.INTERNAL`` or ``RTM.EXTERNAL``.
        """
        return self._target_mode

    @property
    def target_ref(self):
        """
        String in ``Target`` attribute of ``CT_Relationship`` element, a
        relative part reference for internal target mode or an arbitrary URI,
        e.g. an HTTP URL, for external target mode.
        """
        return self._target_ref

    @property
    def target_partname(self):
        """
        |PackURI| instance containing partname targeted by this relationship.
        Raises ``ValueError`` on reference if target_mode is ``'External'``.
        Use :attr:`target_mode` to check before referencing.
        """
        if self.is_external:
            msg = (
                "target_partname attribute on Relationship is undefined w"
                'here TargetMode == "External"'
            )
            raise ValueError(msg)
        # lazy-load _target_partname attribute
        if not hasattr(self, "_target_partname"):
            self._target_partname = PackURI.from_rel_ref(self._baseURI, self.target_ref)
        return self._target_partname


class _SerializedRelationshipCollection(object):
    """
    Read-only sequence of |_SerializedRelationship| instances corresponding
    to the relationships item XML passed to constructor.
    """

    def __init__(self):
        super(_SerializedRelationshipCollection, self).__init__()
        self._srels = []

    def __iter__(self):
        """Support iteration, e.g. 'for x in srels:'"""
        return self._srels.__iter__()

    @staticmethod
    def load_from_xml(baseURI, rels_item_xml):
        """
        Return |_SerializedRelationshipCollection| instance loaded with the
        relationships contained in *rels_item_xml*. Returns an empty
        collection if *rels_item_xml* is |None|.
        """
        srels = _SerializedRelationshipCollection()
        if rels_item_xml is not None:
            rels_elm = parse_xml(rels_item_xml)
            for rel_elm in rels_elm.relationship_lst:
                srels._srels.append(_SerializedRelationship(baseURI, rel_elm))
        return srels


class _ContentTypesItem(object):
    """
    Service class that composes a content types item ([Content_Types].xml)
    based on a list of parts. Not meant to be instantiated directly, its
    single interface method is xml_for(), e.g.
    ``_ContentTypesItem.xml_for(parts)``.
    """

    def __init__(self):
        self._defaults = CaseInsensitiveDict()
        self._overrides = dict()

    @classmethod
    def xml_for(cls, parts):
        """
        Return content types XML mapping each part in *parts* to the
        appropriate content type and suitable for storage as
        ``[Content_Types].xml`` in an OPC package.
        """
        cti = cls()
        cti._defaults["rels"] = CT.OPC_RELATIONSHIPS
        cti._defaults["xml"] = CT.XML
        for part in parts:
            cti._add_content_type(part.partname, part.content_type)
        return cti._xml()

    def _add_content_type(self, partname, content_type):
        """
        Add a content type for the part with *partname* and *content_type*,
        using a default or override as appropriate.
        """
        ext = partname.ext
        if (ext.lower(), content_type) in default_content_types:
            self._defaults[ext] = content_type
        else:
            self._overrides[partname] = content_type

    def _xml(self):
        """
        Return etree element containing the XML representation of this content
        types item, suitable for serialization to the ``[Content_Types].xml``
        item for an OPC package. Although the sequence of elements is not
        strictly significant, as an aid to testing and readability Default
        elements are sorted by extension and Override elements are sorted by
        partname.
        """
        _types_elm = CT_Types.new()
        for ext in sorted(self._defaults.keys()):
            _types_elm.add_default(ext, self._defaults[ext])
        for partname in sorted(self._overrides.keys()):
            _types_elm.add_override(partname, self._overrides[partname])
        return _types_elm
