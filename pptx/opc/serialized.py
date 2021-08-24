# encoding: utf-8

"""API for reading/writing serialized Open Packaging Convention (OPC) package."""

import os
import zipfile

from pptx.compat import Container, is_string
from pptx.exceptions import PackageNotFoundError
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.oxml import CT_Types, serialize_part_xml
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

    def __getitem__(self, pack_uri):
        """Return bytes for part corresponding to `pack_uri`."""
        return self._blob_reader[pack_uri]

    def rels_xml_for(self, partname):
        """Return optional rels item XML for `partname`.

        Returns `None` if no rels item is present for `partname`. `partname` is a
        |PackURI| instance.
        """
        blob_reader, uri = self._blob_reader, partname.rels_uri
        return blob_reader[uri] if uri in blob_reader else None

    @lazyproperty
    def _blob_reader(self):
        """|_PhysPkgReader| subtype providing read access to the package file."""
        return _PhysPkgReader.factory(self._pkg_file)


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


class _PhysPkgReader(Container):
    """Base class for physical package reader objects."""

    def __contains__(self, item):
        """Must be implemented by each subclass."""
        raise NotImplementedError(
            "`%s` must implement `.__contains__()`" % type(self).__name__
        )

    @classmethod
    def factory(cls, pkg_file):
        """Return |_PhysPkgReader| subtype instance appropriage for `pkg_file`."""
        # --- for pkg_file other than str, assume it's a stream and pass it to Zip
        # --- reader to sort out
        if not is_string(pkg_file):
            return _ZipPkgReader(pkg_file)

        # --- otherwise we treat `pkg_file` as a path ---
        if os.path.isdir(pkg_file):
            return _DirPkgReader(pkg_file)

        if zipfile.is_zipfile(pkg_file):
            return _ZipPkgReader(pkg_file)

        raise PackageNotFoundError("Package not found at '%s'" % pkg_file)


class _DirPkgReader(_PhysPkgReader):
    """Implements |PhysPkgReader| interface for OPC package extracted into directory.

    `path` is the path to a directory containing an expanded package.
    """

    def __init__(self, path):
        self._path = os.path.abspath(path)

    def __getitem__(self, pack_uri):
        """Return bytes of file corresponding to `pack_uri` in package directory."""
        path = os.path.join(self._path, pack_uri.membername)
        try:
            with open(path, "rb") as f:
                return f.read()
        except IOError:
            raise KeyError("no member '%s' in package" % pack_uri)


class _ZipPkgReader(_PhysPkgReader):
    """Implements |PhysPkgReader| interface for a zip-file OPC package."""

    def __init__(self, pkg_file):
        self._pkg_file = pkg_file

    def __contains__(self, pack_uri):
        """Return True when part identified by `pack_uri` is present in zip archive."""
        return pack_uri in self._blobs

    def __getitem__(self, pack_uri):
        """Return bytes for part corresponding to `pack_uri`.

        Raises |KeyError| if no matching member is present in zip archive.
        """
        if pack_uri not in self._blobs:
            raise KeyError("no member '%s' in package" % pack_uri)
        return self._blobs[pack_uri]

    @lazyproperty
    def _blobs(self):
        """dict mapping partname to package part binaries."""
        with zipfile.ZipFile(self._pkg_file, "r") as z:
            return {PackURI("/%s" % name): z.read(name) for name in z.namelist()}


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
