"""API for reading/writing serialized Open Packaging Convention (OPC) package."""

from __future__ import annotations

import datetime as dt
import io
import os
import posixpath
import zipfile
from typing import IO, TYPE_CHECKING, Any, Container, Sequence, Tuple, Union

from pptx.exc import PackageNotFoundError, PackageTooLargeError
from pptx.opc._crypto import (
    decrypt_stream,
    encrypt_bytes,
    is_encrypted_stream,
    is_rms_protected_stream,
)
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.oxml import CT_Types, serialize_part_xml
from pptx.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI, PackURI
from pptx.opc.shared import CaseInsensitiveDict
from pptx.opc.spec import default_content_types
from pptx.util import lazyproperty


def _default_max_uncompressed_size() -> int:
    """Return default zip-bomb guard threshold.

    Honors the ``PPTX_MAX_UNCOMPRESSED_SIZE`` environment variable when it
    parses as a positive integer, otherwise defaults to 2 GiB. Setting the
    value to ``0`` disables the guard.
    """
    raw = os.environ.get("PPTX_MAX_UNCOMPRESSED_SIZE")
    if raw is not None:
        try:
            value = int(raw)
            if value >= 0:
                return value
        except ValueError:
            pass
    return 2 * 1024 * 1024 * 1024  # 2 GiB


# -- Zip-bomb guard threshold. Packages whose central-directory entries declare
# -- a combined uncompressed size greater than this value are rejected before
# -- any member bytes are loaded into memory. Set to ``0`` to disable the
# -- guard (not recommended for untrusted inputs). The module attribute is
# -- intentionally mutable so applications that need to process very large
# -- legitimate packages can raise the limit without subclassing.
MAX_UNCOMPRESSED_PACKAGE_SIZE: int = _default_max_uncompressed_size()

if TYPE_CHECKING:
    from pptx.opc.package import Part, _Relationships  # pyright: ignore[reportPrivateUsage]

ZipDateTime = Union[dt.datetime, Tuple[int, int, int, int, int, int]]
"""Value accepted for a fixed zip-member timestamp.

Either a |datetime.datetime| (tzinfo is ignored, only the naive wall-clock
components are used) or a 6-tuple ``(year, month, day, hour, minute, second)``.
Values outside the Zip range (1980-01-01 through 2107-12-31) will be rejected
by the underlying :class:`zipfile.ZipInfo`.
"""


class PackageReader(Container[bytes]):
    """Provides access to package-parts of an OPC package with dict semantics.

    The package may be in zip-format (a .pptx file) or expanded into a directory structure,
    perhaps by unzipping a .pptx file. When `password` is provided, a password-protected
    .pptx file is decrypted before reading; this requires the optional
    ``python-ooxml-crypto`` dependency.
    """

    def __init__(self, pkg_file: str | IO[bytes], password: str | None = None):
        self._pkg_file = pkg_file
        self._password = password

    def __contains__(self, pack_uri: object) -> bool:
        """Return True when part identified by `pack_uri` is present in package."""
        return pack_uri in self._blob_reader

    def __getitem__(self, pack_uri: PackURI) -> bytes:
        """Return bytes for part corresponding to `pack_uri`."""
        return self._blob_reader[pack_uri]

    def rels_xml_for(self, partname: PackURI) -> bytes | None:
        """Return optional rels item XML for `partname`.

        Returns `None` if no rels item is present for `partname`. `partname` is a |PackURI|
        instance.
        """
        blob_reader, uri = self._blob_reader, partname.rels_uri
        # -- _PhysPkgReader implements only `__contains__` / `__getitem__`, not
        # -- `.get()`, so we can't substitute a dict-style `blob_reader.get(uri, None)`.
        return blob_reader[uri] if uri in blob_reader else None  # noqa: SIM401

    @lazyproperty
    def _blob_reader(self) -> _PhysPkgReader:
        """|_PhysPkgReader| subtype providing read access to the package file."""
        return _PhysPkgReader.factory(self._pkg_file, self._password)


class PackageWriter:
    """Writes a zip-format OPC package to `pkg_file`.

    `pkg_file` can be either a path to a zip file (a string) or a file-like object. `pkg_rels` is
    the |_Relationships| object containing relationships for the package. `parts` is a sequence of
    |Part| subtype instance to be written to the package. When `password` is provided the
    package is encrypted using ECMA-376 Agile Encryption (requires the optional
    ``python-ooxml-crypto`` dependency).

    Its single API classmethod is :meth:`write`. This class is not intended to be instantiated.
    """

    def __init__(
        self,
        pkg_file: str | IO[bytes],
        pkg_rels: _Relationships,
        parts: Sequence[Part],
        zip_date_time: ZipDateTime | None = None,
        password: str | None = None,
    ):
        self._pkg_file = pkg_file
        self._pkg_rels = pkg_rels
        self._parts = parts
        self._zip_date_time = zip_date_time
        self._password = password

    @classmethod
    def write(
        cls,
        pkg_file: str | IO[bytes],
        pkg_rels: _Relationships,
        parts: Sequence[Part],
        zip_date_time: ZipDateTime | None = None,
        password: str | None = None,
    ) -> None:
        """Write a physical package (.pptx file) to `pkg_file`.

        The serialized package contains `pkg_rels` and `parts`, a content-types stream based on
        the content type of each part, and a .rels file for each part that has relationships.

        When `zip_date_time` is provided, every zip-member's last-modified timestamp is set to
        that value, producing byte-identical output across builds with otherwise-identical
        inputs (useful for reproducible builds and source-control friendliness).

        When `password` is provided the resulting .pptx is password-protected using ECMA-376
        Agile Encryption. The two keywords are orthogonal: `zip_date_time` stamps the inner
        (plaintext) zip members before encryption wraps the package in an OLE2 CFBF container.
        """
        cls(pkg_file, pkg_rels, parts, zip_date_time, password)._write()

    def _write(self) -> None:
        """Write physical package (.pptx file).

        When a password is configured, the plaintext zip is first built into an in-memory
        buffer (still honoring ``zip_date_time`` for its members) and then encrypted and
        written to the real destination.
        """
        if self._password is None:
            with _PhysPkgWriter.factory(self._pkg_file, self._zip_date_time) as phys_writer:
                self._write_content_types_stream(phys_writer)
                self._write_pkg_rels(phys_writer)
                self._write_parts(phys_writer)
            return

        # -- build the plain zip into an in-memory buffer, then encrypt it --
        plain_buffer = io.BytesIO()
        with _PhysPkgWriter.factory(plain_buffer, self._zip_date_time) as phys_writer:
            self._write_content_types_stream(phys_writer)
            self._write_pkg_rels(phys_writer)
            self._write_parts(phys_writer)

        encrypted = encrypt_bytes(plain_buffer.getvalue(), self._password)

        if isinstance(self._pkg_file, str):
            with open(self._pkg_file, "wb") as f:
                f.write(encrypted)
        else:
            self._pkg_file.write(encrypted)

    def _write_content_types_stream(self, phys_writer: _PhysPkgWriter) -> None:
        """Write `[Content_Types].xml` part to the physical package.

        This part must contain an appropriate content type lookup target for each part in the
        package.
        """
        phys_writer.write(
            CONTENT_TYPES_URI,
            serialize_part_xml(_ContentTypesItem.xml_for(self._parts)),
        )

    def _write_parts(self, phys_writer: _PhysPkgWriter) -> None:
        """Write blob of each part in `parts` to the package.

        A rels item for each part is written when the part has relationships, or when an
        (empty) rels file was physically present for this part in the input package — in
        which case it is preserved verbatim so byte-level round-trips stay faithful.
        """
        for part in self._parts:
            phys_writer.write(part.partname, part.blob)
            has_rels = bool(part._rels)  # pyright: ignore[reportPrivateUsage]
            preserved_empty = getattr(part, "_rels_file_present_on_load", False)
            if has_rels or preserved_empty:
                phys_writer.write(part.partname.rels_uri, part.rels.xml)

    def _write_pkg_rels(self, phys_writer: _PhysPkgWriter) -> None:
        """Write the XML rels item for `pkg_rels` ('/_rels/.rels') to the package."""
        phys_writer.write(PACKAGE_URI.rels_uri, self._pkg_rels.xml)


class _PhysPkgReader(Container[PackURI]):
    """Base class for physical package reader objects."""

    def __contains__(self, item: object) -> bool:
        """Must be implemented by each subclass."""
        raise NotImplementedError(  # pragma: no cover
            "`%s` must implement `.__contains__()`" % type(self).__name__
        )

    def __getitem__(self, pack_uri: PackURI) -> bytes:
        """Blob for part corresponding to `pack_uri`."""
        raise NotImplementedError(  # pragma: no cover
            f"`{type(self).__name__}` must implement `.__contains__()`"
        )

    @classmethod
    def factory(cls, pkg_file: str | IO[bytes], password: str | None = None) -> _PhysPkgReader:
        """Return |_PhysPkgReader| subtype instance appropriage for `pkg_file`.

        When `pkg_file` is an encrypted OOXML package (CFBF / OLE2 container), the file
        is decrypted with `password` and the resulting plaintext bytes are read as a
        zip package. Decryption requires the optional ``python-ooxml-crypto`` dependency.
        """
        # --- for pkg_file other than str, assume it's a stream and pass it to Zip
        # --- reader to sort out
        if not isinstance(pkg_file, str):
            decrypted = cls._maybe_decrypt_stream(pkg_file, password)
            return _ZipPkgReader(decrypted if decrypted is not None else pkg_file)

        # --- otherwise we treat `pkg_file` as a path ---
        if os.path.isdir(pkg_file):
            return _DirPkgReader(pkg_file)

        if zipfile.is_zipfile(pkg_file):
            return _ZipPkgReader(pkg_file)

        # --- an encrypted .pptx is a CFBF container, not a zip, so check the file's
        # --- first bytes before giving up. An absent or unreadable path falls through
        # --- to the PackageNotFoundError below.
        if os.path.isfile(pkg_file):
            with open(pkg_file, "rb") as f:
                decrypted = cls._maybe_decrypt_stream(f, password)
            if decrypted is not None:
                return _ZipPkgReader(decrypted)

        raise PackageNotFoundError("Package not found at '%s'" % pkg_file)

    @staticmethod
    def _maybe_decrypt_stream(stream: IO[bytes], password: str | None) -> io.BytesIO | None:
        """Return decrypted-payload BytesIO if `stream` is encrypted, else None.

        Raises :class:`pptx.exc.EncryptedPackageError` when the stream is encrypted and
        either `password` is None or ``python-ooxml-crypto`` is not installed.

        Raises :class:`pptx.exc.RmsProtectedPackageError` when the stream is wrapped in
        Azure RMS / AIP / IRM protection. RMS-protected packages are CFBF containers
        whose payload is encrypted to the user's Azure AD identity rather than a
        password; python-pptx cannot decrypt them (see the ``rms-protected`` section of
        the user guide for workarounds).
        """
        if not is_encrypted_stream(stream):
            return None

        from pptx.exc import EncryptedPackageError, RmsProtectedPackageError

        if is_rms_protected_stream(stream):
            raise RmsProtectedPackageError(
                "package is wrapped in Azure RMS / AIP / IRM protection; python-pptx "
                "cannot decrypt RMS-protected files — the payload is encrypted to the "
                "user's Microsoft 365 identity, not a password. See the "
                "'RMS / AIP-protected files' section of the python-pptx user guide "
                "for recommended workarounds (delegating decryption to Microsoft "
                "Office automation or the Microsoft Information Protection SDK)."
            )

        if password is None:
            raise EncryptedPackageError(
                "package is password-protected; supply `password` to open it"
            )

        plain = decrypt_stream(stream, password)
        return io.BytesIO(plain)


class _DirPkgReader(_PhysPkgReader):
    """Implements |PhysPkgReader| interface for OPC package extracted into directory.

    `path` is the path to a directory containing an expanded package.
    """

    def __init__(self, path: str):
        self._path = os.path.abspath(path)

    def __contains__(self, pack_uri: object) -> bool:
        """Return True when part identified by `pack_uri` is present in zip archive."""
        if not isinstance(pack_uri, PackURI):
            return False
        return os.path.exists(posixpath.join(self._path, pack_uri.membername))

    def __getitem__(self, pack_uri: PackURI) -> bytes:
        """Return bytes of file corresponding to `pack_uri` in package directory."""
        path = os.path.join(self._path, pack_uri.membername)
        try:
            with open(path, "rb") as f:
                return f.read()
        except IOError:
            raise KeyError("no member '%s' in package" % pack_uri)


class _ZipPkgReader(_PhysPkgReader):
    """Implements |PhysPkgReader| interface for a zip-file OPC package."""

    def __init__(self, pkg_file: str | IO[bytes]):
        self._pkg_file = pkg_file

    def __contains__(self, pack_uri: object) -> bool:
        """Return True when part identified by `pack_uri` is present in zip archive."""
        return pack_uri in self._blobs

    def __getitem__(self, pack_uri: PackURI) -> bytes:
        """Return bytes for part corresponding to `pack_uri`.

        Raises |KeyError| if no matching member is present in zip archive.
        """
        if pack_uri not in self._blobs:
            raise KeyError("no member '%s' in package" % pack_uri)
        return self._blobs[pack_uri]

    @lazyproperty
    def _blobs(self) -> dict[PackURI, bytes]:
        """dict mapping partname to package part binaries."""
        with zipfile.ZipFile(self._pkg_file, "r") as z:
            self._check_uncompressed_size(z)
            return {PackURI("/%s" % name): z.read(name) for name in z.namelist()}

    @staticmethod
    def _check_uncompressed_size(zf: zipfile.ZipFile) -> None:
        """Raise |PackageTooLargeError| when zip declares too-large uncompressed size.

        The guard compares the sum of per-member ``file_size`` values recorded in
        the zip central directory against |MAX_UNCOMPRESSED_PACKAGE_SIZE|. It is a
        lightweight defense against "zip-bomb" packages whose entries declare an
        enormous uncompressed size; rejecting them before reading prevents the
        reader from attempting to materialize gigabytes of member bytes in memory.
        """
        limit = MAX_UNCOMPRESSED_PACKAGE_SIZE
        if limit <= 0:
            return
        total = 0
        for zinfo in zf.infolist():
            # Per-member sanity check catches single enormous members even when
            # `file_size` is a plausibly-normal int; the guard is a partial
            # defense, not a guarantee, and complements standard OS resource
            # limits (see `docs/dev/security.rst`).
            total += zinfo.file_size
            if total > limit:
                raise PackageTooLargeError(
                    "package declared uncompressed size exceeds limit "
                    "of %d bytes (set PPTX_MAX_UNCOMPRESSED_SIZE or "
                    "pptx.opc.serialized.MAX_UNCOMPRESSED_PACKAGE_SIZE to override)"
                    % limit
                )


class _PhysPkgWriter:
    """Base class for physical package writer objects."""

    @classmethod
    def factory(
        cls, pkg_file: str | IO[bytes], zip_date_time: ZipDateTime | None = None
    ) -> _ZipPkgWriter:
        """Return |_PhysPkgWriter| subtype instance appropriage for `pkg_file`.

        Currently the only subtype is `_ZipPkgWriter`, but a `_DirPkgWriter` could be implemented
        or even a `_StreamPkgWriter`.
        """
        return _ZipPkgWriter(pkg_file, zip_date_time)

    def write(self, pack_uri: PackURI, blob: bytes) -> None:
        """Write `blob` to package with membername corresponding to `pack_uri`."""
        raise NotImplementedError(  # pragma: no cover
            f"`{type(self).__name__}` must implement `.write()`"
        )


class _ZipPkgWriter(_PhysPkgWriter):
    """Implements |PhysPkgWriter| interface for a zip-file (.pptx file) OPC package.

    When `zip_date_time` is provided, every zip-member written by this writer is stamped with
    that fixed last-modified date-time, yielding byte-identical output across repeated writes
    (source-control friendly / reproducible builds).
    """

    def __init__(self, pkg_file: str | IO[bytes], zip_date_time: ZipDateTime | None = None):
        self._pkg_file = pkg_file
        self._zip_date_time = _normalize_zip_date_time(zip_date_time)

    def __enter__(self) -> _ZipPkgWriter:
        """Enable use as a context-manager. Opening zip for writing happens here."""
        return self

    def __exit__(self, *exc: list[Any]) -> None:
        """Close the zip archive on exit from context.

        Closing flushes any pending physical writes and releasing any resources it's using.
        """
        self._zipf.close()

    def write(self, pack_uri: PackURI, blob: bytes) -> None:
        """Write `blob` to zip package with membername corresponding to `pack_uri`."""
        if self._zip_date_time is None:
            self._zipf.writestr(pack_uri.membername, blob)
            return
        zinfo = zipfile.ZipInfo(pack_uri.membername, date_time=self._zip_date_time)
        # --- preserve the archive's compression settings; `writestr(name, ...)` picks those up
        # --- from the ZipFile automatically but `writestr(ZipInfo, ...)` does not.
        zinfo.compress_type = self._zipf.compression
        self._zipf.writestr(zinfo, blob, compresslevel=self._zipf.compresslevel)

    @lazyproperty
    def _zipf(self) -> zipfile.ZipFile:
        """`ZipFile` instance open for writing."""
        return zipfile.ZipFile(
            self._pkg_file, "w", compression=zipfile.ZIP_DEFLATED, strict_timestamps=False
        )


def _normalize_zip_date_time(
    zip_date_time: ZipDateTime | None,
) -> tuple[int, int, int, int, int, int] | None:
    """Return `zip_date_time` coerced to the 6-tuple form ZipInfo expects, or `None`."""
    if zip_date_time is None:
        return None
    if isinstance(zip_date_time, dt.datetime):
        return (
            zip_date_time.year,
            zip_date_time.month,
            zip_date_time.day,
            zip_date_time.hour,
            zip_date_time.minute,
            zip_date_time.second,
        )
    # --- accept any 6-length sequence of ints; defend against malformed input at runtime
    t = tuple(zip_date_time)
    if len(t) != 6 or not all(
        isinstance(v, int) for v in t  # pyright: ignore[reportUnnecessaryIsInstance]
    ):
        raise TypeError(
            "zip_date_time must be a datetime or a 6-tuple of ints "
            "(year, month, day, hour, minute, second)"
        )
    return (t[0], t[1], t[2], t[3], t[4], t[5])


class _ContentTypesItem:
    """Composes content-types "part" ([Content_Types].xml) for a collection of parts."""

    def __init__(self, parts: Sequence[Part]):
        self._parts = parts

    @classmethod
    def xml_for(cls, parts: Sequence[Part]) -> CT_Types:
        """Return content-types XML mapping each part in `parts` to a content-type.

        The resulting XML is suitable for storage as `[Content_Types].xml` in an OPC package.
        """
        return cls(parts)._xml

    @lazyproperty
    def _xml(self) -> CT_Types:
        """lxml.etree._Element containing the content-types item.

        This XML object is suitable for serialization to the `[Content_Types].xml` item for an OPC
        package. Although the sequence of elements is not strictly significant, as an aid to
        testing and readability Default elements are sorted by extension and Override elements are
        sorted by partname.
        """
        defaults, overrides = self._defaults_and_overrides
        _types_elm = CT_Types.new()

        for ext, content_type in sorted(defaults.items()):
            _types_elm.add_default(ext, content_type)
        for partname, content_type in sorted(overrides.items()):
            _types_elm.add_override(partname, content_type)

        return _types_elm

    @lazyproperty
    def _defaults_and_overrides(self) -> tuple[dict[str, str], dict[PackURI, str]]:
        """pair of dict (defaults, overrides) accounting for all parts.

        `defaults` is {ext: content_type} and overrides is {partname: content_type}.
        """
        defaults = CaseInsensitiveDict(rels=CT.OPC_RELATIONSHIPS, xml=CT.XML)
        overrides: dict[PackURI, str] = {}

        for part in self._parts:
            partname, content_type = part.partname, part.content_type
            ext = partname.ext
            if (ext.lower(), content_type) in default_content_types:
                defaults[ext] = content_type
            else:
                overrides[partname] = content_type

        return defaults, overrides
