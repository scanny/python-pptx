"""Overall .pptx package."""

from __future__ import annotations

import hashlib
import io
import os
import re
from typing import IO, Iterator, cast

from pptx.exc import UnsupportedImageTypeError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import OpcPackage
from pptx.opc.packuri import PackURI
from pptx.parts.comments import CommentAuthorsPart
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.customprops import CustomProperties, CustomPropertiesPart
from pptx.parts.extprops import ExtendedPropertiesPart
from pptx.parts.image import Image, ImagePart
from pptx.parts.media import MediaPart
from pptx.parts.viewprops import ViewPropsPart
from pptx.util import lazyproperty

# -- Used to detect SVG content in an image byte-stream. Looks at roughly the
# -- first 2 KiB (enough to skip common XML prologues / comments / DOCTYPE) for
# -- an `<svg` root-element opener. This is a lightweight sniff — the caller is
# -- expected to hand over an image blob, not an arbitrary text file.
_SVG_HEAD_BYTES = 2048
# -- matches the opening of an ``<svg`` root element in any of its common
# -- forms: ``<svg>``, ``<svg ...>``, ``<svg\n...``, ``<svg/>``.
_SVG_SNIFF_RE = re.compile(rb"<svg[\s>/]", re.IGNORECASE)


class Package(OpcPackage):
    """An overall .pptx package."""

    @property
    def comment_authors_part(self) -> CommentAuthorsPart | None:
        """The package's legacy |CommentAuthorsPart|, or None if not present.

        The relationship lives on the presentation part per ECMA-376 §13.3.3.

        .. versionadded:: 2026.05.0
        """
        try:
            return cast(
                CommentAuthorsPart,
                self.presentation_part.part_related_by(RT.COMMENT_AUTHORS),
            )
        except KeyError:
            return None

    def get_or_add_comment_authors_part(self) -> CommentAuthorsPart:
        """Return this package's legacy comment-authors part, adding one if needed.

        The relationship is attached to the presentation part.

        .. versionadded:: 2026.05.0
        """
        part = self.comment_authors_part
        if part is not None:
            return part
        part = CommentAuthorsPart.new(self)
        self.presentation_part.relate_to(part, RT.COMMENT_AUTHORS)
        return part

    @lazyproperty
    def core_properties(self) -> CorePropertiesPart:
        """Instance of |CoreProperties| holding read/write Dublin Core doc properties.

        Creates a default core properties part if one is not present (not common).
        """
        try:
            return self.part_related_by(RT.CORE_PROPERTIES)
        except KeyError:
            core_props = CorePropertiesPart.default(self)
            self.relate_to(core_props, RT.CORE_PROPERTIES)
            return core_props

    @lazyproperty
    def custom_properties(self) -> CustomProperties:
        """Dict-like |CustomProperties| facade over the custom-properties part.

        See :class:`pptx.parts.customprops.CustomProperties` for the supported operations
        (``__getitem__`` / ``__setitem__`` / ``__delitem__`` / ``__contains__`` / ``keys``
        / ``items`` / ``update`` / ``clear`` / ``pop`` / ``setdefault``). Reads against a
        package that has no custom-properties part behave as though it were empty; writes
        materialize the part lazily on first use. See issue #259.

        .. versionadded:: 2026.05.0
        """
        return CustomProperties(self)

    @property
    def custom_properties_part(self) -> CustomPropertiesPart:
        """Instance of |CustomPropertiesPart| (``/docProps/custom.xml``).

        Creates a default (empty) custom-properties part if one is not present, wires up
        the relationship from the package root, and returns it. Most callers should use
        the higher-level :attr:`custom_properties` dict-facade instead of this part.

        .. versionadded:: 2026.05.0
        """
        try:
            return self.part_related_by(RT.CUSTOM_PROPERTIES)
        except KeyError:
            custom_props = CustomPropertiesPart.default(self)
            self.relate_to(custom_props, RT.CUSTOM_PROPERTIES)
            return custom_props

    @lazyproperty
    def extended_properties(self) -> ExtendedPropertiesPart:
        """Instance of |ExtendedPropertiesPart| (``/docProps/app.xml``).

        Creates a default extended-properties part if one is not present. This part holds
        application-level properties such as the slide count (which python-pptx updates at
        save-time so downstream consumers like Gmail's attachment preview render correctly).

        .. versionadded:: 2026.05.0
        """
        try:
            return self.part_related_by(RT.EXTENDED_PROPERTIES)
        except KeyError:
            ext_props = ExtendedPropertiesPart.default(self)
            self.relate_to(ext_props, RT.EXTENDED_PROPERTIES)
            return ext_props

    @lazyproperty
    def view_props_part(self) -> ViewPropsPart:
        """|ViewPropsPart| for this package (``ppt/viewProps.xml``).

        Editor-view settings restored by PowerPoint on open — which editor
        view was last active (``@lastView``), whether the comments pane is
        visible (``@showComments``), and per-view zoom / scroll-position
        state. See issue #94.

        Attached to the *presentation part* (``ppt/presentation.xml``) via
        a ``viewProps`` relationship, not to the package root. Creates a
        default (empty) part and wires up the relationship if one is not
        already present.

        .. versionadded:: 2026.05.0
        """
        try:
            return cast(
                ViewPropsPart,
                self.presentation_part.part_related_by(RT.VIEW_PROPS),
            )
        except KeyError:
            view_props = ViewPropsPart.default(self)
            self.presentation_part.relate_to(view_props, RT.VIEW_PROPS)
            return view_props

    def get_or_add_image_part(self, image_file: str | IO[bytes]):
        """
        Return an |ImagePart| object containing the image in *image_file*. If
        the image part already exists in this package, it is reused,
        otherwise a new one is created.
        """
        return self._image_parts.get_or_add_image_part(image_file)

    def get_or_add_svg_part(self, svg_file: str | IO[bytes]) -> ImagePart:
        """Return an |ImagePart| object containing the SVG in *svg_file*.

        The SVG bytes are hashed and deduplicated against other SVG image
        parts already in the package — reinserting the same SVG on multiple
        slides shares a single ``/ppt/media/imageN.svg`` part.

        Bypasses the ``_raise_if_svg`` sniff applied by
        :meth:`get_or_add_image_part`: here SVG content is the expected
        input, not a rejected special case. See issue #358.

        .. versionadded:: 2026.05.0
        """
        return self._image_parts.get_or_add_svg_part(svg_file)

    def get_or_add_media_part(self, media):
        """Return a |MediaPart| object containing the media in *media*.

        If a media part for this media bytestream ("file") is already present
        in this package, it is reused, otherwise a new one is created.
        """
        return self._media_parts.get_or_add_media_part(media)

    def next_image_partname(self, ext: str) -> PackURI:
        """Return a |PackURI| instance representing the next available image partname.

        Partname uses the next available sequence number. *ext* is used as the extention on the
        returned partname.
        """

        def first_available_image_idx():
            image_idxs = sorted(
                [
                    part.partname.idx
                    for part in self.iter_parts()
                    if (
                        part.partname.startswith("/ppt/media/image")
                        and part.partname.idx is not None
                    )
                ]
            )
            for i, image_idx in enumerate(image_idxs):
                idx = i + 1
                if idx < image_idx:
                    return idx
            return len(image_idxs) + 1

        idx = first_available_image_idx()
        return PackURI("/ppt/media/image%d.%s" % (idx, ext))

    def next_media_partname(self, ext):
        """Return |PackURI| instance for next available media partname.

        Partname is first available, starting at sequence number 1. Empty
        sequence numbers are reused. *ext* is used as the extension on the
        returned partname.
        """

        def first_available_media_idx():
            media_idxs = sorted(
                [
                    part.partname.idx
                    for part in self.iter_parts()
                    if part.partname.startswith("/ppt/media/media")
                ]
            )
            for i, media_idx in enumerate(media_idxs):
                idx = i + 1
                if idx < media_idx:
                    return idx
            return len(media_idxs) + 1

        idx = first_available_media_idx()
        return PackURI("/ppt/media/media%d.%s" % (idx, ext))

    @property
    def presentation_part(self):
        """
        Reference to the |Presentation| instance contained in this package.
        """
        return self.main_document_part

    def save(self, pkg_file, zip_date_time=None, password=None):
        """Save this package to `pkg_file`.

        Ensures the extended-properties part (``/docProps/app.xml``) exists and its
        ``<Slides>`` count is refreshed before serialization. See issue #131 — without this,
        Gmail's attachment preview (among others) fails to render.

        `zip_date_time` is forwarded to the base-class save (see issue #702) so reproducible
        timestamps continue to work when the Presentation-level Package override is active.
        `password` is forwarded likewise (see issue #668) so the package can be written
        password-protected. Both keywords may be combined.

        .. versionadded:: 2026.05.0
        """
        # -- trigger lazy creation of the part (no-op if already present) so that it is
        # -- included in the package walk performed by the base-class save.
        _ = self.extended_properties
        super().save(pkg_file, zip_date_time, password=password)

    @lazyproperty
    def _image_parts(self):
        """
        |_ImageParts| object providing access to the image parts in this
        package.
        """
        return _ImageParts(self)

    @lazyproperty
    def _media_parts(self):
        """Return |_MediaParts| object for this package.

        The media parts object provides access to all the media parts in this
        package.
        """
        return _MediaParts(self)


class _ImageParts(object):
    """Provides access to the image parts in a package."""

    def __init__(self, package):
        super(_ImageParts, self).__init__()
        self._package = package

    def __iter__(self) -> Iterator[ImagePart]:
        """Generate a reference to each |ImagePart| object in the package."""
        image_parts = []
        for rel in self._package.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype != RT.IMAGE:
                continue
            image_part = rel.target_part
            if image_part in image_parts:
                continue
            image_parts.append(image_part)
            yield image_part

    def get_or_add_image_part(self, image_file: str | IO[bytes]) -> ImagePart:
        """Return |ImagePart| object containing the image in `image_file`.

        `image_file` can be either a path to an image file or a file-like object
        containing an image. If an image part containing this same image already exists,
        that instance is returned, otherwise a new image part is created.

        Raises |UnsupportedImageTypeError| when `image_file` contains SVG content.
        PowerPoint's SVG support (Office 2016+) requires a companion PNG raster
        fallback paired with the SVG via an ``asvg:svgBlip`` extension, and
        python-pptx does not bundle an SVG rasterizer. Callers must pre-rasterize
        SVG content to PNG (or another supported raster format) before inserting
        it. See the "Inserting SVG images" section of the user guide.

        When `image_file` is a non-seekable file-like object (e.g. a stream
        backed by a blob URL or an HTTP upload reader), its full content is
        read into an in-memory buffer before the SVG sniff and PIL-backed
        format detection run. This ensures format detection succeeds even
        when the stream cannot be rewound. See issue #866.
        """
        image_file = _ensure_seekable(image_file)
        _raise_if_svg(image_file)
        image = Image.from_file(image_file)
        image_part = self._find_by_sha1(image.sha1)
        return image_part if image_part else ImagePart.new(self._package, image)

    def get_or_add_svg_part(self, svg_file: str | IO[bytes]) -> ImagePart:
        """Return |ImagePart| holding the SVG bytes in `svg_file`.

        Deduplicates against existing SVG image parts by SHA1 digest of the
        raw SVG bytes, so embedding the same SVG on multiple slides shares a
        single ``/ppt/media/imageN.svg`` part. Never runs the SVG-rejection
        sniff used by :meth:`get_or_add_image_part` — on this path SVG is
        exactly what the caller supplied. See issue #358.

        .. versionadded:: 2026.05.0
        """
        blob, filename = _read_bytes(svg_file)
        sha1 = hashlib.sha1(blob).hexdigest()
        existing = self._find_svg_by_sha1(sha1)
        if existing is not None:
            return existing
        return ImagePart.new_svg(self._package, blob, filename)

    def _find_svg_by_sha1(self, sha1: str) -> ImagePart | None:
        """Return the package's SVG |ImagePart| matching `sha1`, else |None|.

        Walks the existing image parts in this package and matches any whose
        partname ends in ``.svg`` — the PIL-backed ``sha1`` property used
        for raster images doesn't work here (Pillow can't open SVG), so we
        compute the digest on the part's raw blob instead.
        """
        for image_part in self:
            # -- only consider SVG parts; raster parts use a different hash path --
            if not image_part.partname.endswith(".svg"):
                continue
            if hashlib.sha1(image_part.blob).hexdigest() == sha1:
                return image_part
        return None

    def _find_by_sha1(self, sha1: str) -> ImagePart | None:
        """
        Return an |ImagePart| object belonging to this package or |None| if
        no matching image part is found. The image part is identified by the
        SHA1 hash digest of the image binary it contains.
        """
        for image_part in self:
            # ---defensively skip unsupported image types that may already be
            # ---present in the package (e.g., an SVG part loaded from an
            # ---existing .pptx). See ``_raise_if_svg`` for the write path.
            if not hasattr(image_part, "sha1"):
                continue
            if image_part.sha1 == sha1:
                return image_part
        return None


def _raise_if_svg(image_file: str | IO[bytes]) -> None:
    """Raise |UnsupportedImageTypeError| when `image_file` contains SVG content.

    Sniffs the first few kilobytes of the image byte-stream for an ``<svg ...>``
    root element. Falls back to a filename-extension check (gated on the stream
    at least looking like XML) to catch heavily-commented SVGs whose root tag
    sits beyond the first 2 KiB.

    The check preserves the caller's original stream position so that
    downstream consumers (PIL, etc.) see the full original blob when this
    function returns without raising.
    """
    head, restore = _read_image_head(image_file)
    name = _image_file_name(image_file)

    if _looks_like_svg(head, name):
        hint = f" (file '{name}')" if name else ""
        raise UnsupportedImageTypeError(
            "SVG images are not supported by python-pptx" + hint + ". "
            "PowerPoint's SVG support requires a companion PNG raster fallback; "
            "python-pptx does not include an SVG rasterizer. "
            "Pre-rasterize the SVG to PNG (e.g. with `cairosvg`, `svglib` + "
            "`reportlab`, or `Pillow` via `librsvg`) and insert the resulting "
            "PNG instead. See the 'Inserting SVG images' section of the user "
            "guide for details."
        )

    if restore is not None:
        restore()


def _ensure_seekable(image_file: str | IO[bytes]) -> str | IO[bytes]:
    """Return a seekable form of `image_file`, materializing if necessary.

    Paths and already-seekable file-like objects are returned unchanged. A
    file-like object without a functioning ``seek`` method (e.g. a pipe-backed
    reader, a ``urllib`` response, or a stream backed by a browser blob URL)
    is fully read into a :class:`io.BytesIO` which is then returned in its
    place. The returned BytesIO preserves the original's ``name`` attribute
    when present so downstream filename inference still works.

    Called at the top of :meth:`_ImageParts.get_or_add_image_part` so every
    subsequent consumer (SVG sniff, PIL format detection, SHA1 digest) sees
    the full blob from offset 0. See issue #866.
    """
    if isinstance(image_file, str):
        return image_file

    seek = getattr(image_file, "seek", None)
    tell = getattr(image_file, "tell", None)
    if callable(seek) and callable(tell):
        try:
            pos = tell()
            seek(pos)
        except (OSError, ValueError):
            # -- stream advertises seek/tell but rejects them; treat as non-seekable --
            pass
        else:
            return image_file

    # -- non-seekable: materialize and preserve a `name` attribute if present --
    blob = image_file.read()
    buf = io.BytesIO(blob)
    name = getattr(image_file, "name", None)
    if isinstance(name, str):
        # -- preserves `os.path.basename(buf.name)` downstream filename inference --
        buf.name = name
    return buf


def _read_image_head(image_file: str | IO[bytes]):
    """Return `(head_bytes, restore_fn_or_None)` for SVG sniffing.

    When `image_file` is a file-like object, reads up to `_SVG_HEAD_BYTES` from
    its current position and returns a callable that rewinds the stream back to
    that position so downstream readers see the full blob.
    """
    if isinstance(image_file, str):
        try:
            with open(image_file, "rb") as f:
                return f.read(_SVG_HEAD_BYTES), None
        except OSError:
            return b"", None

    # -- assume file-like object --
    if callable(getattr(image_file, "tell", None)) and callable(getattr(image_file, "seek", None)):
        pos = image_file.tell()
        head = image_file.read(_SVG_HEAD_BYTES)

        def restore():
            image_file.seek(pos)

        return head, restore

    # -- non-seekable stream: best-effort peek (consumes bytes) --
    head = image_file.read(_SVG_HEAD_BYTES)
    return head, None


def _read_bytes(image_file: str | IO[bytes]) -> tuple[bytes, str | None]:
    """Return `(blob, filename)` read from `image_file`.

    `image_file` is either a path (``str``) or a seekable/readable file-like
    object. When `image_file` is a path, its basename is returned as the
    filename. For file-like input a ``name`` attribute is used when present,
    otherwise ``None``.
    """
    if isinstance(image_file, str):
        with open(image_file, "rb") as f:
            blob = f.read()
        return blob, os.path.basename(image_file)

    # -- file-like: rewind if possible so reads from prior positions don't truncate --
    if callable(getattr(image_file, "seek", None)):
        image_file.seek(0)
    blob = image_file.read()
    name_attr = getattr(image_file, "name", None)
    filename = os.path.basename(name_attr) if isinstance(name_attr, str) else None
    return blob, filename


def _image_file_name(image_file: str | IO[bytes]) -> str | None:
    """Return a filename for `image_file` if one is knowable, else |None|."""
    if isinstance(image_file, str):
        return os.path.basename(image_file)
    name = getattr(image_file, "name", None)
    if isinstance(name, str):
        return os.path.basename(name)
    return None


def _looks_like_svg(head: bytes, name: str | None) -> bool:
    """Return True if `head` and/or `name` indicate an SVG document."""
    if _SVG_SNIFF_RE.search(head):
        return True
    # -- the SVG root may live beyond the first 2 KiB in a heavily-commented
    # -- file; fall back to the filename extension in that case (only when the
    # -- content at least looks like XML, to avoid false positives on unrelated
    # -- `.svg`-named binary files).
    if name and name.lower().endswith(".svg"):
        stripped = head.lstrip()
        if stripped.startswith(b"<?xml") or stripped.startswith(b"<!"):
            return True
    return False


class _MediaParts(object):
    """Provides access to the media parts in a package.

    Supports iteration and :meth:`get()` using the media object SHA1 hash as
    its key.
    """

    def __init__(self, package):
        super(_MediaParts, self).__init__()
        self._package = package

    def __iter__(self):
        """Generate a reference to each |MediaPart| object in the package."""
        # A media part can appear in more than one relationship (and commonly
        # does in the case of video). Use media_parts to keep track of those
        # that have been "yielded"; they can be skipped if they occur again.
        # AUDIO relationships are included so click-action sounds (a:snd) can be
        # de-duplicated alongside video media.
        media_parts = []
        for rel in self._package.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype not in (RT.AUDIO, RT.MEDIA, RT.VIDEO):
                continue
            media_part = rel.target_part
            if media_part in media_parts:
                continue
            media_parts.append(media_part)
            yield media_part

    def get_or_add_media_part(self, media):
        """Return a |MediaPart| object containing the media in *media*.

        If this package already contains a media part for the same
        bytestream, that instance is returned, otherwise a new media part is
        created.
        """
        media_part = self._find_by_sha1(media.sha1)
        if media_part is None:
            media_part = MediaPart.new(self._package, media)
        return media_part

    def _find_by_sha1(self, sha1):
        """Return |MediaPart| object having *sha1* hash or None if not found.

        All media parts belonging to this package are considered. A media
        part is identified by the SHA1 hash digest of its bytestream
        ("file"). Parts that were loaded as generic |Part| instances (for
        example an audio clip whose content type was not mapped to
        |MediaPart|) are skipped — they can never match and don't expose a
        ``sha1`` attribute.
        """
        for media_part in self:
            if not isinstance(media_part, MediaPart):
                continue
            if media_part.sha1 == sha1:
                return media_part
        return None
