"""Overall .pptx package."""

from __future__ import annotations

import os
import re
from typing import IO, Iterator

from pptx.exc import UnsupportedImageTypeError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import OpcPackage
from pptx.opc.packuri import PackURI
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.image import Image, ImagePart
from pptx.parts.media import MediaPart
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

    def get_or_add_image_part(self, image_file: str | IO[bytes]):
        """
        Return an |ImagePart| object containing the image in *image_file*. If
        the image part already exists in this package, it is reused,
        otherwise a new one is created.
        """
        return self._image_parts.get_or_add_image_part(image_file)

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
        """
        _raise_if_svg(image_file)
        image = Image.from_file(image_file)
        image_part = self._find_by_sha1(image.sha1)
        return image_part if image_part else ImagePart.new(self._package, image)

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
    if callable(getattr(image_file, "tell", None)) and callable(
        getattr(image_file, "seek", None)
    ):
        pos = image_file.tell()
        head = image_file.read(_SVG_HEAD_BYTES)

        def restore():
            image_file.seek(pos)

        return head, restore

    # -- non-seekable stream: best-effort peek (consumes bytes) --
    head = image_file.read(_SVG_HEAD_BYTES)
    return head, None


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
        media_parts = []
        for rel in self._package.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype not in (RT.MEDIA, RT.VIDEO):
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
        ("file").
        """
        for media_part in self:
            if media_part.sha1 == sha1:
                return media_part
        return None
