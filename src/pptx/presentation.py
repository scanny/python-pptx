"""Main presentation object."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, cast

from pptx.shared import PartElementProxy
from pptx.slide import SlideMasters, Slides
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.opc.serialized import ZipDateTime
    from pptx.oxml.presentation import CT_Presentation, CT_SlideId
    from pptx.parts.presentation import PresentationPart
    from pptx.slide import NotesMaster, SlideLayouts
    from pptx.util import Length


_VALID_FONT_STYLES = ("regular", "bold", "italic", "boldItalic")


def _read_blob(file: str | IO[bytes]) -> bytes:
    """Return the contents of `file` as bytes.

    `file` may be either a filesystem path (str) or a file-like object open for
    binary reading. File-like objects that support `.seek(0)` are rewound before
    reading so repeated calls read from the start.
    """
    if isinstance(file, str):
        with open(file, "rb") as f:
            return f.read()
    seek = getattr(file, "seek", None)
    if callable(seek):
        seek(0)
    return file.read()


class Presentation(PartElementProxy):
    """PresentationML (PML) presentation.

    Not intended to be constructed directly. Use :func:`pptx.Presentation` to open or
    create a presentation.
    """

    _element: CT_Presentation
    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def core_properties(self):
        """|CoreProperties| instance for this presentation.

        Provides read/write access to the Dublin Core document properties for the presentation.
        """
        return self.part.core_properties

    def embed_font(
        self,
        font_file: str | IO[bytes],
        typeface: str,
        style: str = "regular",
    ) -> None:
        """Embed the font in `font_file` under typeface name `typeface`.

        `font_file` is either a path to a TrueType / OpenType font file or a file-like
        object opened in binary-read mode. `typeface` is the font-family name that
        PowerPoint should match against typefaces referenced in run properties
        (e.g. ``"Pacifico"``). `style` selects which of the four style slots of an
        embedded-font entry this file occupies and must be one of ``"regular"``,
        ``"bold"``, ``"italic"``, or ``"boldItalic"``; it defaults to ``"regular"``.

        Calling :meth:`embed_font` a second time with the same `typeface` adds the
        new file under an additional style slot on the existing entry (e.g. a
        regular + bold pair); calling it again for the same `typeface`/`style`
        pair replaces the previously-embedded file.

        Does not automatically set ``embedTrueTypeFonts``/``saveSubsetFonts`` on
        the presentation element; PowerPoint treats presence of an
        ``embeddedFontLst`` entry as authoritative.
        """
        if style not in _VALID_FONT_STYLES:
            raise ValueError("style must be one of %r, got %r" % (list(_VALID_FONT_STYLES), style))

        font_blob = _read_blob(font_file)

        embeddedFontLst = self._element.get_or_add_embeddedFontLst()
        entry = embeddedFontLst.entry_for_typeface(typeface)
        if entry is None:
            entry = embeddedFontLst.add_embeddedFont(typeface)

        # -- drop the relationship for any prior font-data in this slot so the
        # -- replaced FontPart does not remain orphaned in the package.
        prior_rId = entry.rId_for_style(style)

        rId = self.part.add_embedded_font(font_blob)
        entry.set_rId_for_style(style, rId)

        if prior_rId is not None and prior_rId != rId:
            self.part.drop_rel(prior_rId)

    @property
    def embedded_fonts(self) -> tuple[str, ...]:
        """Tuple of typeface names embedded in this presentation, or empty tuple.

        The typeface names are returned in document order.
        """
        embeddedFontLst = self._element.embeddedFontLst
        if embeddedFontLst is None:
            return ()
        return tuple(entry.typeface for entry in embeddedFontLst.embeddedFont_lst)

    @property
    def notes_master(self) -> NotesMaster:
        """Instance of |NotesMaster| for this presentation.

        If the presentation does not have a notes master, one is created from a default template
        and returned. The same single instance is returned on each call.
        """
        return self.part.notes_master

    def save(
        self,
        file: str | IO[bytes],
        zip_date_time: ZipDateTime | None = None,
        password: str | None = None,
    ):
        """Writes this presentation to `file`.

        `file` can be either a file-path or a file-like object open for writing bytes.

        When `zip_date_time` is provided, every zip-member in the saved .pptx is stamped
        with that fixed last-modified timestamp rather than the current wall-clock time.
        This produces byte-identical output across repeated saves of identical content,
        which is convenient for reproducible builds and source-control diffs. The value
        may be a :class:`datetime.datetime` or a 6-tuple ``(year, month, day, hour,
        minute, second)``. The earliest representable Zip date is 1980-01-01.

        When `password` is provided, the saved .pptx is password-protected using ECMA-376
        Agile Encryption (the same scheme PowerPoint uses). Encryption requires the
        optional ``msoffcrypto-tool`` dependency. `zip_date_time` and `password` are
        orthogonal: `zip_date_time` stamps the inner (plaintext) zip members before the
        encryption wrapper is applied.
        """
        self.part.save(file, zip_date_time, password=password)

    @property
    def slide_height(self) -> Length | None:
        """Height of slides in this presentation, in English Metric Units (EMU).

        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cy

    @slide_height.setter
    def slide_height(self, height: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cy = height

    @property
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| collection belonging to the first |SlideMaster| of this presentation.

        A presentation can have more than one slide master and each master will have its own set
        of layouts. This property is a convenience for the common case where the presentation has
        only a single slide master.
        """
        return self.slide_masters[0].slide_layouts

    @property
    def slide_master(self):
        """
        First |SlideMaster| object belonging to this presentation. Typically,
        presentations have only a single slide master. This property provides
        simpler access in that common case.
        """
        return self.slide_masters[0]

    @lazyproperty
    def slide_masters(self) -> SlideMasters:
        """|SlideMasters| collection of slide-masters belonging to this presentation."""
        return SlideMasters(self._element.get_or_add_sldMasterIdLst(), self)

    @property
    def slide_width(self):
        """
        Width of slides in this presentation, in English Metric Units (EMU).
        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cx

    @slide_width.setter
    def slide_width(self, width: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cx = width

    @lazyproperty
    def slides(self):
        """|Slides| object containing the slides in this presentation."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        self.part.rename_slide_parts([cast("CT_SlideId", sldId).rId for sldId in sldIdLst])
        return Slides(sldIdLst, self)
