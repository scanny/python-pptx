"""Main presentation object."""

from __future__ import annotations

import re
import uuid
from typing import IO, TYPE_CHECKING, Iterable, Iterator, cast

from pptx.enum.presentation import PP_VIEW_TYPE
from pptx.shared import PartElementProxy
from pptx.slide import SlideMasters, Slides
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.opc.serialized import ZipDateTime
    from pptx.oxml.presentation import (
        CT_Presentation,
        CT_Section,
        CT_SlideId,
    )
    from pptx.oxml.viewprops import (
        CT_CommonSlideViewProperties,
        CT_CommonViewProperties,
        CT_OutlineViewProperties,
        CT_SlideSorterViewProperties,
        CT_ViewProperties,
    )
    from pptx.parts.extprops import ExtendedPropertiesPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.viewprops import ViewPropsPart
    from pptx.slide import NotesMaster, Slide, SlideLayouts
    from pptx.util import Length


_VALID_FONT_STYLES = ("regular", "bold", "italic", "boldItalic")

# -- e.g. "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}" -- PowerPoint GUID format --
_GUID_RE = re.compile(
    r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}" r"-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
)


def _new_section_id() -> str:
    """Return a freshly-generated section GUID in PowerPoint canonical form."""
    return "{%s}" % str(uuid.uuid4()).upper()


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

    @property
    def custom_properties(self):
        """Dict-like |CustomProperties| instance for this presentation.

        Provides get/set/delete access to user-defined custom document properties — the
        ones PowerPoint surfaces under *File → Info → Properties → Advanced → Custom* and
        that ``{ DOCPROPERTY }`` field codes resolve against. Supported value types are
        ``str``, ``int``, ``float``, ``bool``, and ``datetime.datetime``. See issue #259.

        .. versionadded:: 2026.05.0
        """
        return self.part.package.custom_properties

    @property
    def extended_properties(self) -> ExtendedPropertiesPart:
        """|ExtendedPropertiesPart| for this presentation (``/docProps/app.xml``).

        Provides read/write access to application-level document properties such as
        ``application``, ``app_version``, ``company``, ``manager``, ``hyperlink_base``,
        ``presentation_format``, and ``template``. Also exposes ``slide_count`` (read-only
        in practice -- the count is refreshed automatically at save-time from the
        presentation's slide list; see issue #131). The part is created lazily if the
        package does not already contain one.

        .. versionadded:: 2026.05.0
        """
        return self.part.extended_properties

    @property
    def first_slide_num(self) -> int:
        """Display number PowerPoint uses for the first slide (read/write).

        Corresponds to ``p:presentation/@firstSlideNum`` and the "Number
        slides from" field under *Design → Slide Size → Custom Slide
        Size…*. Defaults to ``1`` when the attribute is absent (the
        schema default). Setting it to ``1`` removes the attribute.

        This controls slide-number *display only* (the slide-number
        placeholder on each slide); it does not reorder slides. See
        issue #94.
        """
        return self._element.firstSlideNum

    @first_slide_num.setter
    def first_slide_num(self, value: int) -> None:
        if isinstance(
            value, bool
        ) or not isinstance(  # pyright: ignore[reportUnnecessaryIsInstance]
            value, int
        ):
            raise TypeError("first_slide_num must be an int, got %s" % type(value).__name__)
        self._element.firstSlideNum = value

    @lazyproperty
    def view_props(self) -> ViewProps:
        """|ViewProps| accessor for the view-properties part of this package.

        Provides read/write access to editor-view settings that PowerPoint
        restores when the presentation is re-opened: the last-active view
        (:attr:`ViewProps.view_type`), whether the comments pane is visible
        (:attr:`ViewProps.show_comments`), the slide-sorter's
        formatting toggle (:attr:`ViewProps.show_formatting`), and the
        per-view zoom (:attr:`ViewProps.slide_view_zoom`,
        :attr:`ViewProps.notes_view_zoom`, :attr:`ViewProps.outline_view_zoom`,
        :attr:`ViewProps.sorter_view_zoom`). The underlying
        ``ppt/viewProps.xml`` part is created lazily if the package does
        not already contain one. See issue #94.

        .. versionadded:: 2026.05.0
        """
        return ViewProps(self.part.package.view_props_part)

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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
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

    def set_auto_advance(
        self,
        seconds: float | int | None,
        advance_on_click: bool = False,
    ) -> None:
        """Configure every slide in this presentation to auto-advance after `seconds`.

        Convenience bulk wrapper around per-slide ``Slide.transition.advance_after_time``
        and ``Slide.transition.advance_on_click``: applies the same auto-advance setting
        to every slide currently in the presentation. `seconds` is the delay before
        each slide auto-advances during a slide show, specified in seconds (accepts
        ``int`` or ``float`` — e.g. ``0.5`` for half a second). ``None`` disables
        auto-advance, removing the ``@advTm`` attribute from every slide's
        ``p:transition`` and restoring click-only advance.

        `advance_on_click` controls whether a mouse click can also advance past the
        transition while the timer is pending. Defaults to ``False`` (kiosk-style:
        timer only). Set ``advance_on_click=True`` to allow either a click OR the
        timer to advance — matching PowerPoint's "On Mouse Click / After" combined
        mode. When `seconds` is ``None`` the `advance_on_click` argument is still
        applied (so ``set_auto_advance(None, advance_on_click=True)`` returns every
        slide to plain click-to-advance).

        Only slides currently in the deck are updated; slides added after this call
        use their schema-default advance behavior until configured individually or
        via another ``set_auto_advance`` invocation.

        .. versionadded:: 2026.05.0
        """
        if seconds is None:
            advance_after_time_ms: int | None = None
        else:
            if isinstance(seconds, bool) or not isinstance(seconds, (int, float)):
                raise TypeError(
                    "seconds must be a non-negative number or None, got %s" % type(seconds).__name__
                )
            if seconds < 0:
                raise ValueError(
                    "seconds must be a non-negative number or None, got %r" % (seconds,)
                )
            advance_after_time_ms = int(round(seconds * 1000))

        if not isinstance(advance_on_click, bool):
            raise TypeError(
                "advance_on_click must be a bool, got %s" % type(advance_on_click).__name__
            )

        for slide in self.slides:
            transition = slide.transition
            transition.advance_after_time = advance_after_time_ms
            transition.advance_on_click = advance_on_click

    def merge(self, other_presentation: Presentation) -> list[Slide]:
        """Append every slide of `other_presentation` to this presentation.

        Returns the list of newly-appended |Slide| objects in the same order
        they appear in `other_presentation`. Each appended slide is a
        full-fidelity deep copy: its shape tree, image and media parts, charts
        (with distinct embedded workbooks), embedded OLE objects, and external
        hyperlinks are all materialised in *this* presentation's package. The
        source presentation is not modified.

        Each cloned slide is bound to the layout at the *same index* in this
        presentation's primary slide master (``slide_masters[0].slide_layouts``)
        as its source's layout occupied in the source presentation's primary
        master. When the target master has fewer layouts than the index
        demanded by the source, the last layout is used as a fallback so the
        merge does not fail on a mismatched template; callers who need strict
        layout mapping should reassign ``slide.slide_layout`` after the merge
        or use :meth:`Slides.add_slide_from_external` for per-slide control.

        Notes-slide relationships on source slides are dropped (a notes slide
        carries a back-reference to its owning slide and so cannot be shared);
        other per-slide metadata (comments, OLE objects, linked-picture rels)
        is carried along.

        Raises |TypeError| if `other_presentation` is not a |Presentation| and
        |ValueError| when called with ``self`` as the argument (to avoid
        inadvertently doubling a presentation's slide count).

        .. versionadded:: 2026.05.0
        """
        if not isinstance(other_presentation, Presentation):
            raise TypeError(
                "other_presentation must be a Presentation, got %s"
                % type(other_presentation).__name__
            )
        if other_presentation is self:
            raise ValueError("cannot merge a presentation into itself")

        target_layouts = self.slide_masters[0].slide_layouts
        if len(target_layouts) == 0:
            raise ValueError("target presentation has no slide layouts on its primary master")

        source_layouts = other_presentation.slide_masters[0].slide_layouts

        appended: list[Slide] = []
        for source_slide in other_presentation.slides:
            # -- Pick the target layout whose index matches the source slide's
            # -- layout index in the source presentation's primary master. Fall
            # -- back to the last layout if the target master is shorter. --
            source_layout = source_slide.slide_layout
            try:
                layout_idx = source_layouts.index(source_layout)
            except (ValueError, KeyError):
                layout_idx = 0
            if layout_idx >= len(target_layouts):
                layout_idx = len(target_layouts) - 1
            target_layout = target_layouts[layout_idx]

            appended.append(self.slides.add_slide_from_external(source_slide, target_layout))

        return appended

    def strip_slides(self) -> Presentation:
        """Remove every slide from this presentation and return ``self``.

        Provides the "use an existing ``.pptx`` as a blank template" workflow
        of issue #310: open an existing deck with :func:`pptx.Presentation`,
        strip the slides it came with, and use the remaining (empty)
        presentation as the starting point for a new deck — the slide
        masters, slide layouts, theme, default table styles, embedded fonts,
        and any other template-level resources are preserved.

        Every slide currently in :attr:`slides` is deleted via
        :meth:`Slides.delete`, so the same reference-counting cleanup applies
        (the slide parts and any image / media / chart parts reachable only
        from them become unreachable from the package root and are omitted on
        the next save; parts shared with the slide master, layouts, or other
        surviving parts are retained).

        Any presentation :attr:`sections` are cleared as well, since sections
        reference slides by slide-id and an empty ``p14:sectionLst`` with
        dangling ids would confuse PowerPoint. The presentation's section
        scaffolding (``p:extLst/p:ext/p14:sectionLst``) is pruned so a
        freshly-stripped presentation round-trips to an empty
        ``sldIdLst`` and no ``sectionLst``.

        Returns ``self`` so the method chains after the
        :func:`pptx.Presentation` factory call::

            blank = Presentation("branded-template.pptx").strip_slides()
            blank.slides.add_slide(blank.slide_layouts[0])

        Calling :meth:`strip_slides` on a presentation that already has no
        slides is a no-op (returns ``self`` unchanged).

        .. versionadded:: 2026.05.0
        """
        # -- Drop every slide using the public delete() API so reference-
        # -- counting and part cleanup run identically to per-slide deletes.
        # -- Iterate over a snapshot since `delete()` mutates the collection. --
        for slide in list(self.slides):
            self.slides.delete(slide)

        # -- Clear sections: an empty section list with stale sldId refs
        # -- would confuse PowerPoint. Removing each section via the Sections
        # -- API prunes the enclosing p:extLst/p:ext scaffolding when the
        # -- last section is removed. --
        for section in list(self.sections):
            self.sections.remove(section)

        return self

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

    def save_ppsx(
        self,
        file: str | IO[bytes],
        zip_date_time: ZipDateTime | None = None,
        password: str | None = None,
    ) -> None:
        """Write this presentation to `file` as a PowerPoint Show (``.ppsx``).

        Identical to :meth:`save` except the presentation part's content-type
        override is written as
        ``application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml``
        rather than the regular ``…presentation.main+xml``. Named with a
        ``.ppsx`` extension, the resulting file opens in PowerPoint straight
        into slide-show playback (no authoring UI).

        `file` is either a filesystem path (``str``) or a file-like object open
        for writing bytes. `zip_date_time` and `password` behave exactly as
        they do for :meth:`save` (reproducible timestamps and ECMA-376 Agile
        Encryption, respectively). See issue #438.

        .. versionadded:: 2026.05.0
        """
        self.part.save_ppsx(file, zip_date_time, password=password)

    def save_flat_xml(self, file: str | IO[bytes]) -> None:
        """Write this presentation to `file` as a Flat OPC single-file XML document.

        Flat OPC is the "XML Presentation" (``.xml``) format defined in ECMA-376
        Part 4 and produced by PowerPoint's *Save As \N{RIGHTWARDS ARROW} XML
        Presentation* command. The entire package (the presentation part, every
        slide, master, layout, theme, image, and so on) is serialized as one XML
        document rooted at ``<pkg:package>``; XML parts are embedded verbatim
        while binary parts (images, embedded fonts, OLE objects, media) are
        base64-encoded inline.

        `file` is either a filesystem path (``str``) or a file-like object open
        for writing bytes. Unlike :meth:`save`, this method does not accept
        ``zip_date_time`` or ``password`` -- the Flat OPC format is an XML
        document, not a zip archive, and is not password-protectable.

        .. versionadded:: 2026.05.0
        """
        self.part.save_flat_xml(file)

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

    @lazyproperty
    def sections(self) -> Sections:
        """|Sections| object providing access to the presentation's section list.

        Supports iteration, ``len()``, and indexed access. Returns an empty collection when
        the presentation has no sections defined. Adding the first section creates the
        ``p:extLst/p:ext/p14:sectionLst`` chain on demand.

        .. versionadded:: 2026.05.0
        """
        return Sections(self._element, self.part)


class Sections:
    """Sequence of |Section| objects belonging to a |Presentation|.

    Has list semantics for indexed access, ``len()``, and iteration. Create sections
    via :meth:`add_section` and remove them with :meth:`remove`. Sections are
    identified by a GUID in the PowerPoint-canonical string form
    ``"{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"``.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, prs_element: CT_Presentation, prs_part: PresentationPart | None):
        super().__init__()
        self._prs_element = prs_element
        self._prs_part = prs_part

    def __getitem__(self, idx: int) -> Section:
        """Provide indexed access to sections, e.g. ``presentation.sections[0]``."""
        sectionLst = self._prs_element.sectionLst
        if sectionLst is None:
            raise IndexError("section index out of range")
        try:
            section_el = sectionLst.section_lst[idx]
        except IndexError as exc:
            raise IndexError("section index out of range") from exc
        return Section(section_el, self._prs_element, self._prs_part)

    def __iter__(self) -> Iterator[Section]:
        """Iterate sections in document order."""
        sectionLst = self._prs_element.sectionLst
        if sectionLst is None:
            return
        for section_el in sectionLst.section_lst:
            yield Section(section_el, self._prs_element, self._prs_part)

    def __len__(self) -> int:
        """Number of sections defined on this presentation (zero when absent)."""
        sectionLst = self._prs_element.sectionLst
        if sectionLst is None:
            return 0
        return len(sectionLst.section_lst)

    def add_section(
        self,
        name: str,
        slides: Iterable[Slide] = (),
        id: str | None = None,
    ) -> Section:
        """Append and return a new |Section| with display name `name`.

        `slides` is an optional iterable of |Slide| objects (each must already belong
        to this presentation) to be assigned to the new section; slides are referenced
        by `sldId` value, not by the slide's position in the deck.

        `id` optionally specifies the section GUID in PowerPoint-canonical form
        ``"{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"``; when omitted a fresh GUID is
        generated. A |ValueError| is raised if `id` is supplied but malformed,
        or if it duplicates an existing section's id.

        .. versionadded:: 2026.05.0
        """
        if id is None:
            section_id = _new_section_id()
        else:
            if not _GUID_RE.match(id):
                raise ValueError(
                    "id must be a GUID in '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}' form, got %r"
                    % id
                )
            for existing in self:
                if existing.id.lower() == id.lower():
                    raise ValueError("a section with id %r already exists" % id)
            section_id = id

        sectionLst = self._prs_element.get_or_add_sectionLst()
        section_el = sectionLst.add_section(name, section_id)
        section = Section(section_el, self._prs_element, self._prs_part)

        for slide in slides:
            section.add_slide(slide)

        return section

    def find_containing(self, slide: Slide) -> Section | None:
        """Return the |Section| that currently owns `slide`, or |None|.

        A slide can only belong to a single section at a time (matching PowerPoint's
        rule that every slide appears in exactly one section once the deck is
        partitioned), so the first containing section in document order is returned.

        Returns |None| when the presentation has no sections or the slide is not
        assigned to any section (e.g. it was added to the deck after the last
        section was created). Raises |ValueError| if `slide` does not belong to
        this presentation.

        .. versionadded:: 2026.05.0
        """
        sldIdLst = self._prs_element.sldIdLst
        if sldIdLst is None or self._prs_part is None:
            raise ValueError("slide does not belong to the owning presentation")
        slide_id: int | None = None
        for sldId in sldIdLst.sldId_lst:
            if self._prs_part.related_slide(sldId.rId) == slide:
                slide_id = sldId.id
                break
        if slide_id is None:
            raise ValueError("slide does not belong to the owning presentation")

        for section in self:
            if slide_id in section._section_slide_ids:  # pyright: ignore[reportPrivateUsage]
                return section
        return None

    def get_by_id(self, id: str) -> Section | None:
        """Return the section whose GUID equals `id` (case-insensitive), or |None|.

        .. versionadded:: 2026.05.0
        """
        target = id.lower()
        for section in self:
            if section.id.lower() == target:
                return section
        return None

    def get_by_name(self, name: str) -> Section | None:
        """Return the first section whose `name` equals `name`, or |None|.

        Section names are *not* required to be unique in the presentation; the first
        match in document order is returned.

        .. versionadded:: 2026.05.0
        """
        for section in self:
            if section.name == name:
                return section
        return None

    def index(self, section: Section) -> int:
        """Return the zero-based position of `section` in this collection.

        Raises |ValueError| if `section` does not belong to this presentation.

        .. versionadded:: 2026.05.0
        """
        sectionLst = self._prs_element.sectionLst
        if sectionLst is not None:
            for idx, section_el in enumerate(sectionLst.section_lst):
                if section_el is section.element:
                    return idx
        raise ValueError("section is not a member of this collection")

    def remove(self, section: Section) -> None:
        """Remove `section` from this presentation.

        The `p14:section` element is removed from the section list. If the last
        section is removed the enclosing `p14:sectionLst` and `p:ext` elements are
        pruned as well, leaving the `p:extLst` clean.

        Raises |ValueError| if `section` does not belong to this presentation.

        .. versionadded:: 2026.05.0
        """
        sectionLst = self._prs_element.sectionLst
        if sectionLst is None or section.element not in list(sectionLst.section_lst):
            raise ValueError("section is not a member of this collection")

        sectionLst.remove(section.element)

        # -- prune empty containers so we don't leave dangling scaffolding --
        if not sectionLst.section_lst:
            ext = sectionLst.getparent()
            extLst = ext.getparent() if ext is not None else None
            if ext is not None and extLst is not None:
                extLst.remove(ext)
            if extLst is not None and len(extLst) == 0:
                prs = extLst.getparent()
                if prs is not None:
                    prs.remove(extLst)


class Section:
    """A single presentation section — a named grouping of slides.

    Exposes :attr:`name` (read/write), :attr:`id` (GUID, read-only), and
    :attr:`slides`, a tuple of the slides currently assigned to this section.
    Not intended to be constructed directly; obtain |Section| instances via
    :attr:`Presentation.sections`.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        section_elm: CT_Section,
        prs_element: CT_Presentation,
        prs_part: PresentationPart | None,
    ):
        super().__init__()
        self._element = section_elm
        self._prs_element = prs_element
        self._prs_part = prs_part

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Section):
            return NotImplemented
        return self._element is other._element

    def __hash__(self) -> int:
        return id(self._element)

    @property
    def element(self) -> CT_Section:
        """The underlying `p14:section` oxml element for this section.

        .. versionadded:: 2026.05.0
        """
        return self._element

    @property
    def id(self) -> str:
        """Section GUID in PowerPoint-canonical form, e.g. ``"{9B1E...DCB9C}"``.

        Read-only; ids are assigned when the section is added and are not modified
        afterwards (matching PowerPoint's behavior).

        .. versionadded:: 2026.05.0
        """
        return self._element.id

    @property
    def index(self) -> int:
        """Zero-based position of this section within the presentation's section list.

        Raises |ValueError| if the section has been removed from its presentation
        (e.g. still held in a local variable after :meth:`Sections.remove`).

        .. versionadded:: 2026.05.0
        """
        sectionLst = self._prs_element.sectionLst
        if sectionLst is not None:
            for idx, section_el in enumerate(sectionLst.section_lst):
                if section_el is self._element:
                    return idx
        raise ValueError("section is not a member of its presentation")

    @property
    def name(self) -> str:
        """Display name of this section (read/write).

        Setting an empty string is permitted and matches PowerPoint's "rename to
        blank" behavior, though the UI will usually re-populate the field.
        """
        return self._element.name

    @name.setter
    def name(self, value: str) -> None:
        self._element.name = value

    @property
    def slides(self) -> tuple[Slide, ...]:
        """Tuple of |Slide| objects assigned to this section, in section order.

        Slides referenced by a `p14:sldId/@id` value that does not match any slide
        currently in the presentation (e.g. stale reference left by third-party
        tools) are silently skipped rather than raising.

        .. versionadded:: 2026.05.0
        """
        sldIdLst = self._prs_element.sldIdLst
        if sldIdLst is None or self._prs_part is None:
            return ()
        id_to_rId = {sldId.id: sldId.rId for sldId in sldIdLst.sldId_lst}
        slides: list[Slide] = []
        for slide_id in self._section_slide_ids:
            rId = id_to_rId.get(slide_id)
            if rId is None:
                continue
            slides.append(self._prs_part.related_slide(rId))
        return tuple(slides)

    def add_slide(self, slide: Slide) -> None:
        """Append `slide` to this section.

        `slide` must belong to the owning presentation; assigning a slide from a
        different presentation raises |ValueError|. If `slide` is already a member
        of this section the call is a no-op (PowerPoint preserves the existing
        position rather than appending a duplicate reference).

        A slide may belong to at most one section. If `slide` is already assigned
        to a *different* section, |ValueError| is raised; call
        :meth:`Section.remove_slide` on that section (or :meth:`move_slide` on the
        target section) before re-assigning.

        .. versionadded:: 2026.05.0
        """
        # -- locate the `p:sldId` for `slide` to obtain its id value --
        slide_id = self._resolve_slide_id(slide)

        if slide_id in self._section_slide_ids:
            return

        # -- reject reassignment from another section; PowerPoint's model is
        # -- one-section-per-slide and silently duplicating the reference would
        # -- corrupt the deck on save.
        other = self._find_other_section_owning(slide_id)
        if other is not None:
            raise ValueError(
                "slide is already assigned to section %r; remove it from that "
                "section first, or call Section.move_slide()" % other.name
            )

        sldIdLst = self._element.get_or_add_sldIdLst()
        sldIdLst.add_sldId(slide_id)

    def move_after(self, other: Section) -> None:
        """Reposition this section immediately after `other` in the section list.

        `other` must belong to the same presentation. Moving a section before or
        after itself is a no-op. Raises |ValueError| if either section is not a
        member of this presentation's section list.

        .. versionadded:: 2026.05.0
        """
        self._reposition(other, before=False)

    def move_before(self, other: Section) -> None:
        """Reposition this section immediately before `other` in the section list.

        `other` must belong to the same presentation. Moving a section before or
        after itself is a no-op. Raises |ValueError| if either section is not a
        member of this presentation's section list.

        .. versionadded:: 2026.05.0
        """
        self._reposition(other, before=True)

    def move_slide(self, slide: Slide) -> None:
        """Reassign `slide` to this section, removing it from its current section.

        Convenience for the common "move this slide into section X" workflow.
        If `slide` is already a member of this section, the call is a no-op.
        Raises |ValueError| if `slide` does not belong to the owning presentation.

        .. versionadded:: 2026.05.0
        """
        slide_id = self._resolve_slide_id(slide)
        if slide_id in self._section_slide_ids:
            return

        other = self._find_other_section_owning(slide_id)
        if other is not None:
            other.remove_slide(slide)

        sldIdLst = self._element.get_or_add_sldIdLst()
        sldIdLst.add_sldId(slide_id)

    def remove_slide(self, slide: Slide) -> None:
        """Remove `slide`'s reference from this section.

        Raises |ValueError| if `slide` is not assigned to this section. The slide
        itself is left untouched in the presentation; only the section-membership
        reference is removed.

        .. versionadded:: 2026.05.0
        """
        slide_id = self._resolve_slide_id(slide)

        sldIdLst = self._element.sldIdLst
        if sldIdLst is None:
            raise ValueError("slide is not a member of this section")

        for entry in sldIdLst.sldId_lst:
            if entry.id == slide_id:
                sldIdLst.remove(entry)
                return
        raise ValueError("slide is not a member of this section")

    @property
    def _section_slide_ids(self) -> tuple[int, ...]:
        """Tuple of `p14:sldId/@id` values currently in this section, in order."""
        sldIdLst = self._element.sldIdLst
        if sldIdLst is None:
            return ()
        return tuple(entry.id for entry in sldIdLst.sldId_lst)

    def _find_other_section_owning(self, slide_id: int) -> Section | None:
        """Return a *different* |Section| that currently contains `slide_id`, or |None|."""
        sectionLst = self._prs_element.sectionLst
        if sectionLst is None:
            return None
        for section_el in sectionLst.section_lst:
            if section_el is self._element:
                continue
            sldIdLst = section_el.sldIdLst
            if sldIdLst is None:
                continue
            for entry in sldIdLst.sldId_lst:
                if entry.id == slide_id:
                    return Section(section_el, self._prs_element, self._prs_part)
        return None

    def _reposition(self, other: Section, before: bool) -> None:
        """Shared implementation for :meth:`move_before` / :meth:`move_after`."""
        if other == self:
            return

        sectionLst = self._prs_element.sectionLst
        if sectionLst is None:
            raise ValueError("section is not a member of this presentation")

        members = list(sectionLst.section_lst)
        if self._element not in members or other.element not in members:
            raise ValueError("section is not a member of this presentation")

        # -- pull self out of the list then re-insert relative to `other` --
        sectionLst.remove(self._element)

        # -- `other` index may shift once `self` is removed, so recompute --
        anchor_idx = list(sectionLst.section_lst).index(other.element)
        target_idx = anchor_idx if before else anchor_idx + 1
        # -- lxml.etree._Element.insert; typed-stub gap in CT_SectionList --
        sectionLst.insert(  # pyright: ignore[reportUnknownMemberType,reportAttributeAccessIssue]
            target_idx, self._element
        )

    def _resolve_slide_id(self, slide: Slide) -> int:
        """Return the `p:sldId/@id` integer value of `slide` in the owning presentation.

        Raises |ValueError| if `slide` does not belong to this presentation.
        """
        sldIdLst = self._prs_element.sldIdLst
        if sldIdLst is not None and self._prs_part is not None:
            for sldId in sldIdLst.sldId_lst:
                if self._prs_part.related_slide(sldId.rId) == slide:
                    return sldId.id
        raise ValueError("slide does not belong to the owning presentation")


class ViewProps:
    """Editor-view settings for a presentation — issue #94 MVP.

    Wraps the ``p:viewPr`` element (``ppt/viewProps.xml``) and exposes the
    most commonly-used fields: the last-active view (:attr:`view_type`),
    the comments-pane toggle (:attr:`show_comments`), the slide-sorter's
    text-formatting toggle (:attr:`show_formatting`), and the per-view
    zoom properties (:attr:`slide_view_zoom`, :attr:`notes_view_zoom`,
    :attr:`outline_view_zoom`, :attr:`sorter_view_zoom`).

    Not constructed directly; obtain an instance via
    :attr:`Presentation.view_props`. The underlying part is created
    lazily the first time :attr:`Presentation.view_props` is read.

    Raw :attr:`element` access is provided as an escape-hatch for callers
    that need to read or write attributes that the MVP does not yet model
    (e.g. ``showOutlineIcons`` on ``p:normalViewPr``).

    .. versionadded:: 2026.05.0
    """

    def __init__(self, part: ViewPropsPart):
        super().__init__()
        self._part = part

    @property
    def element(self) -> CT_ViewProperties:
        """The underlying ``p:viewPr`` lxml element (round-trip escape hatch).

        .. versionadded:: 2026.05.0
        """
        return self._part._element  # pyright: ignore[reportPrivateUsage]

    @property
    def part(self) -> ViewPropsPart:
        """The underlying |ViewPropsPart| (``ppt/viewProps.xml``).

        .. versionadded:: 2026.05.0
        """
        return self._part

    # -- `@lastView` on `p:viewPr` — which editor view PowerPoint will
    # -- restore on the next open.

    @property
    def view_type(self) -> PP_VIEW_TYPE:
        """The :class:`.PP_VIEW_TYPE` value of ``p:viewPr/@lastView``.

        Defaults to :attr:`PP_VIEW_TYPE.NORMAL` (``sldView``) when the
        attribute is absent, per the schema.
        """
        return PP_VIEW_TYPE.from_xml(self.element.lastView)

    @view_type.setter
    def view_type(self, value: PP_VIEW_TYPE) -> None:
        if not isinstance(value, PP_VIEW_TYPE):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(
                "view_type must be a PP_VIEW_TYPE member, got %s" % type(value).__name__
            )
        xml_value = value.xml_value
        if xml_value is None:
            raise ValueError("PP_VIEW_TYPE member %r has no XML mapping" % value.name)
        self.element.lastView = xml_value

    # -- `@showComments` on `p:viewPr` — whether the comments pane is
    # -- visible when PowerPoint opens the deck.

    @property
    def show_comments(self) -> bool:
        """``bool`` value of ``p:viewPr/@showComments``. Defaults to ``True``."""
        return self.element.showComments

    @show_comments.setter
    def show_comments(self, value: bool) -> None:
        if not isinstance(value, bool):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("show_comments must be a bool, got %s" % type(value).__name__)
        self.element.showComments = value

    # -- `@showFormatting` on `p:sorterViewPr` — whether slide-sorter
    # -- thumbnails render text formatting.

    @property
    def show_formatting(self) -> bool:
        """``bool`` value of ``p:sorterViewPr/@showFormatting``. Defaults to ``True``.

        Controls whether slide-sorter thumbnails render text formatting.
        Reading returns the schema default (``True``) when the
        ``p:sorterViewPr`` element is absent.
        """
        sorter = self.element.sorterViewPr
        if sorter is None:
            return True
        return sorter.showFormatting

    @show_formatting.setter
    def show_formatting(self, value: bool) -> None:
        if not isinstance(value, bool):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("show_formatting must be a bool, got %s" % type(value).__name__)
        sorter = self.element.get_or_add_sorterViewPr()
        sorter.showFormatting = value

    # -- per-view zoom --

    @property
    def slide_view_zoom(self) -> float:
        """Zoom of the slide-editing view as a ratio, e.g. ``1.0`` = 100%.

        Reads ``p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx``;
        returns ``1.0`` when any element in the chain is absent.
        """
        slide_view = self.element.slideViewPr
        if slide_view is None or slide_view.cSldViewPr is None:
            return 1.0
        cViewPr = slide_view.cSldViewPr.cViewPr
        return _zoom_of(cViewPr)

    @slide_view_zoom.setter
    def slide_view_zoom(self, value: float) -> None:
        _validate_zoom(value)
        slide_view = self.element.get_or_add_slideViewPr()
        cSldViewPr = slide_view.get_or_add_cSldViewPr()
        _set_zoom_on(cSldViewPr, value)

    @property
    def notes_view_zoom(self) -> float:
        """Zoom of the notes-page view as a ratio. See :attr:`slide_view_zoom`."""
        notes_view = self.element.notesViewPr
        if notes_view is None or notes_view.cSldViewPr is None:
            return 1.0
        return _zoom_of(notes_view.cSldViewPr.cViewPr)

    @notes_view_zoom.setter
    def notes_view_zoom(self, value: float) -> None:
        _validate_zoom(value)
        notes_view = self.element.get_or_add_notesViewPr()
        cSldViewPr = notes_view.get_or_add_cSldViewPr()
        _set_zoom_on(cSldViewPr, value)

    @property
    def outline_view_zoom(self) -> float:
        """Zoom of the outline view as a ratio. See :attr:`slide_view_zoom`."""
        outline_view = self.element.outlineViewPr
        if outline_view is None:
            return 1.0
        return _zoom_of(outline_view.cViewPr)

    @outline_view_zoom.setter
    def outline_view_zoom(self, value: float) -> None:
        _validate_zoom(value)
        outline_view = self.element.get_or_add_outlineViewPr()
        _set_zoom_on_cviewpr_parent(outline_view, value)

    @property
    def sorter_view_zoom(self) -> float:
        """Zoom of the slide-sorter view as a ratio. See :attr:`slide_view_zoom`."""
        sorter = self.element.sorterViewPr
        if sorter is None:
            return 1.0
        return _zoom_of(sorter.cViewPr)

    @sorter_view_zoom.setter
    def sorter_view_zoom(self, value: float) -> None:
        _validate_zoom(value)
        sorter = self.element.get_or_add_sorterViewPr()
        _set_zoom_on_cviewpr_parent(sorter, value)


def _validate_zoom(value: float) -> None:
    """Raise for any value that cannot be encoded as a ``CT_Ratio``."""
    if isinstance(value, bool) or not isinstance(  # pyright: ignore[reportUnnecessaryIsInstance]
        value, (int, float)
    ):
        raise TypeError("zoom must be a non-negative number, got %s" % type(value).__name__)
    if value <= 0:
        raise ValueError("zoom must be > 0, got %r" % (value,))


def _zoom_of(cViewPr: CT_CommonViewProperties | None) -> float:
    """Return the `float` zoom read from ``cViewPr/p:scale/a:sx``, or ``1.0``.

    Returns ``1.0`` when `cViewPr` is |None| or when the ``p:scale`` child
    is absent.
    """
    if cViewPr is None:
        return 1.0
    scale = cViewPr.scale
    if scale is None:
        return 1.0
    return scale.zoom


def _set_zoom_on(cSldViewPr: CT_CommonSlideViewProperties, value: float) -> None:
    """Set `value` as the zoom on ``cSldViewPr/p:cViewPr/p:scale``."""
    from pptx.oxml.viewprops import CT_CommonViewProperties as _CT_CVP

    cViewPr = cSldViewPr.cViewPr
    if cViewPr is None:
        cViewPr = _CT_CVP.new_default()
        cSldViewPr.append(cViewPr)
    scale = cViewPr.get_or_add_scale()
    scale.set_zoom(value)


def _set_zoom_on_cviewpr_parent(
    parent: CT_OutlineViewProperties | CT_SlideSorterViewProperties,
    value: float,
) -> None:
    """Set zoom on a view element that carries a direct ``p:cViewPr`` child.

    Used for ``p:outlineViewPr`` and ``p:sorterViewPr``, which wrap
    ``p:cViewPr`` directly (unlike ``p:slideViewPr`` / ``p:notesViewPr``
    which go through an intermediate ``p:cSldViewPr``).
    """
    from pptx.oxml.viewprops import CT_CommonViewProperties as _CT_CVP

    cViewPr = parent.cViewPr
    if cViewPr is None:
        cViewPr = _CT_CVP.new_default()
        parent.append(cViewPr)
    scale = cViewPr.get_or_add_scale()
    scale.set_zoom(value)
