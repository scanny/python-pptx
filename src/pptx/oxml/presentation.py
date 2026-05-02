"""Custom element classes for presentation-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.oxml.simpletypes import ST_SlideId, ST_SlideSizeCoordinate, XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from pptx.oxml.text import CT_TextFont
    from pptx.util import Length


class CT_Presentation(BaseOxmlElement):
    """`p:presentation` element, root of the Presentation part stored as `/ppt/presentation.xml`."""

    get_or_add_sldSz: Callable[[], CT_SlideSize]
    get_or_add_sldIdLst: Callable[[], CT_SlideIdList]
    get_or_add_sldMasterIdLst: Callable[[], CT_SlideMasterIdList]
    get_or_add_embeddedFontLst: Callable[[], CT_EmbeddedFontList]
    get_or_add_extLst: Callable[[], CT_ExtensionList]

    sldMasterIdLst: CT_SlideMasterIdList | None = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "p:sldMasterIdLst",
            successors=(
                "p:notesMasterIdLst",
                "p:handoutMasterIdLst",
                "p:sldIdLst",
                "p:sldSz",
                "p:notesSz",
            ),
        )
    )
    sldIdLst: CT_SlideIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldIdLst", successors=("p:sldSz", "p:notesSz")
    )
    sldSz: CT_SlideSize | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldSz", successors=("p:notesSz",)
    )
    embeddedFontLst: CT_EmbeddedFontList | None = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "p:embeddedFontLst",
            successors=(
                "p:custShowLst",
                "p:photoAlbum",
                "p:custDataLst",
                "p:kinsoku",
                "p:defaultTextStyle",
                "p:modifyVerifier",
                "p:extLst",
            ),
        )
    )
    extLst: CT_ExtensionList | None = ZeroOrOne("p:extLst")  # pyright: ignore[reportAssignmentType]

    # -- URI identifying the 2010 sections extension (p14:sectionLst) --
    _SECTION_LIST_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"

    @property
    def sectionLst(self) -> CT_SectionList | None:
        """The `p14:sectionLst` element for this presentation, or |None| if not present.

        The section list is nested under `p:extLst/p:ext[uri="{521415D9-…}"]/p14:sectionLst`;
        this property walks the chain and returns the innermost element when all links exist.
        """
        extLst = self.extLst
        if extLst is None:
            return None
        ext = extLst.get_ext_by_uri(self._SECTION_LIST_EXT_URI)
        if ext is None:
            return None
        return ext.sectionLst

    def get_or_add_sectionLst(self) -> CT_SectionList:
        """Return `p14:sectionLst`, creating the extLst/ext/sectionLst chain if needed."""
        extLst = self.get_or_add_extLst()
        ext = extLst.get_or_add_ext_by_uri(self._SECTION_LIST_EXT_URI)
        return ext.get_or_add_sectionLst()


class CT_SlideId(BaseOxmlElement):
    """`p:sldId` element.

    Direct child of `p:sldIdLst` that contains an `rId` reference to a slide in the presentation.
    """

    id: int = RequiredAttribute("id", ST_SlideId)  # pyright: ignore[reportAssignmentType]
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideIdList(BaseOxmlElement):
    """`p:sldIdLst` element.

    Direct child of <p:presentation> that contains a list of the slide parts in the presentation.
    """

    sldId_lst: list[CT_SlideId]

    _add_sldId: Callable[..., CT_SlideId]
    sldId = ZeroOrMore("p:sldId")

    def add_sldId(self, rId: str) -> CT_SlideId:
        """Create and return a reference to a new `p:sldId` child element.

        The new `p:sldId` element has its r:id attribute set to `rId`.
        """
        return self._add_sldId(id=self._next_id, rId=rId)

    @property
    def _next_id(self) -> int:
        """The next available slide ID as an `int`.

        Valid slide IDs start at 256. The next integer value greater than the max value in use is
        chosen, which minimizes that chance of reusing the id of a deleted slide.
        """
        MIN_SLIDE_ID = 256
        MAX_SLIDE_ID = 2147483647

        used_ids = [int(s) for s in cast("list[str]", self.xpath("./p:sldId/@id"))]
        simple_next = max([MIN_SLIDE_ID - 1] + used_ids) + 1
        if simple_next <= MAX_SLIDE_ID:
            return simple_next

        # -- fall back to search for next unused from bottom --
        valid_used_ids = sorted(id for id in used_ids if (MIN_SLIDE_ID <= id <= MAX_SLIDE_ID))
        return (
            next(
                candidate_id
                for candidate_id, used_id in enumerate(valid_used_ids, start=MIN_SLIDE_ID)
                if candidate_id != used_id
            )
            if valid_used_ids
            else 256
        )


class CT_SlideMasterIdList(BaseOxmlElement):
    """`p:sldMasterIdLst` element.

    Child of `p:presentation` containing references to the slide masters that belong to the
    presentation.
    """

    sldMasterId_lst: list[CT_SlideMasterIdListEntry]

    sldMasterId = ZeroOrMore("p:sldMasterId")


class CT_SlideMasterIdListEntry(BaseOxmlElement):
    """
    ``<p:sldMasterId>`` element, child of ``<p:sldMasterIdLst>`` containing
    a reference to a slide master.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideSize(BaseOxmlElement):
    """`p:sldSz` element.

    Direct child of <p:presentation> that contains the width and height of slides in the
    presentation.
    """

    cx: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cx", ST_SlideSizeCoordinate
    )
    cy: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cy", ST_SlideSizeCoordinate
    )


class CT_EmbeddedFontDataId(BaseOxmlElement):
    """`p:regular`, `p:bold`, `p:italic`, or `p:boldItalic` element.

    Child of `p:embeddedFont` carrying an `r:id` relationship reference to an embedded
    `.fntdata` part for a specific typeface style.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_EmbeddedFontListEntry(BaseOxmlElement):
    """`p:embeddedFont` element.

    One entry in `p:embeddedFontLst`. Identifies a typeface (via the required child
    `p:font` element, type `a:CT_TextFont`) and carries up to four style-specific
    font-data relationship children (`p:regular`, `p:bold`, `p:italic`,
    `p:boldItalic`).
    """

    get_or_add_regular: Callable[[], CT_EmbeddedFontDataId]
    get_or_add_bold: Callable[[], CT_EmbeddedFontDataId]
    get_or_add_italic: Callable[[], CT_EmbeddedFontDataId]
    get_or_add_boldItalic: Callable[[], CT_EmbeddedFontDataId]

    font: CT_TextFont = OneAndOnlyOne("p:font")  # pyright: ignore[reportAssignmentType]
    regular: CT_EmbeddedFontDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:regular", successors=("p:bold", "p:italic", "p:boldItalic")
    )
    bold: CT_EmbeddedFontDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bold", successors=("p:italic", "p:boldItalic")
    )
    italic: CT_EmbeddedFontDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:italic", successors=("p:boldItalic",)
    )
    boldItalic: CT_EmbeddedFontDataId | None = ZeroOrOne(
        "p:boldItalic"
    )  # pyright: ignore[reportAssignmentType]

    @property
    def typeface(self) -> str:
        """The typeface name of this embedded-font entry."""
        return self.font.typeface

    @typeface.setter
    def typeface(self, value: str) -> None:
        self.font.typeface = value

    def rId_for_style(self, style: str) -> str | None:
        """Return the `rId` of the font-data part for `style`, or None if absent.

        `style` must be one of `"regular"`, `"bold"`, `"italic"`, or `"boldItalic"`.
        """
        child = getattr(self, style)
        return None if child is None else child.rId

    def set_rId_for_style(self, style: str, rId: str) -> None:
        """Set the `rId` of the font-data child-element for `style`.

        Creates the child element if it does not already exist. `style` must be one of
        `"regular"`, `"bold"`, `"italic"`, or `"boldItalic"`.
        """
        add_method = getattr(self, "get_or_add_%s" % style)
        child = add_method()
        child.rId = rId


class CT_EmbeddedFontList(BaseOxmlElement):
    """`p:embeddedFontLst` element.

    Child of `p:presentation` that contains the list of embedded-font entries.
    """

    embeddedFont_lst: list[CT_EmbeddedFontListEntry]

    _add_embeddedFont: Callable[..., CT_EmbeddedFontListEntry]
    _insert_embeddedFont: Callable[[CT_EmbeddedFontListEntry], CT_EmbeddedFontListEntry]
    embeddedFont = ZeroOrMore("p:embeddedFont")

    def add_embeddedFont(self, typeface: str) -> CT_EmbeddedFontListEntry:
        """Return a newly-added `p:embeddedFont` entry whose typeface is `typeface`.

        The new `p:embeddedFont` is created with its required `p:font` child already
        populated with `typeface` so the element is schema-valid immediately.
        """
        embeddedFont = cast("CT_EmbeddedFontListEntry", OxmlElement("p:embeddedFont"))
        font = cast("CT_TextFont", OxmlElement("p:font"))
        font.typeface = typeface
        embeddedFont.append(font)
        self._insert_embeddedFont(embeddedFont)
        return embeddedFont

    def entry_for_typeface(self, typeface: str) -> CT_EmbeddedFontListEntry | None:
        """Return existing `p:embeddedFont` child whose typeface matches `typeface`, else None."""
        for entry in self.embeddedFont_lst:
            if entry.typeface == typeface:
                return entry
        return None


class CT_Extension(BaseOxmlElement):
    """`p:ext` element.

    An extension point that wraps arbitrary extension-namespaced content,
    scoped by a required ``uri`` attribute whose value identifies the
    extension.
    """

    get_or_add_sectionLst: Callable[[], CT_SectionList]

    uri: str = RequiredAttribute("uri", XsdString)  # pyright: ignore[reportAssignmentType]

    sectionLst: CT_SectionList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p14:sectionLst"
    )


class CT_ExtensionList(BaseOxmlElement):
    """`p:extLst` element.

    A container for one or more `p:ext` extension elements. Each `p:ext` is
    identified by its ``uri`` attribute; extensions carry namespaced content
    defined outside the core ECMA-376 schemas (Office 2010/2013/2016 additions,
    notably ``p14:sectionLst`` for PowerPoint sections).
    """

    ext_lst: list[CT_Extension]

    ext = ZeroOrMore("p:ext")

    def get_ext_by_uri(self, uri: str) -> CT_Extension | None:
        """Return the `p:ext` child whose `uri` attribute equals `uri`, or |None|."""
        for ext in self.ext_lst:
            if ext.uri == uri:
                return ext
        return None

    def get_or_add_ext_by_uri(self, uri: str) -> CT_Extension:
        """Return the `p:ext` child for `uri`, creating it if not already present."""
        ext = self.get_ext_by_uri(uri)
        if ext is None:
            ext = cast("CT_Extension", OxmlElement("p:ext"))
            ext.uri = uri
            self.append(ext)
        return ext


class CT_SectionList(BaseOxmlElement):
    """`p14:sectionLst` element.

    Top-level container for presentation sections; nested under
    ``p:extLst/p:ext[uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"]``.
    """

    section_lst: list[CT_Section]

    _add_section: Callable[..., CT_Section]
    section = ZeroOrMore("p14:section")

    def add_section(self, name: str, id: str) -> CT_Section:
        """Create and return a new `p14:section` child with an empty `p14:sldIdLst`.

        `id` must be a GUID string in the PowerPoint-canonical form
        ``"{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"``. `name` is the
        user-visible section name.
        """
        section = cast("CT_Section", OxmlElement("p14:section"))
        section.name = name
        section.id = id
        # -- every p14:section element carries a p14:sldIdLst, even if empty --
        section.get_or_add_sldIdLst()
        self.append(section)
        return section


class CT_Section(BaseOxmlElement):
    """`p14:section` element.

    One section in a presentation's section list. A section has a display
    `name`, a stable `id` (GUID), and an owned `p14:sldIdLst` that lists
    the slides belonging to the section by their `p:sldId/@id` value.
    """

    get_or_add_sldIdLst: Callable[[], CT_SectionSlideIdList]

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    id: str = RequiredAttribute("id", XsdString)  # pyright: ignore[reportAssignmentType]

    sldIdLst: CT_SectionSlideIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p14:sldIdLst"
    )


class CT_SectionSlideIdList(BaseOxmlElement):
    """`p14:sldIdLst` element (the section-scoped one).

    Child of `p14:section` that lists the slide-ids of the slides assigned to
    the containing section, in presentation order.
    """

    sldId_lst: list[CT_SectionSlideId]

    _add_sldId: Callable[..., CT_SectionSlideId]
    sldId = ZeroOrMore("p14:sldId")

    def add_sldId(self, slide_id: int) -> CT_SectionSlideId:
        """Append a new `p14:sldId` child whose `id` attribute is `slide_id`."""
        return self._add_sldId(id=slide_id)


class CT_SectionSlideId(BaseOxmlElement):
    """`p14:sldId` element, child of `p14:sldIdLst` (the section-scoped list).

    References a slide by the value of its corresponding `p:sldId/@id`
    attribute in the main `p:sldIdLst`.
    """

    id: int = RequiredAttribute("id", ST_SlideId)  # pyright: ignore[reportAssignmentType]
