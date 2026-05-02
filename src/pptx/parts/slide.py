"""Slide and related objects."""

from __future__ import annotations

import copy
import io
from typing import IO, TYPE_CHECKING, cast

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartRelationshipCloner, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide, CT_Slide
from pptx.oxml.theme import CT_OfficeStyleSheet
from pptx.parts.chart import ChartPart
from pptx.parts.comments import CommentsPart
from pptx.parts.embeddedpackage import EmbeddedPackagePart
from pptx.parts.media import MediaPart
from pptx.slide import NotesMaster, NotesSlide, Slide, SlideLayout, SlideMaster
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.media import Video
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.package import Package
    from pptx.media import Audio, Video
    from pptx.parts.image import Image, ImagePart


class BaseSlidePart(XmlPart):
    """Base class for slide parts.

    This includes slide, slide-layout, and slide-master parts, but also notes-slide,
    notes-master, and handout-master parts.
    """

    _element: CT_Slide

    def get_image(self, rId: str) -> Image:
        """Return an |Image| object containing the image related to this slide by *rId*.

        Raises |KeyError| if no image is related by that id, which would generally indicate a
        corrupted .pptx file.
        """
        return cast("ImagePart", self.related_part(rId)).image

    def get_or_add_image_part(self, image_file: str | IO[bytes]):
        """Return `(image_part, rId)` pair corresponding to `image_file`.

        The returned |ImagePart| object contains the image in `image_file` and is
        related to this slide with the key `rId`. If either the image part or
        relationship already exists, they are reused, otherwise they are newly created.
        """
        image_part = self._package.get_or_add_image_part(image_file)
        rId = self.relate_to(image_part, RT.IMAGE)
        return image_part, rId

    @property
    def name(self) -> str:
        """Internal name of this slide."""
        return self._element.cSld.name


class NotesMasterPart(BaseSlidePart):
    """Notes master part.

    Corresponds to package file `ppt/notesMasters/notesMaster1.xml`.
    """

    @classmethod
    def create_default(cls, package):
        """
        Create and return a default notes master part, including creating the
        new theme it requires.
        """
        notes_master_part = cls._new(package)
        theme_part = cls._new_theme_part(package)
        notes_master_part.relate_to(theme_part, RT.THEME)
        return notes_master_part

    @lazyproperty
    def notes_master(self):
        """
        Return the |NotesMaster| object that proxies this notes master part.
        """
        return NotesMaster(self._element, self)

    @classmethod
    def _new(cls, package):
        """
        Create and return a standalone, default notes master part based on
        the built-in template (without any related parts, such as theme).
        """
        return NotesMasterPart(
            PackURI("/ppt/notesMasters/notesMaster1.xml"),
            CT.PML_NOTES_MASTER,
            package,
            CT_NotesMaster.new_default(),
        )

    @classmethod
    def _new_theme_part(cls, package):
        """Return new default theme-part suitable for use with a notes master."""
        return ThemePart(
            package.next_partname("/ppt/theme/theme%d.xml"),
            CT.OFC_THEME,
            package,
            CT_OfficeStyleSheet.new_default(),
        )


class ThemePart(XmlPart):
    """Theme part.

    Corresponds to package files ``ppt/theme/theme[1-9][0-9]*.xml`` and holds the
    color scheme, font scheme, and effect scheme that drive inheritance throughout
    a presentation. Exposed primarily as an anchor for accessors such as
    :attr:`~pptx.oxml.theme.CT_OfficeStyleSheet.hlink_color` (see issue #940).
    """

    _element: CT_OfficeStyleSheet

    @property
    def theme(self) -> CT_OfficeStyleSheet:
        """The `a:theme` XML element proxied by this part."""
        return self._element


class NotesSlidePart(BaseSlidePart):
    """Notes slide part.

    Contains the slide notes content and the layout for the slide handout page.
    Corresponds to package file `ppt/notesSlides/notesSlide[1-9][0-9]*.xml`.
    """

    @classmethod
    def new(cls, package, slide_part):
        """Return new |NotesSlidePart| for the slide in `slide_part`.

        The new notes-slide part is based on the (singleton) notes master and related to
        both the notes-master part and `slide_part`. If no notes-master is present,
        one is created based on the default template.
        """
        notes_master_part = package.presentation_part.notes_master_part
        notes_slide_part = cls._add_notes_slide_part(package, slide_part, notes_master_part)
        notes_slide = notes_slide_part.notes_slide
        notes_slide.clone_master_placeholders(notes_master_part.notes_master)
        return notes_slide_part

    @lazyproperty
    def notes_master(self):
        """Return the |NotesMaster| object this notes slide inherits from."""
        notes_master_part = self.part_related_by(RT.NOTES_MASTER)
        return notes_master_part.notes_master

    @lazyproperty
    def notes_slide(self):
        """Return the |NotesSlide| object that proxies this notes slide part."""
        return NotesSlide(self._element, self)

    @classmethod
    def _add_notes_slide_part(cls, package, slide_part, notes_master_part):
        """Create and return a new notes-slide part.

        The return part is fully related, but has no shape content (i.e. placeholders
        not cloned).
        """
        notes_slide_part = NotesSlidePart(
            package.next_partname("/ppt/notesSlides/notesSlide%d.xml"),
            CT.PML_NOTES_SLIDE,
            package,
            CT_NotesSlide.new(),
        )
        notes_slide_part.relate_to(notes_master_part, RT.NOTES_MASTER)
        notes_slide_part.relate_to(slide_part, RT.SLIDE)
        return notes_slide_part


class SlidePart(BaseSlidePart):
    """Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml."""

    @classmethod
    def new(cls, partname, package, slide_layout_part):
        """Return newly-created blank slide part.

        The new slide-part has `partname` and a relationship to `slide_layout_part`.
        """
        slide_part = cls(partname, CT.PML_SLIDE, package, CT_Slide.new())
        slide_part.relate_to(slide_layout_part, RT.SLIDE_LAYOUT)
        return slide_part

    @classmethod
    def clone_from(
        cls,
        source_slide_part: SlidePart,
        partname: PackURI,
        package: Package,
        slide_layout_part: SlideLayoutPart,
    ) -> SlidePart:
        """Return new |SlidePart| cloned from `source_slide_part` in `package`.

        The new slide part is bound to `slide_layout_part` (which must belong to
        `package`) and its shape tree is a deep copy of the source slide's shape
        tree. Every relationship referenced by the source slide -- image, chart,
        embedded OLE object, media, external hyperlink -- is re-established on
        the new slide part with a fresh target part materialised in `package`
        (or reused when a content-equivalent part already exists, e.g. an
        identical image). The cloned shape-tree's `r:id` / `r:embed` / `r:link`
        attributes are rewritten to match the freshly-allocated relationships.

        This is the full-fidelity cross-presentation slide copy path
        (promoted from the #1036 "basic" variant now that Foundation F1 and F5
        are in place). Charts bring their embedded workbook along as a
        *distinct* :class:`EmbeddedXlsxPart` so PowerPoint's "Edit Data"
        dialog continues to work; images and media are content-deduplicated
        against `package`'s existing parts; and the notes-slide relationship
        (which carries a back-reference to its owning slide) is dropped.
        """
        # -- Start with a deep-copied `p:sld` element. --
        cloned_sld = copy.deepcopy(source_slide_part._element)

        new_slide_part = cls(partname, CT.PML_SLIDE, package, cloned_sld)

        # -- Always relate the new slide to the *target* slide layout. --
        layout_rId = new_slide_part.relate_to(slide_layout_part, RT.SLIDE_LAYOUT)

        # -- Build an rId remap table and re-establish non-layout rels in the --
        # -- target package. Each rel-type gets the handler appropriate for   --
        # -- full-fidelity cross-package transport:                           --
        # --   SLIDE_LAYOUT -> remapped to the supplied target layout         --
        # --   NOTES_SLIDE  -> dropped (notes slides back-reference their     --
        # --                   owning slide, so they can't be shared)         --
        # --   IMAGE        -> deduplicated via package.get_or_add_image_part --
        # --   MEDIA/VIDEO  -> deduplicated via package.get_or_add_media_part --
        # --   /AUDIO          (shared MediaPart, two rel types per video)   --
        # --   CHART        -> ChartPart.clone_from (deep; distinct xlsx)    --
        # --   external     -> URI verbatim                                  --
        # --   other (OLE,  -> F1 shallow clone (new part, same blob)        --
        # --    PACKAGE,                                                     --
        # --    comments…)                                                   --
        rId_map: dict[str, str] = {}
        # -- Track already-cloned MediaParts (keyed by source rel target part
        # -- identity) so that a video's MEDIA + VIDEO rel pair resolves to one
        # -- MediaPart in the target package rather than two divergent clones. --
        media_part_cache: dict[int, MediaPart] = {}
        for old_rId, rel in source_slide_part.rels.items():
            if rel.reltype == RT.SLIDE_LAYOUT:
                rId_map[old_rId] = layout_rId
                continue
            if rel.reltype == RT.NOTES_SLIDE:
                # -- Dropped; any stray `r:id` reference will be rewritten to "". --
                rId_map[old_rId] = ""
                continue
            if rel.is_external:
                # -- External rels (e.g. hyperlinks) copy just the target URL. --
                new_rId = new_slide_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                rId_map[old_rId] = new_rId
                continue
            if rel.reltype == RT.IMAGE:
                target_image_part = cls._clone_image_part_into(
                    cast("ImagePart", rel.target_part), package
                )
                new_rId = new_slide_part.relate_to(target_image_part, RT.IMAGE)
                rId_map[old_rId] = new_rId
                continue
            if rel.reltype in (RT.MEDIA, RT.VIDEO, RT.AUDIO):
                src_media_part = cast(MediaPart, rel.target_part)
                tgt_media_part = cls._clone_media_part_into(
                    src_media_part, package, media_part_cache
                )
                new_rId = new_slide_part.relate_to(tgt_media_part, rel.reltype)
                rId_map[old_rId] = new_rId
                continue
            if rel.reltype == RT.CHART:
                src_chart_part = cast(ChartPart, rel.target_part)
                tgt_chart_part = ChartPart.clone_from(src_chart_part, package)
                new_rId = new_slide_part.relate_to(tgt_chart_part, RT.CHART)
                rId_map[old_rId] = new_rId
                continue
            # -- Any remaining internal rel (OLE object, embedded package,
            # -- comments, etc.) gets the F1 shallow-clone treatment: a new
            # -- part is materialised in the target package with the same
            # -- content-type and blob, and a fresh rId is allocated. Parts
            # -- that already belong to the target package are reused. --
            src_target_part = rel.target_part
            if src_target_part.package is package:
                tgt_target_part = src_target_part
            else:
                tgt_target_part = cls._shallow_clone_part_into(src_target_part, package)
            new_rId = new_slide_part.relate_to(tgt_target_part, rel.reltype)
            rId_map[old_rId] = new_rId

        # -- Rewrite every `r:id` / `r:embed` / `r:link` on cloned element. --
        cls._remap_rel_ids(cloned_sld, rId_map)

        return new_slide_part

    @classmethod
    def clone_within(
        cls,
        source_slide_part: SlidePart,
        partname: PackURI,
    ) -> SlidePart:
        """Return new |SlidePart| cloned from `source_slide_part` in the same package.

        Used to implement :meth:`Slides.duplicate` for duplicating a slide
        within its own presentation. The new slide part is bound to the same
        slide layout as the source and its shape tree is a deep copy of the
        source slide's shape tree. Every relationship referenced from the
        source shape tree (image, chart, OLE object, media, hyperlink, ...) is
        re-established on the new slide part using the F1
        |PartRelationshipCloner|, which reuses the existing target parts
        because they already belong to this package.

        The notes-slide relationship (if any) on the source is intentionally
        *dropped* on the duplicate: a notes slide carries a back-reference to
        its owning slide, so having two slides share one notes-slide part
        would be invalid. Users who need notes on the duplicate can create
        them on the returned slide in the usual way.
        """
        package = source_slide_part.package

        # -- Seed the new part with a minimal `p:sld` root; `PartRelationshipCloner.clone`
        # -- below produces the real (rId-remapped) element tree and replaces it. --
        placeholder_sld = CT_Slide.new()
        new_slide_part = cls(partname, CT.PML_SLIDE, package, placeholder_sld)

        # -- Always relate the duplicate to the same slide-layout as the source. --
        slide_layout_part = source_slide_part.part_related_by(RT.SLIDE_LAYOUT)
        new_slide_part.relate_to(slide_layout_part, RT.SLIDE_LAYOUT)

        # -- Delegate per-rId cloning to the F1 PartRelationshipCloner. It deep-copies the
        # -- source `p:sld`, walks every `r:id` / `r:embed` / `r:link` on the copy, creates
        # -- a matching relationship on `new_slide_part` (reusing existing target parts
        # -- because this is a same-package clone), and rewrites the rId attributes in
        # -- place. The returned element is the duplicate's authoritative shape tree. --
        new_slide_part._element = PartRelationshipCloner.clone(
            source_slide_part, new_slide_part, source_slide_part._element
        )

        return new_slide_part

    @staticmethod
    def _clone_image_part_into(source_image_part: ImagePart, package: Package) -> ImagePart:
        """Return an |ImagePart| in `package` mirroring `source_image_part`.

        Re-uses an existing image part in `package` when one with matching
        content is already present; otherwise creates a new one from the source
        part's blob.
        """
        return package.get_or_add_image_part(io.BytesIO(source_image_part.blob))

    @staticmethod
    def _clone_media_part_into(
        source_media_part: MediaPart,
        package: Package,
        cache: dict[int, MediaPart],
    ) -> MediaPart:
        """Return a |MediaPart| in `package` mirroring `source_media_part`.

        A single video yields two relationships on the owning slide -- MEDIA
        and VIDEO -- both pointing at the same |MediaPart|. The `cache` keeps
        both rels pointing at the one clone in the target package so the
        duplicate media part isn't materialised twice.

        Already-in-target-package parts are reused directly; otherwise a
        new |MediaPart| is created with a non-colliding partname and the
        source blob.
        """
        key = id(source_media_part)
        hit = cache.get(key)
        if hit is not None:
            return hit
        if source_media_part.package is package:
            cache[key] = source_media_part
            return source_media_part
        # -- allocate a non-colliding partname in the target package using the
        # -- source's extension (extracted from its partname). --
        src_partname = str(source_media_part.partname)
        dot = src_partname.rfind(".")
        ext = src_partname[dot + 1 :] if dot != -1 else "bin"
        partname = package.next_media_partname(ext)
        cloned = MediaPart(
            partname,
            source_media_part.content_type,
            package,
            source_media_part.blob,
        )
        cache[key] = cloned
        return cloned

    @staticmethod
    def _shallow_clone_part_into(source_part, package):
        """Return a new part in `package` mirroring `source_part` (shallow duplicate).

        Used for OLE objects, embedded packages, and any other slide-owned rel
        targets that don't have a content-aware deep-clone helper. The new
        partname is allocated via the source part's template so it doesn't
        collide with existing parts in `package`. Relationships *within* the
        source part are not recursed into -- the expected use-cases (OLE,
        generic embedded-package) treat the target as an opaque blob.
        """
        from pptx.opc.package import _partname_template_for

        partname_tmpl = _partname_template_for(source_part.partname)
        new_partname = package.next_partname(partname_tmpl)
        return type(source_part)(
            new_partname, source_part.content_type, package, source_part.blob
        )

    @staticmethod
    def _remap_rel_ids(element: BaseOxmlElement, rId_map: dict[str, str]) -> None:
        """Rewrite every relationship-id attribute under `element` per `rId_map`.

        The attributes considered are `r:id`, `r:embed`, and `r:link` -- the
        three ways OOXML references a relationship from within a part's XML
        body. When a mapped value is an empty string the attribute is removed
        (used for dropped-rel types such as notes-slide references).
        """
        for pfx_name, attr_qn in (
            ("r:id", qn("r:id")),
            ("r:embed", qn("r:embed")),
            ("r:link", qn("r:link")),
        ):
            xpath_expr = "descendant-or-self::*[@%s]" % pfx_name
            for elt in cast("list[BaseOxmlElement]", element.xpath(xpath_expr)):
                old_rId = elt.get(attr_qn)
                if old_rId is None or old_rId not in rId_map:
                    continue
                new_rId = rId_map[old_rId]
                if new_rId:
                    elt.set(attr_qn, new_rId)
                else:
                    del elt.attrib[attr_qn]

    def add_chart_part(self, chart_type: XL_CHART_TYPE, chart_data: ChartData):
        """Return str rId of new |ChartPart| object containing chart of `chart_type`.

        The chart depicts `chart_data` and is related to the slide contained in this
        part by `rId`.
        """
        return self.relate_to(ChartPart.new(chart_type, chart_data, self._package), RT.CHART)

    def add_embedded_ole_object_part(
        self,
        prog_id: PROG_ID | str,
        ole_object_file: str | IO[bytes],
        extension: str | None = None,
    ):
        """Return rId of newly-added OLE-object part formed from `ole_object_file`.

        `extension` is an optional file-extension hint (without the leading dot, e.g.
        ``"zip"``) passed through to the |EmbeddedPackagePart| factory to produce a
        readable part-name for generic OLE-object embeds.
        """
        # -- MS-Office "package" file-types (DOCX/PPTX/XLSX members of PROG_ID) get
        # -- the RT.PACKAGE relationship-type; every other case -- including non-Office
        # -- PROG_ID members and arbitrary str progIds -- gets RT.OLE_OBJECT.
        is_office_package = (
            isinstance(prog_id, PROG_ID) and prog_id.is_office_package
        )
        relationship_type = RT.PACKAGE if is_office_package else RT.OLE_OBJECT
        return self.relate_to(
            EmbeddedPackagePart.factory(
                prog_id,
                self._blob_from_file(ole_object_file),
                self._package,
                extension,
            ),
            relationship_type,
        )

    def get_or_add_sound_media_part(self, audio: Audio) -> str:
        """Return rId of AUDIO relationship to a media part containing `audio`.

        A new |MediaPart| object is created if the package does not already contain a
        media part with this audio's binary content (for example when the same sound
        appears on more than one click-action). A single AUDIO relationship to the part
        is created from this slide part and its rId is returned.
        """
        media_part = self._package.get_or_add_media_part(audio)
        return self.relate_to(media_part, RT.AUDIO)

    def get_or_add_video_media_part(self, video: Video) -> tuple[str, str]:
        """Return rIds for media and video relationships to media part.

        A new |MediaPart| object is created if it does not already exist
        (such as would occur if the same video appeared more than once in
         a presentation). Two relationships to the media part are created,
        one each with MEDIA and VIDEO relationship types. The need for two
        appears to be for legacy support for an earlier (pre-Office 2010)
        PowerPoint media embedding strategy.
        """
        media_part = self._package.get_or_add_media_part(video)
        media_rId = self.relate_to(media_part, RT.MEDIA)
        video_rId = self.relate_to(media_part, RT.VIDEO)
        return media_rId, video_rId

    @property
    def comments_part(self) -> CommentsPart | None:
        """Legacy |CommentsPart| for this slide, or None if no comments are attached."""
        try:
            return cast(CommentsPart, self.part_related_by(RT.COMMENTS))
        except KeyError:
            return None

    def get_or_add_comments_part(self) -> CommentsPart:
        """Return this slide's legacy |CommentsPart|, creating one if needed."""
        part = self.comments_part
        if part is not None:
            return part
        partname = self.package.next_partname("/ppt/comments/comment%d.xml")
        part = CommentsPart.new(self.package, partname)
        self.relate_to(part, RT.COMMENTS)
        return part

    @property
    def has_notes_slide(self):
        """
        Return True if this slide has a notes slide, False otherwise. A notes
        slide is created by the :attr:`notes_slide` property when one doesn't
        exist; use this property to test for a notes slide without the
        possible side-effect of creating one.
        """
        try:
            self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            return False
        return True

    @lazyproperty
    def notes_slide(self) -> NotesSlide:
        """The |NotesSlide| instance associated with this slide.

        If the slide does not have a notes slide, a new one is created. The same single instance
        is returned on each call.
        """
        try:
            notes_slide_part = self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            notes_slide_part = self._add_notes_slide_part()
        return notes_slide_part.notes_slide

    @lazyproperty
    def slide(self):
        """
        The |Slide| object representing this slide part.
        """
        return Slide(self._element, self)

    @property
    def slide_id(self) -> int:
        """Return the slide identifier stored in the presentation part for this slide part."""
        presentation_part = self.package.presentation_part
        return presentation_part.slide_id(self)

    @property
    def slide_layout(self) -> SlideLayout:
        """|SlideLayout| object the slide in this part inherits appearance from."""
        slide_layout_part = self.part_related_by(RT.SLIDE_LAYOUT)
        return slide_layout_part.slide_layout

    def _add_notes_slide_part(self):
        """
        Return a newly created |NotesSlidePart| object related to this slide
        part. Caller is responsible for ensuring this slide doesn't already
        have a notes slide part.
        """
        notes_slide_part = NotesSlidePart.new(self.package, self)
        self.relate_to(notes_slide_part, RT.NOTES_SLIDE)
        return notes_slide_part


class SlideLayoutPart(BaseSlidePart):
    """Slide layout part.

    Corresponds to package files ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """

    @lazyproperty
    def slide_layout(self):
        """
        The |SlideLayout| object representing this part.
        """
        return SlideLayout(self._element, self)

    @property
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide layout inherits properties."""
        return self.part_related_by(RT.SLIDE_MASTER).slide_master


class SlideMasterPart(BaseSlidePart):
    """Slide master part.

    Corresponds to package files ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """

    def related_slide_layout(self, rId: str) -> SlideLayout:
        """Return |SlideLayout| related to this slide-master by key `rId`."""
        return self.related_part(rId).slide_layout

    @lazyproperty
    def slide_master(self):
        """
        The |SlideMaster| object representing this part.
        """
        return SlideMaster(self._element, self)
