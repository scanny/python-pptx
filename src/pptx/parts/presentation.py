"""Presentation part, the main part in a .pptx package."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterable

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.parts.font import FontPart
from pptx.parts.slide import NotesMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.opc.serialized import ZipDateTime
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.slide import NotesMaster, Slide, SlideLayout, SlideMaster


class PresentationPart(XmlPart):
    """Top level class in object model.

    Represents the contents of the /ppt directory of a .pptx file.
    """

    def add_slide(self, slide_layout: SlideLayout):
        """Return (rId, slide) pair of a newly created blank slide.

        New slide inherits appearance from `slide_layout`.
        """
        partname = self._next_slide_partname
        slide_layout_part = slide_layout.part
        slide_part = SlidePart.new(partname, self.package, slide_layout_part)
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    def add_embedded_font(self, font_blob: bytes) -> str:
        """Return the rId of a new |FontPart| added from `font_blob`.

        `font_blob` is the raw bytes of a TrueType / OpenType font file. A new
        |FontPart| is added to the package (partname
        `/ppt/fonts/font{n}.fntdata`) and related to this presentation part via
        an `http://schemas.openxmlformats.org/officeDocument/2006/relationships/font`
        relationship whose rId is returned.
        """
        font_part = FontPart.new(font_blob, self.package)
        return self.relate_to(font_part, RT.FONT)

    def add_slide_from_external(
        self, source_slide: Slide, slide_layout: SlideLayout
    ) -> tuple[str, Slide]:
        """Return (rId, slide) pair for a slide cloned from `source_slide`.

        The new slide is added to this presentation with appearance inherited
        from `slide_layout` (which must belong to this presentation). The
        returned |Slide| contains a deep copy of the shape tree of
        `source_slide`; image parts are copied into this package (or reused
        when a matching image is already present).

        Raises |NotImplementedError| when the source slide references a part
        type not handled by the basic (F1-independent) copy path -- currently
        anything other than slide-layout, image, hyperlink, or notes-slide
        relationships.
        """
        partname = self._next_slide_partname
        slide_part = SlidePart.clone_from(
            source_slide.part, partname, self.package, slide_layout.part
        )
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    def duplicate_slide(self, source_slide: Slide) -> tuple[str, Slide]:
        """Return (rId, slide) pair for a duplicate of `source_slide` in this presentation.

        `source_slide` must belong to this presentation. The returned |Slide|
        has a deep copy of the source slide's shape tree, is bound to the same
        slide layout as the source, and shares (by relationship reuse) the same
        image / chart / OLE / media / hyperlink targets.

        The caller is responsible for adding a matching `p:sldId` to the
        `p:sldIdLst` at the desired position; this method only creates the
        duplicate slide part and registers a relationship from this
        presentation part to it.
        """
        partname = self._next_slide_partname
        slide_part = SlidePart.clone_within(source_slide.part, partname)
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    @property
    def core_properties(self) -> CorePropertiesPart:
        """A |CoreProperties| object for the presentation.

        Provides read/write access to the Dublin Core properties of this presentation.
        """
        return self.package.core_properties

    def get_slide(self, slide_id: int) -> Slide | None:
        """Return optional related |Slide| object identified by `slide_id`.

        Returns |None| if no slide with `slide_id` is related to this presentation.
        """
        for sldId in self._element.sldIdLst:
            if sldId.id == slide_id:
                return self.related_part(sldId.rId).slide
        return None

    @lazyproperty
    def notes_master(self) -> NotesMaster:
        """
        Return the |NotesMaster| object for this presentation. If the
        presentation does not have a notes master, one is created from
        a default template. The same single instance is returned on each
        call.
        """
        return self.notes_master_part.notes_master

    @lazyproperty
    def notes_master_part(self) -> NotesMasterPart:
        """Return the |NotesMasterPart| object for this presentation.

        If the presentation does not have a notes master, one is created from a default template.
        The same single instance is returned on each call.
        """
        try:
            return self.part_related_by(RT.NOTES_MASTER)
        except KeyError:
            notes_master_part = NotesMasterPart.create_default(self.package)
            self.relate_to(notes_master_part, RT.NOTES_MASTER)
            return notes_master_part

    @lazyproperty
    def presentation(self):
        """
        A |Presentation| object providing access to the content of this
        presentation.
        """
        return Presentation(self._element, self)

    def related_slide(self, rId: str) -> Slide:
        """Return |Slide| object for related |SlidePart| related by `rId`."""
        return self.related_part(rId).slide

    def related_slide_master(self, rId: str) -> SlideMaster:
        """Return |SlideMaster| object for |SlideMasterPart| related by `rId`."""
        return self.related_part(rId).slide_master

    def rename_slide_parts(self, rIds: Iterable[str]):
        """Assign incrementing partnames to the slide parts identified by `rIds`.

        Partnames are like `/ppt/slides/slide9.xml` and are assigned in the order their id appears
        in the `rIds` sequence. The name portion is always `slide`. The number part forms a
        continuous sequence starting at 1 (e.g. 1, 2, ... 10, ...). The extension is always
        `.xml`.
        """
        for idx, rId in enumerate(rIds):
            slide_part = self.related_part(rId)
            slide_part.partname = PackURI("/ppt/slides/slide%d.xml" % (idx + 1))

    def save(
        self,
        path_or_stream: str | IO[bytes],
        zip_date_time: ZipDateTime | None = None,
        password: str | None = None,
    ):
        """Save this presentation package to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object. When `zip_date_time` is provided, every zip-member in the saved
        package is stamped with that fixed last-modified timestamp. When `password` is
        provided, the saved package is encrypted using ECMA-376 Agile Encryption (see
        :meth:`pptx.presentation.Presentation.save`). The two keywords are orthogonal.
        """
        self.package.save(path_or_stream, zip_date_time, password=password)

    def save_flat_xml(self, path_or_stream: str | IO[bytes]) -> None:
        """Save this presentation package to `path_or_stream` as Flat OPC XML.

        Flat OPC is the single-file "XML Presentation" format defined in
        ECMA-376 Part 4. See :meth:`pptx.presentation.Presentation.save_flat_xml`
        for details.
        """
        self.package.save_flat_xml(path_or_stream)

    def slide_id(self, slide_part):
        """Return the slide-id associated with `slide_part`."""
        for sldId in self._element.sldIdLst:
            if self.related_part(sldId.rId) is slide_part:
                return sldId.id
        raise ValueError("matching slide_part not found")

    @property
    def _next_slide_partname(self):
        """Return |PackURI| instance containing next available slide partname."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        partname_str = "/ppt/slides/slide%d.xml" % (len(sldIdLst) + 1)
        return PackURI(partname_str)
