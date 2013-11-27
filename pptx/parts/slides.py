# encoding: utf-8

"""
Slide objects, including Slide and SlideMaster.
"""

from __future__ import absolute_import

from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import Element, serialize_part_xml, SubElement
from pptx.parts.part import PartCollection
from pptx.shapes.shapetree import ShapeCollection
from pptx.util import lazyproperty


class _BaseSlide(Part):
    """
    Base class for slide parts, e.g. slide, slideLayout, slideMaster,
    notesSlide, notesMaster, and handoutMaster.
    """
    def __init__(self, partname, content_type, element, package):
        super(_BaseSlide, self).__init__(
            partname, content_type, package=package
        )
        self._element = element

    @property
    def blob(self):
        return serialize_part_xml(self._element)

    @classmethod
    def load(cls, partname, content_type, blob, package):
        slide_elm = parse_xml_bytes(blob)
        slide = cls(partname, content_type, slide_elm, package)
        return slide

    @property
    def name(self):
        """Internal name of this slide-like object."""
        cSld = self._element.cSld
        return cSld.get('name', default='')

    @lazyproperty
    def shapes(self):
        """
        Collection of shape objects belonging to this slide.
        """
        return ShapeCollection(self._element.cSld.spTree, self)

    def _add_image(self, img_file):
        """
        Return 2-tuple ``(image, rId)`` representing an |Image| part
        corresponding to the image in *img_file*, newly created if no
        matching image part is already present. If the slide already has a
        relationship to an existing image, that relationship is reused.
        """
        image = self._package._images.add_image(img_file)
        rId = self.relate_to(image, RT.IMAGE)
        return (image, rId)


class Slide(_BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    @classmethod
    def new(cls, slidelayout, partname, package):
        """
        Return a new slide based on *slidelayout* and having *partname*,
        created from scratch.
        """
        slide_elm = cls._minimal_element()
        slide = cls(partname, CT.PML_SLIDE, slide_elm, package)
        slide.shapes._clone_layout_placeholders(slidelayout)
        slide.relate_to(slidelayout, RT.SLIDE_LAYOUT)
        return slide

    @property
    def slidelayout(self):
        """
        |SlideLayout| object this slide inherits appearance from.
        """
        return self.part_related_by(RT.SLIDE_LAYOUT)

    @staticmethod
    def _minimal_element():
        """
        Return element containing the minimal XML for a slide, based on what
        is required by the XMLSchema.
        """
        sld = Element('p:sld')
        SubElement(sld, 'p:cSld')
        SubElement(sld.cSld, 'p:spTree')
        SubElement(sld.cSld.spTree, 'p:nvGrpSpPr')
        SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvPr')
        SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvGrpSpPr')
        SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:nvPr')
        SubElement(sld.cSld.spTree, 'p:grpSpPr')
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('id', '1')
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('name', '')
        return sld


class SlideCollection(object):
    """
    Immutable sequence of slides belonging to an instance of |Presentation|,
    with methods for manipulating the slides in the presentation.
    """
    def __init__(self, sldIdLst, prs):
        super(SlideCollection, self).__init__()
        self._sldIdLst = sldIdLst
        self._prs = prs

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. 'slides[0]').
        """
        if idx >= len(self._sldIdLst):
            raise IndexError('slide index out of range')
        rId = self._sldIdLst[idx].rId
        return self._prs.related_parts[rId]

    def __iter__(self):
        """
        Support iteration (e.g. 'for slide in slides:').
        """
        for sldId in self._sldIdLst:
            rId = sldId.rId
            yield self._prs.related_parts[rId]

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldIdLst)

    def add_slide(self, slidelayout):
        """
        Return a newly added slide that inherits layout from *slidelayout*.
        """
        partname = self._next_partname
        package = self._prs.package
        slide = Slide.new(slidelayout, partname, package)
        rId = self._prs.relate_to(slide, RT.SLIDE)
        self._sldIdLst.add_sldId(rId)
        return slide

    def rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self):
            partname_str = '/ppt/slides/slide%d.xml' % (idx+1)
            slide.partname = PackURI(partname_str)

    @property
    def _next_partname(self):
        """
        Return |PackURI| instance containing the partname for a slide to be
        appended to this slide collection, e.g. ``/ppt/slides/slide9.xml``
        for a slide collection containing 8 slides.
        """
        partname_str = '/ppt/slides/slide%d.xml' % (len(self)+1)
        return PackURI(partname_str)


class SlideLayout(_BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    @property
    def slidemaster(self):
        """
        Slide master from which this slide layout inherits properties.
        """
        return self.part_related_by(RT.SLIDE_MASTER)


class SlideMaster(_BaseSlide):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """
    # TECHNOTE: In the Microsoft API, Master is a general type that all of
    # SlideMaster, SlideLayout (CustomLayout), HandoutMaster, and NotesMaster
    # inherit from. So might look into why that is and consider refactoring
    # the various masters a bit later.
    @lazyproperty
    def slidelayouts(self):
        """
        Collection of slide layout objects belonging to this slide master.
        """
        slidelayouts = PartCollection()
        sl_rels = [r for r in self.rels if r.reltype == RT.SLIDE_LAYOUT]
        for sl_rel in sl_rels:
            slide_layout = sl_rel.target_part
            slidelayouts.add_part(slide_layout)
        return slidelayouts
