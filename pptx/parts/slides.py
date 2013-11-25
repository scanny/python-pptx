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
    def __init__(self, partname, content_type, element):
        super(_BaseSlide, self).__init__(partname, content_type)
        self._element = element

    @property
    def blob(self):
        return serialize_part_xml(self._element)

    @classmethod
    def load(cls, partname, content_type, blob):
        slide_elm = parse_xml_bytes(blob)
        slide = cls(partname, content_type, slide_elm)
        return slide

    @property
    def name(self):
        """Internal name of this slide-like object."""
        cSld = self._element.cSld
        return cSld.get('name', default='')

    @property
    def shapes(self):
        """
        Collection of shape objects belonging to this slide.
        """
        if not hasattr(self, '_shapes'):
            self._shapes = ShapeCollection(self._element.cSld.spTree, self)
        return self._shapes

    def _add_image(self, file):
        """
        Return a tuple ``(image, relationship)`` representing the |Image| part
        specified by *file*. If a matching image part already exists it is
        reused. If the slide already has a relationship to an existing image,
        that relationship is reused.
        """
        image = self._package._images.add_image(file)
        rel = self._rels.get_or_add(RT.IMAGE, image)
        return (image, rel)

    @property
    def _package(self):
        """Reference to |Package| containing this slide"""
        # !!! --- GET RID OF THIS, PASS PACKAGE TO PART ON CONSTRUCTION !!!
        from pptx.presentation import Package
        # !!! =============================================================
        return Package.containing(self)


class Slide(_BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    def __init__(self, partname, content_type, element):
        super(Slide, self).__init__(partname, content_type, element)

    @classmethod
    def new(cls, slidelayout, partname):
        """
        Return a new slide based on *slidelayout* and having *partname*,
        created from scratch.
        """
        slide_elm = cls._minimal_element()
        slide = cls(partname, CT.PML_SLIDE, slide_elm)
        slide._slidelayout = slidelayout
        slide.shapes._clone_layout_placeholders(slidelayout)
        slide._rels.get_or_add(RT.SLIDE_LAYOUT, slidelayout)
        return slide

    def after_unmarshal(self):
        # selectively unmarshal relationships for now
        self._slidelayout = self._rels.part_with_reltype(RT.SLIDE_LAYOUT)

    @property
    def slidelayout(self):
        """
        |SlideLayout| object this slide inherits appearance from.
        """
        return self._slidelayout

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
    def __init__(self, sldIdLst, prs_rels, presentation):
        super(SlideCollection, self).__init__()
        self._sldIdLst = sldIdLst
        self._prs_rels = prs_rels
        self._presentation = presentation

    def __getitem__(self, key):
        """
        Provide indexed access, (e.g. 'slides[0]').
        """
        if key >= len(self._sldIdLst):
            raise IndexError('slide index out of range')
        sldId = self._sldIdLst[key]
        return self._slide_from_sldId(sldId)

    def __iter__(self):
        """
        Support iteration (e.g. 'for slide in slides:').
        """
        return self._slides

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldIdLst)

    def add_slide(self, slidelayout):
        """
        Return a newly added slide that inherits layout from *slidelayout*.
        """
        temp_partname = PackURI('/ppt/slides/slide1.xml')
        slide = Slide.new(slidelayout, temp_partname)
        rel = self._presentation._rels.get_or_add(RT.SLIDE, slide)
        self._sldIdLst.add_sldId(rel.rId)
        self._rename_slides()  # assigns partname as side effect
        return slide

    def _rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self._slides):
            partname_str = '/ppt/slides/slide%d.xml' % (idx+1)
            slide.partname = PackURI(partname_str)

    def _slide_from_sldId(self, sldId):
        """
        Return the |Slide| instance referenced by *sldId*.
        """
        return self._prs_rels.part_with_rId(sldId.rId)

    @property
    def _slides(self):
        """
        Iterator over slides in collection.
        """
        for sldId in self._sldIdLst:
            yield self._slide_from_sldId(sldId)


class SlideLayout(_BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    def __init__(self, partname, content_type, element):
        super(SlideLayout, self).__init__(partname, content_type, element)

    def after_unmarshal(self):
        # selectively unmarshal relationships we need
        self._slidemaster = self._rels.part_with_reltype(RT.SLIDE_MASTER)

    @property
    def slidemaster(self):
        """Slide master from which this slide layout inherits properties."""
        assert self._slidemaster is not None, (
            "SlideLayout.slidemaster referenced before assigned"
        )
        return self._slidemaster


class SlideMaster(_BaseSlide):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """
    # TECHNOTE: In the Microsoft API, Master is a general type that all of
    # SlideMaster, SlideLayout (CustomLayout), HandoutMaster, and NotesMaster
    # inherit from. So might look into why that is and consider refactoring
    # the various masters a bit later.
    def __init__(self, partname, content_type, element):
        super(SlideMaster, self).__init__(partname, content_type, element)

    @lazyproperty
    def slidelayouts(self):
        """
        Collection of slide layout objects belonging to this slide master.
        """
        slidelayouts = PartCollection()
        sl_rels = [r for r in self._rels if r.reltype == RT.SLIDE_LAYOUT]
        for sl_rel in sl_rels:
            slide_layout = sl_rel.target_part
            slidelayouts.add_part(slide_layout)
        return slidelayouts
