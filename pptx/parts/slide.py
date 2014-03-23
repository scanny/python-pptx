# encoding: utf-8

"""
Slide objects, including Slide and SlideMaster.
"""

from __future__ import absolute_import

from warnings import warn

from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import Element, SubElement, qn
from pptx.oxml.ns import nsmap, _nsmap
from pptx.shapes.placeholder import BasePlaceholder
from pptx.shapes.shapetree import (
    BaseShapeFactory, BaseShapeTree, ShapeCollection
)
from pptx.util import lazyproperty


class BaseSlide(Part):
    """
    Base class for slide parts, e.g. slide, slideLayout, slideMaster,
    notesSlide, notesMaster, and handoutMaster.
    """
    def __init__(self, partname, content_type, element, package):
        super(BaseSlide, self).__init__(
            partname, content_type, element=element, package=package
        )

    @classmethod
    def load(cls, partname, content_type, blob, package):
        slide_elm = parse_xml_bytes(blob)
        slide = cls(partname, content_type, slide_elm, package)
        return slide

    @property
    def name(self):
        """
        Internal name of this slide.
        """
        cSld = self._element.cSld
        return cSld.get('name', default='')

    @property
    def part(self):
        """
        Part of the parent protocol, "children" of the slide will not know
        the part that contains them so must ask their parent object. That
        chain of delegation ends here for slide child objects.
        """
        return self

    @property
    def spTree(self):
        """
        Reference to ``<p:spTree>`` element for this slide
        """
        spTree_lst = self._element.xpath(
            './p:cSld/p:spTree', namespaces=_nsmap
        )
        return spTree_lst[0]

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


class Slide(BaseSlide):
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

    @lazyproperty
    def shapes(self):
        """
        Instance of |ShapeCollection| containing sequence of shape objects
        appearing on this slide.
        """
        return ShapeCollection(self._element.cSld.spTree, self)

    @lazyproperty
    def shapes_new(self):
        """
        Instance of |ShapeTree| containing sequence of shape objects
        appearing on this slide.
        """
        return _SlideShapeTree(self)

    @property
    def slide_layout(self):
        """
        |SlideLayout| object this slide inherits appearance from.
        """
        return self.part_related_by(RT.SLIDE_LAYOUT)

    @property
    def slidelayout(self):
        """
        Deprecated. Use ``.slide_layout`` property instead.
        """
        msg = (
            'Slide.slidelayout property is deprecated. Use .slide_layout '
            'instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_layout

    @staticmethod
    def _minimal_element():
        """
        Return element containing the minimal XML for a slide, based on what
        is required by the XMLSchema.
        """
        sld = Element('p:sld', nsmap('a', 'p', 'r'))
        SubElement(sld, 'p:cSld')
        SubElement(sld.cSld, 'p:spTree')
        SubElement(sld.cSld.spTree, 'p:nvGrpSpPr')
        SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvPr')
        SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvGrpSpPr')
        SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:nvPr')
        SubElement(sld.cSld.spTree, 'p:grpSpPr')
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('id', '1')
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('name', '')
        SubElement(sld, 'p:clrMapOvr')
        SubElement(sld.clrMapOvr, 'a:masterClrMapping')
        return sld


class SlideCollection(object):
    """
    Sequence of slides belonging to an instance of |Presentation|, having list
    semantics for access to individual slides. Supports indexed access,
    len(), and iteration.
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


class _SlideShapeTree(BaseShapeTree):
    """
    Sequence of shapes appearing on a slide. The first shape in the sequence
    is the backmost in z-order and the last shape is topmost. Supports indexed
    access, len(), index(), and iteration.
    """
    def _clone_layout_placeholders(self, slidelayout):
        """
        Add placeholder shapes based on those in *slidelayout*. Z-order of
        placeholders is preserved. Latent placeholders (date, slide number,
        and footer) are not cloned.
        """
        raise NotImplementedError

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return _SlideShapeFactory(shape_elm, self)


def _SlideShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return _SlidePlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class _SlidePlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide. Inherits shape properties from its
    corresponding slide layout placeholder.
    """
