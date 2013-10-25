# encoding: utf-8

"""
Slide objects, including _Slide and _SlideMaster.
"""

from __future__ import absolute_import

from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml import _Element, _SubElement
from pptx.parts.part import _BasePart, _PartCollection
from pptx.shapes.shapetree import _ShapeCollection


class _BaseSlide(_BasePart):
    """
    Base class for slide parts, e.g. slide, slideLayout, slideMaster,
    notesSlide, notesMaster, and handoutMaster.
    """
    def __init__(self, content_type=None):
        """
        Needs content_type parameter so newly created parts (not loaded from
        package) can register their content type.
        """
        super(_BaseSlide, self).__init__(content_type)
        self._shapes = None

    @property
    def name(self):
        """Internal name of this slide-like object."""
        cSld = self._element.cSld
        return cSld.get('name', default='')

    @property
    def shapes(self):
        """Collection of shape objects belonging to this slide."""
        assert self._shapes is not None, ("_BaseSlide.shapes referenced "
                                          "before assigned")
        return self._shapes

    def _add_image(self, file):
        """
        Return a tuple ``(image, relationship)`` representing the |Image| part
        specified by *file*. If a matching image part already exists it is
        reused. If the slide already has a relationship to an existing image,
        that relationship is reused.
        """
        image = self._package._images.add_image(file)
        rel = self._add_relationship(RT.IMAGE, image)
        return (image, rel)

    def _load(self, pkgpart, part_dict):
        """Handle aspects of loading that are general to slide types."""
        # call parent to do generic aspects of load
        super(_BaseSlide, self)._load(pkgpart, part_dict)
        # unmarshal shapes
        self._shapes = _ShapeCollection(self._element.cSld.spTree, self)
        # return self-reference to allow generative calling
        return self

    @property
    def _package(self):
        """Reference to |_Package| containing this slide"""
        # !!! --- GET RID OF THIS, PASS PACKAGE TO PART ON CONSTRUCTION !!!
        from pptx.presentation import _Package
        # !!! =============================================================
        return _Package.containing(self)


class _Slide(_BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    def __init__(self, slidelayout=None):
        super(_Slide, self).__init__(CT.PML_SLIDE)
        self._slidelayout = slidelayout
        self._element = self._minimal_element
        self._shapes = _ShapeCollection(self._element.cSld.spTree, self)
        # if slidelayout, this is a slide being added, not one being loaded
        if slidelayout:
            self._shapes._clone_layout_placeholders(slidelayout)
            # add relationship to slideLayout part
            self._add_relationship(RT.SLIDE_LAYOUT, slidelayout)

    @property
    def slidelayout(self):
        """
        |_SlideLayout| object this slide inherits appearance from.
        """
        return self._slidelayout

    def _load(self, pkgpart, part_dict):
        """
        Load slide from package part.
        """
        # call parent to do generic aspects of load
        super(_Slide, self)._load(pkgpart, part_dict)
        # selectively unmarshal relationships for now
        for rel in self._relationships:
            if rel._reltype == RT.SLIDE_LAYOUT:
                self._slidelayout = rel._target
        return self

    @property
    def _minimal_element(self):
        """
        Return element containing the minimal XML for a slide, based on what
        is required by the XMLSchema.
        """
        sld = _Element('p:sld')
        _SubElement(sld, 'p:cSld')
        _SubElement(sld.cSld, 'p:spTree')
        _SubElement(sld.cSld.spTree, 'p:nvGrpSpPr')
        _SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvPr')
        _SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvGrpSpPr')
        _SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:nvPr')
        _SubElement(sld.cSld.spTree, 'p:grpSpPr')
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('id', '1')
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('name', '')
        return sld


class _SlideCollection(_PartCollection):
    """
    Immutable sequence of slides belonging to an instance of |Presentation|,
    with methods for manipulating the slides in the presentation.
    """
    def __init__(self, presentation):
        super(_SlideCollection, self).__init__()
        self._presentation = presentation

    def add_slide(self, slidelayout):
        """Add a new slide that inherits layout from *slidelayout*."""
        # 1. construct new slide
        slide = _Slide(slidelayout)
        # 2. add it to this collection
        self._values.append(slide)
        # 3. assign its partname
        self._rename_slides()
        # 4. add presentation->slide relationship
        self._presentation._add_relationship(RT.SLIDE, slide)
        # 5. return reference to new slide
        return slide

    def _rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self._values):
            slide.partname = '/ppt/slides/slide%d.xml' % (idx+1)


class _SlideLayout(_BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    def __init__(self):
        super(_SlideLayout, self).__init__(CT.PML_SLIDE_LAYOUT)
        self.__slidemaster = None

    @property
    def slidemaster(self):
        """Slide master from which this slide layout inherits properties."""
        assert self.__slidemaster is not None, ("_SlideLayout.slidemaster "
                                                "referenced before assigned")
        return self.__slidemaster

    def _load(self, pkgpart, part_dict):
        """
        Load slide layout from package part.
        """
        # call parent to do generic aspects of load
        super(_SlideLayout, self)._load(pkgpart, part_dict)

        # selectively unmarshal relationships we need
        for rel in self._relationships:
            # get slideMaster from which this slideLayout inherits properties
            if rel._reltype == RT.SLIDE_MASTER:
                self.__slidemaster = rel._target

        # return self-reference to allow generative calling
        return self


class _SlideMaster(_BaseSlide):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """
    # TECHNOTE: In the Microsoft API, Master is a general type that all of
    # SlideMaster, SlideLayout (CustomLayout), HandoutMaster, and NotesMaster
    # inherit from. So might look into why that is and consider refactoring
    # the various masters a bit later.

    def __init__(self):
        super(_SlideMaster, self).__init__(CT.PML_SLIDE_MASTER)
        self.__slidelayouts = _PartCollection()

    @property
    def slidelayouts(self):
        """
        Collection of slide layout objects belonging to this slide master.
        """
        return self.__slidelayouts

    def _load(self, pkgpart, part_dict):
        """
        Load slide master from package part.
        """
        # call parent to do generic aspects of load
        super(_SlideMaster, self)._load(pkgpart, part_dict)

        # selectively unmarshal relationships for now
        for rel in self._relationships:
            if rel._reltype == RT.SLIDE_LAYOUT:
                self.__slidelayouts._loadpart(rel._target)
        return self
