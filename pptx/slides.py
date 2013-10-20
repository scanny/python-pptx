# encoding: utf-8

"""
Slide objects, including _Slide and _SlideMaster.
"""

from pptx.opc_constants import RELATIONSHIP_TYPE as RT
from pptx.part import _BasePart
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
