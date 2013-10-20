# -*- coding: utf-8 -*-
#
# presentation.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
API classes for dealing with presentations and other objects one typically
encounters as an end-user of the PowerPoint user interface.
"""

import hashlib

try:
    from PIL import Image as PIL_Image
except ImportError:
    import Image as PIL_Image

import os
import posixpath
import weakref

from datetime import datetime
from StringIO import StringIO

import pptx.packaging
import pptx.spec as spec

from pptx.exceptions import InvalidPackageError
from pptx.opc_constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml import CT_CoreProperties, _Element, _SubElement, qn
from pptx.part import _BasePart, _PartCollection
from pptx.rels import _Relationship, _RelationshipCollection
from pptx.shapes.shapetree import _ShapeCollection
from pptx.spec import namespaces
from pptx.util import Px

# default namespace map for use in lxml calls
_nsmap = namespaces('a', 'r', 'p')


def _child(element, child_tagname, nsmap=None):
    """
    Return direct child of *element* having *child_tagname* or |None| if no
    such child element is present.
    """
    # use default nsmap if not specified
    if nsmap is None:
        nsmap = _nsmap
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


# ============================================================================
# _Package
# ============================================================================

class _Package(object):
    """
    Return an instance of |_Package| loaded from *file*, where *file* can be a
    path (a string) or a file-like object. If *file* is a path, it can be
    either a path to a PowerPoint `.pptx` file or a path to a directory
    containing an expanded presentation file, as would result from unzipping
    a `.pptx` file. If *file* is |None|, the default presentation template is
    loaded.
    """
    # track instances as weakrefs so .containing() can be computed
    __instances = []

    def __init__(self, file_=None):
        super(_Package, self).__init__()
        self.__presentation = None
        self.__core_properties = None
        self.__relationships = _RelationshipCollection()
        self.__images = _ImageCollection()
        self.__instances.append(weakref.ref(self))
        if file_ is None:
            file_ = self.__default_pptx_path
        self.__open(file_)

    @classmethod
    def containing(cls, part):
        """Return package instance that contains *part*"""
        for pkg in cls.instances():
            if part in pkg._parts:
                return pkg
        raise KeyError("No package contains part %r" % part)

    @property
    def core_properties(self):
        """
        Instance of |_CoreProperties| holding the read/write Dublin Core
        document properties for this presentation.
        """
        assert self.__core_properties, ('_Package.__core_properties referenc'
                                        'ed before assigned')
        return self.__core_properties

    @classmethod
    def instances(cls):
        """Return tuple of _Package instances that have been created"""
        # clean garbage collected pkgs out of __instances
        cls.__instances[:] = [wkref for wkref in cls.__instances
                              if wkref() is not None]
        # return instance references in a tuple
        pkgs = [wkref() for wkref in cls.__instances]
        return tuple(pkgs)

    @property
    def presentation(self):
        """
        Reference to the |Presentation| instance contained in this package.
        """
        return self.__presentation

    def save(self, file):
        """
        Save this package to *file*, where *file* can be either a path to a
        file (a string) or a file-like object.
        """
        pkgng_pkg = pptx.packaging.Package().marshal(self)
        pkgng_pkg.save(file)

    @property
    def _images(self):
        return self.__images

    @property
    def _relationships(self):
        return self.__relationships

    def __load(self, pkgrels):
        """
        Load all the model-side parts and relationships from the on-disk
        package by loading package-level relationship parts and propagating
        the load down the relationship graph.
        """
        # keep track of which parts are already loaded
        part_dict = {}

        # discard any previously loaded relationships
        self.__relationships = _RelationshipCollection()

        # add model-side rel for each pkg-side one, and load target parts
        for pkgrel in pkgrels:
            # unpack working values for part to be loaded
            reltype = pkgrel.reltype
            pkgpart = pkgrel.target
            partname = pkgpart.partname
            content_type = pkgpart.content_type

            # create target part
            part = _Part(reltype, content_type)
            part_dict[partname] = part
            part._load(pkgpart, part_dict)

            # create model-side package relationship
            model_rel = _Relationship(pkgrel.rId, reltype, part)
            self.__relationships._additem(model_rel)

        # gather references to image parts into __images
        self.__images = _ImageCollection()
        image_parts = [part for part in self._parts
                       if part.__class__.__name__ == '_Image']
        for image in image_parts:
            self.__images._loadpart(image)

    def __open(self, file):
        """
        Load presentation contained in *file* into this package.
        """
        pkg = pptx.packaging.Package().open(file)
        self.__load(pkg.relationships)
        # unmarshal relationships selectively for now
        for rel in self.__relationships:
            if rel._reltype == RT.OFFICE_DOCUMENT:
                self.__presentation = rel._target
            elif rel._reltype == RT.CORE_PROPERTIES:
                self.__core_properties = rel._target
        if self.__core_properties is None:
            core_props = _CoreProperties._default()
            self.__core_properties = core_props
            rId = self.__relationships._next_rId
            rel = _Relationship(rId, RT.CORE_PROPERTIES, core_props)
            self.__relationships._additem(rel)

    @property
    def __default_pptx_path(self):
        """
        The path of the default presentation, used when no path is specified
        on construction.
        """
        thisdir = os.path.split(__file__)[0]
        return os.path.join(thisdir, 'templates', 'default.pptx')

    @property
    def _parts(self):
        """
        Return a list containing a reference to each of the parts in this
        package.
        """
        return [part for part in _Package.__walkparts(self.__relationships)]

    @staticmethod
    def __walkparts(rels, parts=None):
        """
        Recursive function, walk relationships to iterate over all parts in
        this package. Leave out *parts* parameter in call to visit all parts.
        """
        # initial call can leave out parts parameter as a signal to initialize
        if parts is None:
            parts = []
        for rel in rels:
            part = rel._target
            # only visit each part once (graph is cyclic)
            if part in parts:
                continue
            parts.append(part)
            yield part
            for part in _Package.__walkparts(part._relationships, parts):
                yield part


# ============================================================================
# Parts
# ============================================================================

class _ImageCollection(_PartCollection):
    """
    Immutable sequence of images, typically belonging to an instance of
    |_Package|. An image part containing a particular image blob appears only
    once in an instance, regardless of how many times it is referenced by a
    pic shape in a slide.
    """
    def __init__(self):
        super(_ImageCollection, self).__init__()

    def add_image(self, file):
        """
        Return image part containing the image in *file*, which is either a
        path to an image file or a file-like object containing an image. If an
        image instance containing this same image already exists, that
        instance is returned. If it does not yet exist, a new one is created.
        """
        # use _Image constructor to validate and characterize image file
        image = _Image(file)
        # return matching image if found
        for existing_image in self._values:
            if existing_image._sha1 == image._sha1:
                return existing_image
        # otherwise add it to collection and return new image
        self._values.append(image)
        self.__rename_images()
        return image

    def __rename_images(self):
        """
        Assign partnames like ``/ppt/media/image9.png`` to all images in the
        collection. The name portion is always ``image``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is preserved during renaming.
        """
        for idx, image in enumerate(self._values):
            image.partname = '/ppt/media/image%d%s' % (idx+1, image.ext)


class _Part(object):
    """
    Part factory. Returns an instance of the appropriate custom part type for
    part types that have them, _BasePart otherwise.
    """
    def __new__(cls, reltype, content_type):
        """
        *reltype* is the relationship type, e.g. ``RT.SLIDE``, corresponding
        to the type of part to be created. For at least one part type, in
        particular the presentation part type, *content_type* is also required
        in order to fully specify the part to be created.
        """
        PRS_MAIN_CONTENT_TYPES = (
            CT.PML_PRESENTATION_MAIN, CT.PML_TEMPLATE_MAIN,
            CT.PML_SLIDESHOW_MAIN
        )
        if reltype == RT.CORE_PROPERTIES:
            return _CoreProperties()
        if reltype == RT.OFFICE_DOCUMENT:
            if content_type in PRS_MAIN_CONTENT_TYPES:
                return Presentation()
            else:
                tmpl = "Not a presentation content type, got '%s'"
                raise InvalidPackageError(tmpl % content_type)
        elif reltype == RT.SLIDE:
            return _Slide()
        elif reltype == RT.SLIDE_LAYOUT:
            return _SlideLayout()
        elif reltype == RT.SLIDE_MASTER:
            return _SlideMaster()
        elif reltype == RT.IMAGE:
            return _Image()
        return _BasePart()


class _CoreProperties(_BasePart):
    """
    Corresponds to part named ``/docProps/core.xml``, containing the core
    document properties for this document package.
    """
    _propnames = (
        'author', 'category', 'comments', 'content_status', 'created',
        'identifier', 'keywords', 'language', 'last_modified_by',
        'last_printed', 'modified', 'revision', 'subject', 'title', 'version'
    )

    def __init__(self, partname=None):
        super(_CoreProperties, self).__init__(CT.OPC_CORE_PROPERTIES, partname)

    @classmethod
    def _default(cls):
        core_props = _CoreProperties('/docProps/core.xml')
        core_props._element = CT_CoreProperties.new_coreProperties()
        core_props.title = 'PowerPoint Presentation'
        core_props.last_modified_by = 'python-pptx'
        core_props.revision = 1
        core_props.modified = datetime.utcnow()
        return core_props

    def __getattribute__(self, name):
        """
        Intercept attribute access to generalize property getters.
        """
        if name in _CoreProperties._propnames:
            return getattr(self._element, name)
        else:
            return super(_CoreProperties, self).__getattribute__(name)

    def __setattr__(self, name, value):
        """
        Intercept attribute assignment to generalize assignment to properties
        """
        if name in _CoreProperties._propnames:
            setattr(self._element, name, value)
        else:
            super(_CoreProperties, self).__setattr__(name, value)


class Presentation(_BasePart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    def __init__(self):
        super(Presentation, self).__init__()
        self.__slidemasters = _PartCollection()
        self.__slides = _SlideCollection(self)

    @property
    def slidemasters(self):
        """
        List of |_SlideMaster| objects belonging to this presentation.
        """
        return tuple(self.__slidemasters)

    @property
    def slides(self):
        """
        |_SlideCollection| object containing the slides in this presentation.
        """
        return self.__slides

    @property
    def _blob(self):
        """
        Rewrite sldId elements in sldIdLst before handing over to super for
        transformation of _element into a blob.
        """
        self.__rewrite_sldIdLst()
        # # at least the following needs to be added before using
        # # _reltype_ordering again for Presentation
        # self.__rewrite_notesMasterIdLst()
        # self.__rewrite_handoutMasterIdLst()
        # self.__rewrite_sldMasterIdLst()
        return super(Presentation, self)._blob

    def _load(self, pkgpart, part_dict):
        """
        Load presentation from package part.
        """
        # call parent to do generic aspects of load
        super(Presentation, self)._load(pkgpart, part_dict)

        # side effect of setting reltype ordering is that rId values can be
        # changed (renumbered during resequencing), so must complete rewrites
        # of all four IdLst elements (notesMasterIdLst, etc.) internal to
        # presentation.xml to reflect any possible changes. Not sure if good
        # order in the .rels files is worth the trouble just yet, so
        # commenting this out for now.

        # # set reltype ordering so rels file ordering is readable
        # self._relationships._reltype_ordering = (RT.SLIDE_MASTER,
        #     RT.NOTES_MASTER, RT.HANDOUT_MASTER, RT.SLIDE,
        #     RT.PRES_PROPS, RT.VIEW_PROPS, RT.TABLE_STYLES, RT.THEME)

        # selectively unmarshal relationships for now
        for rel in self._relationships:
            if rel._reltype == RT.SLIDE_MASTER:
                self.__slidemasters._loadpart(rel._target)
            elif rel._reltype == RT.SLIDE:
                self.__slides._loadpart(rel._target)
        return self

    def __rewrite_sldIdLst(self):
        """
        Rewrite the ``<p:sldIdLst>`` element in ``<p:presentation>`` to
        reflect current ordering of slide relationships and possible
        renumbering of ``rId`` values.
        """
        sldIdLst = _child(self._element, 'p:sldIdLst')
        if sldIdLst is None:
            sldIdLst = self.__add_sldIdLst()
        sldIdLst.clear()
        sld_rels = self._relationships.rels_of_reltype(RT.SLIDE)
        for idx, rel in enumerate(sld_rels):
            sldId = _Element('p:sldId')
            sldIdLst.append(sldId)
            sldId.set('id', str(256+idx))
            sldId.set(qn('r:id'), rel._rId)

    def __add_sldIdLst(self):
        """
        Add a <p:sldIdLst> element to <p:presentation> in the right sequence
        among its siblings.
        """
        sldIdLst = _child(self._element, 'p:sldIdLst')
        assert sldIdLst is None, '__add_sldIdLst() called where '\
                                 '<p:sldIdLst> already exists'
        sldIdLst = _Element('p:sldIdLst')
        # insert new sldIdLst element in right sequence
        sldSz = _child(self._element, 'p:sldSz')
        if sldSz is not None:
            sldSz.addprevious(sldIdLst)
        else:
            notesSz = _child(self._element, 'p:notesSz')
            notesSz.addprevious(sldIdLst)
        return sldIdLst


class _Image(_BasePart):
    """
    Return new Image part instance. *file* may be |None|, a path to a file (a
    string), or a file-like object. If *file* is |None|, no image is loaded
    and :meth:`_load` must be called before using the instance. Otherwise, the
    file referenced or contained in *file* is loaded. Corresponds to package
    files ppt/media/image[1-9][0-9]*.*.
    """
    def __init__(self, file=None):
        super(_Image, self).__init__()
        self.__filepath = None
        self.__ext = None
        if file is not None:
            self.__load_image_from_file(file)

    @property
    def ext(self):
        """
        Return file extension for this image. Includes the leading dot, e.g.
        ``'.png'``.
        """
        assert self.__ext, "_Image.__ext referenced before assigned"
        return self.__ext

    @property
    def _desc(self):
        """
        Return filename associated with this image, either the filename of the
        original image file the image was created with or a synthetic name of
        the form ``image.ext`` where ``.ext`` is appropriate to the image file
        format, e.g. ``'.jpg'``.
        """
        if self.__filepath is not None:
            return os.path.split(self.__filepath)[1]
        # return generic filename if original filename is unknown
        return 'image%s' % self.ext

    def _scale(self, width, height):
        """
        Return scaled image dimensions based on supplied parameters. If
        *width* and *height* are both |None|, the native image size is
        returned. If neither *width* nor *height* is |None|, their values are
        returned unchanged. If a value is provided for either *width* or
        *height* and the other is |None|, the dimensions are scaled,
        preserving the image's aspect ratio.
        """
        native_width_px, native_height_px = self._size
        native_width = Px(native_width_px)
        native_height = Px(native_height_px)

        if width is None and height is None:
            width = native_width
            height = native_height
        elif width is None:
            scaling_factor = float(height) / float(native_height)
            width = int(round(native_width * scaling_factor))
        elif height is None:
            scaling_factor = float(width) / float(native_width)
            height = int(round(native_height * scaling_factor))
        return width, height

    @property
    def _sha1(self):
        """Return SHA1 hash digest for image"""
        return hashlib.sha1(self._blob).hexdigest()

    @property
    def _size(self):
        """
        Return *width*, *height* tuple representing native dimensions of
        image in pixels.
        """
        image_stream = StringIO(self._blob)
        width_px, height_px = PIL_Image.open(image_stream).size
        image_stream.close()
        return width_px, height_px

    @property
    def _blob(self):
        """
        For an image, _blob is always _load_blob, image file content is not
        manipulated.
        """
        return self._load_blob

    def _load(self, pkgpart, part_dict):
        """Handle aspects of loading that are particular to image parts."""
        # call parent to do generic aspects of load
        super(_Image, self)._load(pkgpart, part_dict)
        # set file extension
        self.__ext = posixpath.splitext(pkgpart.partname)[1]
        # return self-reference to allow generative calling
        return self

    @staticmethod
    def __image_ext_content_type(ext):
        """Return the content type corresponding to filename extension *ext*"""
        if ext not in spec.default_content_types:
            tmpl = "unsupported image file extension '%s'"
            raise TypeError(tmpl % (ext))
        content_type = spec.default_content_types[ext]
        if not content_type.startswith('image/'):
            tmpl = "'%s' is not an image content type; ext '%s'"
            raise TypeError(tmpl % (content_type, ext))
        return content_type

    @staticmethod
    def __ext_from_image_stream(stream):
        """
        Return the filename extension appropriate to the image file contained
        in *stream*.
        """
        ext_map = {'GIF': '.gif', 'JPEG': '.jpg', 'PNG': '.png',
                   'TIFF': '.tiff', 'WMF': '.wmf'}
        stream.seek(0)
        format = PIL_Image.open(stream).format
        if format not in ext_map:
            tmpl = "unsupported image format, expected one of: %s, got '%s'"
            raise ValueError(tmpl % (ext_map.keys(), format))
        return ext_map[format]

    def __load_image_from_file(self, file):
        """
        Load image from *file*, which is either a path to an image file or a
        file-like object.
        """
        if isinstance(file, basestring):  # file is a path
            self.__filepath = file
            self.__ext = os.path.splitext(self.__filepath)[1]
            self._content_type = self.__image_ext_content_type(self.__ext)
            with open(self.__filepath, 'rb') as f:
                self._load_blob = f.read()
        else:  # assume file is a file-like object
            self.__ext = self.__ext_from_image_stream(file)
            self._content_type = self.__image_ext_content_type(self.__ext)
            file.seek(0)
            self._load_blob = file.read()


# ============================================================================
# Slide Parts
# ============================================================================

class _SlideCollection(_PartCollection):
    """
    Immutable sequence of slides belonging to an instance of |Presentation|,
    with methods for manipulating the slides in the presentation.
    """
    def __init__(self, presentation):
        super(_SlideCollection, self).__init__()
        self.__presentation = presentation

    def add_slide(self, slidelayout):
        """Add a new slide that inherits layout from *slidelayout*."""
        # 1. construct new slide
        slide = _Slide(slidelayout)
        # 2. add it to this collection
        self._values.append(slide)
        # 3. assign its partname
        self.__rename_slides()
        # 4. add presentation->slide relationship
        self.__presentation._add_relationship(RT.SLIDE, slide)
        # 5. return reference to new slide
        return slide

    def __rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self._values):
            slide.partname = '/ppt/slides/slide%d.xml' % (idx+1)


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
        return _Package.containing(self)


class _Slide(_BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    def __init__(self, slidelayout=None):
        super(_Slide, self).__init__(CT.PML_SLIDE)
        self.__slidelayout = slidelayout
        self._element = self.__minimal_element
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
        return self.__slidelayout

    def _load(self, pkgpart, part_dict):
        """
        Load slide from package part.
        """
        # call parent to do generic aspects of load
        super(_Slide, self)._load(pkgpart, part_dict)
        # selectively unmarshal relationships for now
        for rel in self._relationships:
            if rel._reltype == RT.SLIDE_LAYOUT:
                self.__slidelayout = rel._target
        return self

    @property
    def __minimal_element(self):
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
