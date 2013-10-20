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

import os
import weakref

import pptx.packaging

from pptx.exceptions import InvalidPackageError
from pptx.opc_constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml import _Element, qn
from pptx.parts.coreprops import _CoreProperties
from pptx.parts.image import _Image, _ImageCollection
from pptx.parts.part import _BasePart, _PartCollection
from pptx.parts.slides import (
    _Slide, _SlideCollection, _SlideLayout, _SlideMaster
)
from pptx.rels import _Relationship, _RelationshipCollection
from pptx.spec import namespaces

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
