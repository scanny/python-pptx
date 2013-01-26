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
import weakref

from lxml import etree

import pptx.packaging
import pptx.spec as spec
import pptx.util as util

from pptx.exceptions import InvalidPackageError

from pptx.spec import namespaces, qname
from pptx.spec import (CT_PRESENTATION, CT_SLIDE, CT_SLIDELAYOUT,
    CT_SLIDEMASTER, CT_SLIDESHOW, CT_TEMPLATE)
from pptx.spec import (RT_HANDOUTMASTER, RT_IMAGE, RT_NOTESMASTER,
    RT_OFFICEDOCUMENT, RT_PRESPROPS, RT_SLIDE, RT_SLIDELAYOUT, RT_SLIDEMASTER,
    RT_TABLESTYLES, RT_THEME, RT_VIEWPROPS)
from pptx.spec import (PH_TYPE_BODY, PH_TYPE_CTRTITLE, PH_TYPE_DT,
    PH_TYPE_FTR, PH_TYPE_OBJ, PH_TYPE_SLDNUM, PH_TYPE_SUBTITLE, PH_TYPE_TBL,
    PH_TYPE_TITLE, PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL)
from pptx.spec import slide_ph_basenames

from pptx.util import Px

import logging
log = logging.getLogger('pptx.presentation')
log.setLevel(logging.DEBUG)
# log.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - '
                              '%(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)


def _child(element, child_tagname, nsmap):
    """
    Return direct child of *element* having *child_tagname* or :class:`None`
    if no such child element is present.
    """
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


# ============================================================================
# Package
# ============================================================================

class Package(object):
    """
    Root class of presentation object hierarchy.
    """
    # track instances as weakrefs so .containing() can be computed
    __instances = []
    
    def __init__(self):
        super(Package, self).__init__()
        self.__presentation = None
        self.__relationships = _RelationshipCollection()
        self.__images = ImageCollection()
        self.__instances.append(weakref.ref(self))
    
    @classmethod
    def containing(cls, part):
        """Return package instance that contains *part*"""
        for pkg in cls.instances():
            if part in pkg._parts:
                return pkg
        raise KeyError("No package contains part %d" % part)
    
    @classmethod
    def instances(cls):
        """Return tuple of Package instances that have been created"""
        # clean garbage collected pkgs out of __instances
        cls.__instances[:] = [wkref for wkref in cls.__instances
                              if wkref() is not None]
        # return instance references in a tuple
        pkgs = [wkref() for wkref in cls.__instances]
        return tuple(pkgs)
    
    @property
    def presentation(self):
        return self.__presentation
    
    def open(self, path):
        """
        Open presentation file located at *path*. Returns self-reference to
        allow generative calling structure, e.g.
        ``pkg = Package().open(path)``.
        """
        pkg = pptx.packaging.Package().open(path)
        self.__load(pkg.relationships)
        # unmarshal relationships selectively for now
        for rel in self.__relationships:
            if rel._reltype == RT_OFFICEDOCUMENT:
                self.__presentation = rel._target
        return self
    
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
            # log.debug("%s -- %s", reltype, partname)
            
            # create target part
            part = Part(reltype, content_type)
            part_dict[partname] = part
            part._load(pkgpart, part_dict)
            
            # create model-side package relationship
            model_rel = _Relationship(pkgrel.rId, reltype, part)
            self.__relationships._additem(model_rel)
        
        # gather references to image parts into __images
        self.__images = ImageCollection()
        image_parts = [part for part in self._parts
                       if part.__class__.__name__ == 'Image']
        for image in image_parts:
            self.__images._loadpart(image)
    
    @property
    def _parts(self):
        """
        Return a list containing a reference to each of the parts in this
        package.
        """
        return [part for part in Package.__walkparts(self.__relationships)]
    
    @staticmethod
    def __walkparts(rels, parts=None):
        """
        Recursive function, walk relationships to iterate over all parts in
        this package. Leave out *parts* parameter in call to visit all parts.
        """
        # initial call can leave out parts parameter as a signal to initialize
        if parts is None:
            parts = []
        # log.debug("in __walkparts(), len(parts)==%d", len(parts))
        for rel in rels:
            # log.debug("rel.target.partname==%s", rel.target.partname)
            part = rel._target
            if part in parts: # only visit each part once (graph is cyclic)
                continue
            parts.append(part)
            yield part
            for part in Package.__walkparts(part._relationships, parts):
                yield part
    


# ============================================================================
# Base classes
# ============================================================================

class Collection(object):
    """
    Base class for collection classes. May also be used for part collections
    that don't yet have any custom methods.
    
    Has the following characteristics.:
    
    * Container (implements __contains__)
    * Iterable (implements __iter__)
    * Sized (implements __len__)
    * Sequence (implements __getitem__)
    """
    def __init__(self):
        # log.debug('Collection.__init__() called')
        super(Collection, self).__init__()
        self.__values = []
    
    @property
    def _values(self):
        """Return read-only reference to collection values (list)."""
        return self.__values
    
    def __contains__(self, item):  # __iter__ would do this job by itself
        """Supports 'in' operator (e.g. 'x in collection')."""
        return (item in self.__values)
    
    def __getitem__(self, key):
        """Provides indexed access, (e.g. 'collection[0]')."""
        return self.__values.__getitem__(key)
    
    def __iter__(self):
        """Supports iteration (e.g. 'for x in collection: pass')."""
        return self.__values.__iter__()
    
    def __len__(self):
        """Supports len() function (e.g. 'len(collection) == 1')."""
        return len(self.__values)
    
    def index(self, item):
        """Supports index method (e.g. '[1, 2, 3].index(2) == 1')."""
        return self.__values.index(item)
    

class Observable(object):
    """
    Simple observer pattern mixin. Limitations:
    
    * observers get all message types from subject (Observable), subscription
      is on subject basis, not subject + event_type.
    
    * notifications are oriented toward "value has been updated", which seems
      like only one possible event, could also be something like "load has
      completed" or "preparing to load".
    
    """
    def __init__(self):
        super(Observable, self).__init__()
        self._observers = []
    
    def _notify_observers(self, name, value):
        # value = getattr(self, name)
        for observer in self._observers:
            observer.notify(self, name, value)
    
    def add_observer(self, observer):
        """
        Begin notifying *observer* of events. *observer* must implement method
        ``notify(observed, name, new_value)``
        """
        if observer not in self._observers:
            self._observers.append(observer)
    
    def remove_observer(self, observer):
        """Remove *observer* from notification list."""
        assert observer in self._observers, "remove_observer called for"\
                                            "unsubscribed object"
        self._observers.remove(observer)
    


# ============================================================================
# Relationships
# ============================================================================

class _RelationshipCollection(Collection):
    """
    Sequence of relationships maintained in rId order. Maintaining the
    relationships in sorted order makes the .rels files both repeatable and
    more readable, which is very helpful for debugging. 
    :class:`RelationshipCollection` has an attribute *_reltype_ordering* which
    is a sequence (tuple) of reltypes. If *_reltype_ordering* contains one or
    more reltype, the collection is maintained in reltype + partname.idx
    order and relationship ids (rIds) are renumbered to match that sequence
    and any numbering gaps are filled in.
    """
    def __init__(self):
        super(_RelationshipCollection, self).__init__()
        self.__reltype_ordering = ()
    
    def _additem(self, relationship):
        """
        Insert *relationship* into the appropriate position in this ordered
        collection.
        """
        rIds = [rel._rId for rel in self._values]
        if relationship._rId in rIds:
            tmpl = "cannot add relationship with duplicate rId '%s'"
            raise ValueError(tmpl % relationship._rId)
        self._values.append(relationship)
        self.__resequence()
        # register as observer of partname changes
        relationship._target.add_observer(self)
    
    @property
    def _next_rId(self):
        """
        Next available rId in collection, starting from 'rId1' and making use
        of any gaps in numbering, e.g. 'rId2' for rIds ['rId1', 'rId3'].
        """
        tmpl = 'rId%d'
        next_rId_num = 1
        for relationship in self._values:
            if relationship._num > next_rId_num:
                return tmpl % next_rId_num
            next_rId_num += 1
        return tmpl % next_rId_num
    
    @property
    def _reltype_ordering(self):
        """
        Tuple of relationship types, e.g. ``(RT_SLIDE, RT_SLIDELAYOUT)``. If
        present, relationships of those types are grouped, and those groups
        are ordered in the same sequence they appear in the tuple. In
        addition, relationships of the same type are sequenced in order of
        partname number (e.g. 16 for /ppt/slides/slide16.xml). If empty, the
        collection is maintained in existing rId order; rIds are not
        renumbered and any gaps in numbering are left to remain. Specifying
        this value for a collection with members causes it to be immediately
        reordered. The ordering is maintained as relationships are added or
        removed, renumbering rIds whenever necessary to also maintain the
        sequence in rId order.
        """
        return self.__reltype_ordering
    
    @_reltype_ordering.setter
    def _reltype_ordering(self, ordering):
        self.__reltype_ordering = tuple(ordering)
        self.__resequence()
    
    def rels_of_reltype(self, reltype):
        """
        Return a :class:`list` containing the subset of relationships in this
        collection of type *reltype*. The returned list is ordered by rId.
        Returns an empty list if there are no relationships of type *reltype*
        in the collection.
        """
        return [rel for rel in self._values if rel._reltype == reltype]
    
    def notify(self, subject, name, value):
        """RelationshipCollection implements the Observer interface"""
        if isinstance(subject, BasePart):
            if name == 'partname':
                self.__resequence()
    
    def __resequence(self):
        """
        Sort relationships and renumber if necessary to maintain values in rId
        order.
        """
        if self.__reltype_ordering:
            def reltype_key(rel):
                reltype = rel._reltype
                if reltype in self.__reltype_ordering:
                    return self.__reltype_ordering.index(reltype)
                return len(self.__reltype_ordering)
            def partname_idx_key(rel):
                partname = util.Partname(rel._target.partname)
                if partname.idx is None:
                    return 0
                return partname.idx
            self._values.sort(key=lambda rel: partname_idx_key(rel))
            self._values.sort(key=lambda rel: reltype_key(rel))
            # renumber consistent with new sort order
            for idx, relationship in enumerate(self._values):
                relationship._rId = 'rId%d' % (idx+1)
        else:
            self._values.sort(key=lambda rel: rel._num)
    

class _Relationship(object):
    """
    Relationship to a part from a package or part. *rId* must be unique in any
    |_RelationshipCollection| this relationship is added to; use
    :attr:`_RelationshipCollection._next_rId` to get a unique rId.
    """
    def __init__(self, rId, reltype, target):
        super(_Relationship, self).__init__()
        # can't get _num right if rId is non-standard form
        assert rId.startswith('rId'), "rId in non-standard form: '%s'" % rId
        self.__rId = rId
        self.__reltype = reltype
        self.__target = target
    
    @property
    def _rId(self):
        """
        Relationship id for this relationship. Must be of the form
        ``rId[1-9][0-9]*``.
        """
        return self.__rId
    
    @property
    def _reltype(self):
        """
        Relationship type URI for this relationship. Corresponds roughly to
        the part type of the target part.
        """
        return self.__reltype
    
    @property
    def _target(self):
        """
        Target part of this relationship. Relationships are directed, from a
        source and a target. The target is always a part.
        """
        return self.__target
    
    @property
    def _num(self):
        """
        The numeric portion of the rId of this |_Relationship|, expressed as
        an :class:`int`. For example, :attr:`_num` for a relationship with an
        rId of ``'rId12'`` would be ``12``.
        """
        return int(self.__rId[3:])
    
    @_rId.setter
    def _rId(self, value):
        self.__rId = value
    


# ============================================================================
# Parts
# ============================================================================

class PartCollection(Collection):
    """
    Sequence of parts. Sensitive to partname index when ordering parts added
    via _loadpart(), e.g. ``/ppt/slide/slide2.xml`` appears before
    ``/ppt/slide/slide10.xml`` rather than after it as it does in a
    lexicographical sort.
    """
    def __init__(self):
        super(PartCollection, self).__init__()
    
    def _loadpart(self, part):
        """
        Insert a new part loaded from a package, such that list remains
        sorted in logical partname order (e.g. slide10.xml comes after
        slide9.xml).
        """
        new_partidx = util.Partname(part.partname).idx
        for idx, seq_part in enumerate(self._values):
            partidx = util.Partname(seq_part.partname).idx
            if partidx > new_partidx:
                self._values.insert(idx, part)
                return
        self._values.append(part)
    

class ImageCollection(PartCollection):
    """
    Immutable sequence of images, typically belonging to an instance of
    :class:`Package`. An image part containing a particular image blob appears
    only once in an instance, regardless of how many times it is referenced by
    a pic shape in a slide.
    """
    def __init__(self):
        super(ImageCollection, self).__init__()
    
    def add_image(self, path):
        """
        Return image part containing the image at *path*. If the image does
        not yet exist, a new one is created.
        """
        # use Image constructor to validate and characterize image file
        image = Image(path)
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
    

class Part(object):
    """
    Part factory. Returns an instance of the appropriate custom part type for
    part types that have them, BasePart otherwise.
    """
    def __new__(cls, reltype, content_type):
        """
        *reltype* is the relationship type, e.g. ``RT_SLIDE``, corresponding
        to the type of part to be created. For at least one part type, in
        particular the presentation part type, *content_type* is also required
        in order to fully specify the part to be created.
        """
        # log.debug("Creating Part for %s", reltype)
        if reltype == RT_OFFICEDOCUMENT:
            if content_type in (CT_PRESENTATION, CT_TEMPLATE, CT_SLIDESHOW):
                return Presentation()
            else:
                tmpl = "Not a presentation content type, got '%s'"
                raise InvalidPackageError(tmpl % content_type)
        elif reltype == RT_SLIDE:
            return Slide()
        elif reltype == RT_SLIDELAYOUT:
            return SlideLayout()
        elif reltype == RT_SLIDEMASTER:
            return SlideMaster()
        elif reltype == RT_IMAGE:
            return Image()
        return BasePart()
    

class BasePart(Observable):
    """
    Base class for presentation model parts. Provides common code to all parts
    and is the class we instantiate for parts we don't unmarshal or manipulate
    yet.
    
    .. attribute:: _element
    
       ElementTree element for XML parts. ``None`` for binary parts.
    
    .. attribute:: _load_blob
    
       Contents of part as a byte string extracted from the package file. May
       be set to ``None`` by subclasses that override ._blob after content is
       unmarshaled, to free up memory.
    
    .. attribute:: _relationships
    
       :class:`RelationshipCollection` instance containing the relationships
       for this part.
    
    """
    _nsmap = namespaces('a', 'r', 'p')
    
    def __init__(self, content_type=None):
        """
        Needs content_type parameter so newly created parts (not loaded from
        package) can register their content type.
        """
        super(BasePart, self).__init__()
        self.__content_type = content_type
        self.__partname = None
        self._element = None
        self._load_blob = None
        self._relationships = _RelationshipCollection()
    
    @property
    def _blob(self):
        """
        Default is to return unchanged _load_blob. Dynamic parts will
        override. Raises :class:`ValueError` if _load_blob is None.
        """
        if self.partname.endswith('.xml'):
            assert self._element is not None, 'BasePart._blob is undefined '\
                                 'for xml parts when part.__element is None'
            xml = etree.tostring(self._element, encoding='UTF-8',
                                 pretty_print=True, standalone=True)
            return xml
        # default for binary parts is to return _load_blob unchanged
        assert self._load_blob, "BasePart._blob called on part with no "\
               "_load_blob; perhaps _blob not overridden by sub-class?"
        return self._load_blob
    
    @property
    def _content_type(self):
        """
        Return content type of this part, e.g.
        'application/vnd.openxmlformats-officedocument.theme+xml'. Throws on
        access before content type is set (by load or otherwise).
        """
        if self.__content_type is None:
            msg = "_content_type called on part with no content type"
            raise ValueError(msg)
        return self.__content_type
    
    @_content_type.setter
    def _content_type(self, content_type):
        self.__content_type = content_type
    
    @property
    def partname(self):
        """Part name of this part, e.g. '/ppt/slides/slide1.xml'."""
        assert self.__partname, "BasePart.partname referenced before assigned"
        return self.__partname
    
    @partname.setter
    def partname(self, partname):
        self.__partname = partname
        self._notify_observers('partname', self.__partname)
    
    def _add_relationship(self, reltype, target_part):
        """
        Return new relationship of *reltype* to *target_part* after adding it
        to the relationship collection of this part.
        """
        rId = self._relationships._next_rId
        rel = _Relationship(rId, reltype, target_part)
        self._relationships._additem(rel)
        return rel
    
    def _load(self, pkgpart, part_dict):
        """
        Load part and relationships from package part, and propagate load
        process down the relationship graph. *pkgpart* is an instance of
        :class:`pptx.packaging.Part` containing the part contents read from
        the on-disk package. *part_dict* is a dictionary of already-loaded
        parts, keyed by partname.
        """
        # log.debug("loading part %s", pkgpart.partname)
        
        # # set attributes from package part
        self.__content_type = pkgpart.content_type
        self.__partname = pkgpart.partname
        if pkgpart.partname.endswith('.xml'):
            self._element = etree.fromstring(pkgpart.blob)
        else:
            self._load_blob = pkgpart.blob
        
        # discard any previously loaded relationships
        self._relationships = _RelationshipCollection()
        
        # load relationships and propagate load for related parts
        for pkgrel in pkgpart.relationships:
            # unpack working values for part to be loaded
            reltype = pkgrel.reltype
            target_pkgpart = pkgrel.target
            partname = target_pkgpart.partname
            content_type = target_pkgpart.content_type
            
            # create target part
            if partname in part_dict:
                part = part_dict[partname]
            else:
                part = Part(reltype, content_type)
                part_dict[partname] = part
                part._load(target_pkgpart, part_dict)
            
            # create model-side package relationship
            model_rel = _Relationship(pkgrel.rId, reltype, part)
            self._relationships._additem(model_rel)
        return self
    

class Presentation(BasePart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    def __init__(self):
        super(Presentation, self).__init__()
        self.__slidemasters = PartCollection()
        self.__slides = SlideCollection(self)
    
    @property
    def slidemasters(self):
        """
        List of :class:`SlideMaster` objects belonging to this presentation.
        """
        return tuple(self.__slidemasters)
    
    @property
    def slides(self):
        """
        :class:`SlideCollection` object containing the slides in this
        presentation.
        """
        return self.__slides
    
    @property
    def _blob(self):
        """
        Rewrite sldId elements in sldIdLst before handing over to super for
        transformation of _element into a blob.
        """
        sldIdLst = (self.__sldIdLst if self.__sldIdLst is not None
                                    else self.__add_sldIdLst())
        sldIdLst.clear()
        sld_rels = self._relationships.rels_of_reltype(RT_SLIDE)
        for idx, rel in enumerate(sld_rels):
            sldId = etree.SubElement(sldIdLst, qname('p', 'sldId'))
            sldId.set('id', str(256+idx))
            sldId.set(qname('r', 'id'), rel._rId)
        return super(Presentation, self)._blob
    
    def _load(self, pkgpart, part_dict):
        """
        Load presentation from package part.
        """
        # call parent to do generic aspects of load
        super(Presentation, self)._load(pkgpart, part_dict)
        
        # set reltype ordering so rels file ordering is readable
        self._relationships._reltype_ordering = (RT_SLIDEMASTER,
            RT_NOTESMASTER, RT_HANDOUTMASTER, RT_SLIDE, RT_PRESPROPS,
            RT_VIEWPROPS, RT_TABLESTYLES, RT_THEME)
        
        # selectively unmarshal relationships for now
        for rel in self._relationships:
            # log.debug("Presentation Relationship %s", rel._reltype)
            if rel._reltype == RT_SLIDEMASTER:
                self.__slidemasters._loadpart(rel._target)
            elif rel._reltype == RT_SLIDE:
                self.__slides._loadpart(rel._target)
        return self
    
    def __add_sldIdLst(self):
        """
        Add a <p:sldIdLst> element to <p:presentation> in the right sequence
        among its siblings.
        """
        assert self.__sldIdLst is None, "__add_sldIdLst() called where "\
                                        "<p:sldIdLst> already exists"
        sldIdLst = etree.Element(qname('p', 'sldIdLst'))
        # insert new sldIdLst element in right sequence
        if self.__sldSz is not None:
            self.__sldSz.addprevious(sldIdLst)
        else:
            self.__notesSz.addprevious(sldIdLst)
        return sldIdLst
    
    @property
    def __notesSz(self):
        """Bookmark to ``<p:notesSz>`` child element"""
        return _child(self._element, 'p:notesSz', self._nsmap)
    
    @property
    def __sldIdLst(self):
        """Bookmark to ``<p:sldIdLst>`` child element"""
        return _child(self._element, 'p:sldIdLst', self._nsmap)
    
    @property
    def __sldSz(self):
        """Bookmark to ``<p:sldSz>`` child element"""
        return _child(self._element, 'p:sldSz', self._nsmap)
    

class Image(BasePart):
    """
    Image part. Corresponds to package files ppt/media/image[1-9][0-9]*.*.
    If *path* parameter is used, the image file at that location is loaded.
    """
    def __init__(self, path=None):
        super(Image, self).__init__()
        self.__ext = None
        if path:
            self.__load_image_from_file(path)
    
    @property
    def ext(self):
        """Return file extension for this image"""
        assert self.__ext, "Image.__ext referenced before assigned"
        return self.__ext
    
    @property
    def _sha1(self):
        """Return SHA1 hash digest for image"""
        return hashlib.sha1(self._blob).hexdigest()
    
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
        super(Image, self)._load(pkgpart, part_dict)
        # set file extension
        self.__ext = os.path.splitext(pkgpart.partname)[1]
        # return self-reference to allow generative calling
        return self
    
    @staticmethod
    def __image_file_content_type(path):
        """Return the content type of graphic image file at *path*"""
        ext = os.path.splitext(path)[1]
        if ext not in spec.default_content_types:
            tmpl = "unsupported image file extension '%s' at '%s'"
            raise TypeError(tmpl % (ext, path))
        content_type = spec.default_content_types[ext]
        if not content_type.startswith('image/'):
            tmpl = "'%s' is not an image content type; path '%s'"
            raise TypeError(tmpl % (content_type, path))
        return content_type
    
    def __load_image_from_file(self, path):
        """
        Load image file at *path*.
        """
        # set extension
        self.__ext = os.path.splitext(path)[1]
        # set content type
        self._content_type = self.__image_file_content_type(path)
        # lodge blob
        with open(path, 'rb') as f:
            self._load_blob = f.read()
    


# ============================================================================
# Slide Parts
# ============================================================================

class SlideCollection(PartCollection):
    """
    Immutable sequence of slides, typically belonging to an instance of
    :class:`Presentation`, with model-domain methods for manipulating the
    slides in the presentation.
    """
    def __init__(self, presentation):
        super(SlideCollection, self).__init__()
        self.__presentation = presentation
    
    def add_slide(self, slidelayout):
        """Add a new slide that inherits layout from *slidelayout*."""
        # 1. construct new slide
        slide = Slide(slidelayout)
        # 2. add it to this collection
        self._values.append(slide)
        # 3. assign its partname
        self.__rename_slides()
        # 4. add presentation->slide relationship
        self.__presentation._add_relationship(RT_SLIDE, slide)
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
    

class BaseSlide(BasePart):
    """
    Base class for slide parts, e.g. slide, slideLayout, slideMaster,
    notesSlide, notesMaster, and handoutMaster.
    """
    def __init__(self, content_type=None):
        """
        Needs content_type parameter so newly created parts (not loaded from
        package) can register their content type.
        """
        super(BaseSlide, self).__init__(content_type)
        self._shapes = None
    
    @property
    def name(self):
        """Internal name of this slide-like object."""
        root = self._element
        cSld = root.xpath('./p:cSld', namespaces=self._nsmap)[0]
        name = cSld.get('name', default='')
        return name
    
    @property
    def shapes(self):
        """Collection of shape objects belonging to this slide."""
        assert self._shapes is not None, ("BaseSlide.shapes referenced"
                                                      " before assigned")
        return self._shapes
    
    def _load(self, pkgpart, part_dict):
        """Handle aspects of loading that are general to slide types."""
        # call parent to do generic aspects of load
        super(BaseSlide, self)._load(pkgpart, part_dict)
        # unmarshal shapes
        xpath = './p:cSld/p:spTree'
        spTree = self._element.xpath(xpath, namespaces=self._nsmap)[0]
        self._shapes = ShapeCollection(spTree, self)
        # return self-reference to allow generative calling
        return self
    

class Slide(BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    def __init__(self, slidelayout=None):
        super(Slide, self).__init__(CT_SLIDE)
        self.__slidelayout = slidelayout
        self._element = self.__minimal_element
        spTree = self._element.xpath('//p:spTree', namespaces=self._nsmap)[0]
        self._shapes = ShapeCollection(spTree, self)
        # if slidelayout, this is a slide being added, not one being loaded
        if slidelayout:
            self._shapes._clone_layout_placeholders(slidelayout)
            # add relationship to slideLayout part
            self._add_relationship(RT_SLIDELAYOUT, slidelayout)
    
    @property
    def slidelayout(self):
        """
        :class:`SlideLayout` object this slide inherits appearance from.
        """
        return self.__slidelayout
    
    def _load(self, pkgpart, part_dict):
        """
        Load slide from package part.
        """
        # call parent to do generic aspects of load
        super(Slide, self)._load(pkgpart, part_dict)
        # selectively unmarshal relationships for now
        for rel in self._relationships:
            # log.debug("SlideMaster Relationship %s", rel._reltype)
            if rel._reltype == RT_SLIDELAYOUT:
                self.__slidelayout = rel._target
        return self
    
    @property
    def __minimal_element(self):
        """
        Return element containing the minimal XML for a slide, based on what
        is required by the XMLSchema.
        """
        sld = etree.Element(qname('p', 'sld'), nsmap=self._nsmap)
        cSld = etree.SubElement(sld, qname('p', 'cSld'))
        spTree = etree.SubElement(cSld, qname('p', 'spTree'))
        nvGrpSpPr = etree.SubElement(spTree, qname('p', 'nvGrpSpPr'))
        cNvPr = etree.SubElement(nvGrpSpPr, qname('p', 'cNvPr'))
        cNvPr.set('id', '1')
        cNvPr.set('name', '')
        cNvGrpSpPr = etree.SubElement(nvGrpSpPr, qname('p', 'cNvGrpSpPr'))
        nvPr = etree.SubElement(nvGrpSpPr, qname('p', 'nvPr'))
        grpSpPr = etree.SubElement(spTree, qname('p', 'grpSpPr'))
        return sld
    

class SlideLayout(BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ppt/slideLayouts/slideLayout[1-9][0-9]*.xml.
    """
    def __init__(self):
        super(SlideLayout, self).__init__(CT_SLIDELAYOUT)
        self.__slidemaster = None
    
    @property
    def slidemaster(self):
        """Slide master from which this slide layout inherits properties."""
        assert self.__slidemaster is not None, ("SlideLayout.slidemaster "
                                                "referenced before assigned")
        return self.__slidemaster
    
    def _load(self, pkgpart, part_dict):
        """
        Load slide layout from package part.
        """
        # call parent to do generic aspects of load
        super(SlideLayout, self)._load(pkgpart, part_dict)
        
        # selectively unmarshal relationships we need
        for rel in self._relationships:
            # log.debug("SlideLayout Relationship %s", rel._reltype)
            # get slideMaster from which this slideLayout inherits properties
            if rel._reltype == RT_SLIDEMASTER:
                self.__slidemaster = rel._target
        
        # return self-reference to allow generative calling
        return self
    

class SlideMaster(BaseSlide):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    
    TECHNOTE: In the Microsoft API, Master is a general type that all of
    SlideMaster, SlideLayout (CustomLayout), HandoutMaster, and NotesMaster
    inherit from. So might look into why that is and consider refactoring the
    various masters a bit later.
    """
    def __init__(self):
        super(SlideMaster, self).__init__(CT_SLIDEMASTER)
        self.__slidelayouts = PartCollection()
    
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
        super(SlideMaster, self)._load(pkgpart, part_dict)
        
        # selectively unmarshal relationships for now
        for rel in self._relationships:
            # log.debug("SlideMaster Relationship %s", rel._reltype)
            if rel._reltype == RT_SLIDELAYOUT:
                self.__slidelayouts._loadpart(rel._target)
        return self
    


# ============================================================================
# Shapes
# ============================================================================

class BaseShape(object):
    """
    Base class for shape objects.
    """
    _nsmap = namespaces('a', 'r', 'p')
    
    def __init__(self, shape_element):
        # log.debug('BaseShape.__init__() called w/element 0x%X',
        #            id(shape_element))
        super(BaseShape, self).__init__()
        self._element = shape_element
        self.__cNvPr = shape_element.xpath('./*[1]/p:cNvPr',
                                           namespaces=self._nsmap)[0]
    
    @property
    def has_textframe(self):
        """
        True if this shape has a txBody element and can support text.
        """
        xpath = './p:txBody'
        elements = self._element.xpath(xpath, namespaces=self._nsmap)
        return len(elements) > 0
    
    @property
    def id(self):
        """
        Id of this shape. Note that ids are constrained to positive integers.
        """
        return int(self.__cNvPr.get('id'))
    
    @property
    def is_placeholder(self):
        """
        True if this shape is a placeholder. A shape is a placeholder if it
        has a <p:ph> element.
        """
        xpath = './*[1]/p:nvPr/p:ph'
        ph_elms = self._element.xpath(xpath, namespaces=self._nsmap)
        return len(ph_elms) > 0
    
    @property
    def name(self):
        """Name of this shape."""
        return self.__cNvPr.get('name', default='')
    
    def set_text(self, value):
        """Replace all text with single run containing *value*"""
        if not self.has_textframe:
            raise TypeError("cannot set text of shape with no text frame")
        self.textframe.text = value
    
    text = property(None, set_text)
    
    @property
    def textframe(self):
        """
        TextFrame instance for this shape. Raises :class:`ValueError` if shape
        has no text frame. Use :meth:`has_textframe` to check whether a shape
        has a text frame.
        """
        xpath = './p:txBody'
        elements = self._element.xpath(xpath, namespaces=self._nsmap)
        if len(elements) == 0:
            raise ValueError('shape has no text frame')
        txBody = elements[0]
        return TextFrame(txBody)
    
    @property
    def _is_title(self):
        """
        True if this shape is a title placeholder.
        """
        xpath = './*[1]/p:nvPr/p:ph'
        ph_elms = self._element.xpath(xpath, namespaces=self._nsmap)
        if len(ph_elms)==0:
            return False
        idx = ph_elms[0].get('idx', '0')
        return idx=='0'
    

class ShapeCollection(BaseShape, Collection):
    """
    Sequence of shapes. Corresponds to CT_GroupShape in pml schema. Note that
    while spTree in a slide is a group shape, the group shape is recursive in
    that a group shape can include other group shapes within it.
    """
    NVGRPSPPR    = qname('p', 'nvGrpSpPr')
    GRPSPPR      = qname('p', 'grpSpPr')
    SP           = qname('p', 'sp')
    GRPSP        = qname('p', 'grpSp')
    GRAPHICFRAME = qname('p', 'graphicFrame')
    CXNSP        = qname('p', 'cxnSp')
    PIC          = qname('p', 'pic')
    CONTENTPART  = qname('p', 'contentPart')
    EXTLST       = qname('p', 'extLst')
    
    def __init__(self, spTree, slide=None):
        # log.debug('ShapeCollect.__init__() called w/element 0x%X', id(spTree))
        super(ShapeCollection, self).__init__(spTree)
        self.__spTree = spTree
        self.__slide = slide
        # unmarshal shapes
        for elm in spTree:
            # log.debug('elm.tag == %s', elm.tag[60:])
            if elm.tag in (self.NVGRPSPPR, self.GRPSPPR, self.EXTLST):
                continue
            elif elm.tag == self.SP:
                shape = Shape(elm)
            elif elm.tag == self.PIC:
                shape = Picture(elm)
            elif elm.tag == self.GRPSP:
                shape = ShapeCollection(elm)
            elif elm.tag == self.CONTENTPART:
                msg = "first time 'contentPart' shape encountered in the "\
                      "wild, please let developer know and send example"
                raise ValueError(msg)
            else:
                shape = BaseShape(elm)
            self._values.append(shape)
    
    @property
    def placeholders(self):
        """
        Immutable sequence containing the placeholder shapes in this shape
        collection, sorted in *idx* order.
        """
        placeholders =\
            [Placeholder(sp) for sp in self._values if sp.is_placeholder]
        placeholders.sort(key=lambda ph: ph.idx)
        return tuple(placeholders)
    
    @property
    def title(self):
        """The title shape in collection or None if no title placeholder."""
        for shape in self._values:
            if shape._is_title:
                return shape
        return None
    
    def add_picture(self, path, top, left):
        """
        Add picture shape displaying image in file located at *path*, placing
        the shape such that its top left corner is at *top* inches from the
        top of the slide and *left* inches from the left of the slide.
        """
        pkg = Package.containing(self.__slide)
        image = pkg._images.add_image(path)
        rel = self.__slide._add_relationship(RT_IMAGE, image)
        pic = self.__pic(rel._rId, path, top, left)
        self.__spTree.append(pic)
        picture = Picture(pic)
        self._values.append(picture)
        return picture
    
    def _clone_layout_placeholders(self, slidelayout):
        """
        Add placeholder shapes based on those in *slidelayout*. Z-order of
        placeholders is preserved. Latent placeholders (date, slide number,
        and footer) are not cloned.
        """
        latent_ph_types = (PH_TYPE_DT, PH_TYPE_SLDNUM, PH_TYPE_FTR)
        for sp in slidelayout.shapes:
            # log.debug("considering clone of sp.id == %d  %r", sp.id, sp)
            if not sp.is_placeholder:
                continue
            ph = Placeholder(sp)
            if ph.type in latent_ph_types:
                continue
            self.__clone_layout_placeholder(ph)
    
    def __clone_layout_placeholder(self, layout_ph):
        """
        Add a new placeholder shape based on the slide layout placeholder
        *layout_ph*.
        """
        id = self.__next_shape_id
        ph_type = layout_ph.type
        orient = layout_ph.orient
        name = self.__next_ph_name(ph_type, id, orient)
        # <p:sp>
        sp = etree.SubElement(self.__spTree, qname('p', 'sp'))
        #   <p:nvSpPr>
        nvSpPr = etree.SubElement(sp, qname('p', 'nvSpPr'))
        #     <p:cNvPr id="9" name="Thing Placeholder 8"/>
        cNvPr = etree.SubElement(nvSpPr, qname('p', 'cNvPr'))
        cNvPr.set('id', str(id))
        cNvPr.set('name', name)
        #     <p:cNvSpPr>
        cNvSpPr = etree.SubElement(nvSpPr, qname('p', 'cNvSpPr'))
        #       <a:spLocks noGrp="1"/>
        spLocks = etree.SubElement(cNvSpPr, qname('a', 'spLocks'))
        spLocks.set('noGrp', '1')
        #     </p:cNvSpPr>
        #     <p:nvPr>
        nvPr = etree.SubElement(nvSpPr, qname('p', 'nvPr'))
        #       <p:ph type="body" orient="vert" sz="half" idx="9"/>
        ph = etree.SubElement(nvPr, qname('p', 'ph'))
        if ph_type != PH_TYPE_OBJ:
            ph.set('type', ph_type)
        if layout_ph.orient != PH_ORIENT_HORZ:
            ph.set('orient', layout_ph.orient)
        if layout_ph.sz != PH_SZ_FULL:
            ph.set('sz', layout_ph.sz)
        if layout_ph.idx != 0:
            ph.set('idx', str(layout_ph.idx))
        #     </p:nvPr>
        #   </p:nvSpPr>
        #   <p:spPr/>
        spPr = etree.SubElement(sp, qname('p', 'spPr'))
        #   <p:txBody>
        if ph_type in (PH_TYPE_TITLE, PH_TYPE_CTRTITLE, PH_TYPE_SUBTITLE,
                       PH_TYPE_BODY, PH_TYPE_OBJ):
            txBody = etree.SubElement(sp, qname('p', 'txBody'))
            bodyPr = etree.SubElement(txBody, qname('a', 'bodyPr'))
            lstStyle = etree.SubElement(txBody, qname('a', 'lstStyle'))
            p = etree.SubElement(txBody, qname('a', 'p'))
        #   </p:txBody>
        # </p:sp>
        shape = Shape(sp)
        self._values.append(shape)
        # log.debug("\n%s", etree.tostring(sp, pretty_print=True))
        return shape
    
    def __next_ph_name(self, ph_type, id, orient):
        """
        Next unique placeholder name for placeholder shape of type *ph_type*,
        with id number *id* and orientation *orient*. Usually will be standard
        placeholder root name suffixed with id-1, e.g.
        __next_ph_name(PH_TYPE_TBL, 4, 'horz') ==> 'Table Placeholder 3'. The
        number is incremented as necessary to make the name unique within the
        collection. If *orient* is ``'vert'``, the placeholder name is
        prefixed with ``'Vertical '``.
        """
        basename = slide_ph_basenames[ph_type]
        # prefix rootname with 'Vertical ' if orient is 'vert'
        if orient == PH_ORIENT_VERT:
            basename = 'Vertical %s' % basename
        # increment numpart as necessary to make name unique
        numpart = id - 1
        names = self.__spTree.xpath('//p:cNvPr/@name', namespaces=self._nsmap)
        while True:
            name = '%s %d' % (basename, numpart)
            if name not in names:
                break
            numpart += 1
        # log.debug("assigned placeholder name '%s'" % name)
        return name
    
    @property
    def __next_shape_id(self):
        """
        Next available drawing object id number in collection, starting from 1
        and making use of any gaps in numbering. In practice, the minimum id
        is 2 because the spTree element is always assigned id="1".
        """
        cNvPrs = self.__spTree.xpath('//p:cNvPr', namespaces=self._nsmap)
        ids = [int(cNvPr.get('id')) for cNvPr in cNvPrs]
        ids.sort()
        # first gap in sequence wins, or falls off the end as max(ids)+1
        next_id = 1
        for id in ids:
            if id > next_id:
                break
            next_id += 1
        return next_id
    
    def __pic(self, rId, path, top, left):
        """Return minimal ``<p:pic>`` element based on *rId* and *path*."""
        id = self.__next_shape_id
        shapename = 'Picture %d' % (id-1)
        filename = os.path.split(path)[1]
        
        cx_px, cy_px = PIL_Image.open(path).size
        cx = Px(cx_px)
        cy = Px(cy_px)
        
        pic = etree.Element(qname('p', 'pic'), nsmap=self._nsmap)
        
        nvPicPr  = etree.SubElement(pic,      qname('p', 'nvPicPr'))
        cNvPr    = etree.SubElement(nvPicPr,  qname('p', 'cNvPr'))
        cNvPr.set('id',    str(id))
        cNvPr.set('name',  shapename)
        cNvPr.set('descr', filename)
        cNvPicPr = etree.SubElement(nvPicPr,  qname('p', 'cNvPicPr'))
        nvPr     = etree.SubElement(nvPicPr,  qname('p', 'nvPr'))
        
        blipFill = etree.SubElement(pic,      qname('p', 'blipFill'))
        blip     = etree.SubElement(blipFill, qname('a', 'blip'))
        blip.set(qname('r', 'embed'), rId)
        stretch  = etree.SubElement(blipFill, qname('a', 'stretch'))
        fillRect = etree.SubElement(stretch,  qname('a', 'fillRect'))
        
        spPr = etree.SubElement(pic,  qname('p', 'spPr'))
        xfrm = etree.SubElement(spPr, qname('a', 'xfrm'))
        off  = etree.SubElement(xfrm, qname('a', 'off'))
        off.set('x', str(left))
        off.set('y', str(top))
        ext  = etree.SubElement(xfrm, qname('a', 'ext'))
        ext.set('cx', str(cx))
        ext.set('cy', str(cy))
        
        prstGeom = etree.SubElement(spPr, qname('a', 'prstGeom'))
        prstGeom.set('prst', 'rect')
        avLst    = etree.SubElement(prstGeom, qname('a', 'avLst'))
        
        return pic
    

class Placeholder(object):
    """
    Decorator (pattern) class for adding placeholder properties to a shape
    that contains a placeholder element, e.g. ``<p:ph>``.
    """
    def __new__(cls, shape):
        cls = type('PlaceholderDecorator',(Placeholder, shape.__class__), {})
        return object.__new__(cls)
    
    def __init__(self, shape):
        self.__decorated = shape
        xpath = './*[1]/p:nvPr/p:ph'
        self.__ph = self._element.xpath(xpath, namespaces=self._nsmap)[0]
    
    def __getattr__(self, name):
        """
        Called when *name* is not found in ``self`` or in class tree. In this
        case, delegate attribute lookup to decorated (it's probably in its
        instance namespace).
        """
        return getattr(self.__decorated, name)
    
    @property
    def type(self):
        """Placeholder type, e.g. PH_TYPE_CTRTITLE"""
        type = self.__ph.get('type')
        return type if type else PH_TYPE_OBJ
    
    @property
    def orient(self):
        """Placeholder 'orient' attribute, e.g. PH_ORIENT_HORZ"""
        orient = self.__ph.get('orient')
        return orient if orient else PH_ORIENT_HORZ
    
    @property
    def sz(self):
        """Placeholder 'sz' attribute, e.g. PH_SZ_FULL"""
        sz = self.__ph.get('sz')
        return sz if sz else PH_SZ_FULL
    
    @property
    def idx(self):
        """Placeholder 'idx' attribute, e.g. '0'"""
        idx = self.__ph.get('idx')
        return int(idx) if idx else 0
    

class Picture(BaseShape):
    """
    A picture shape, one that places an image on a slide. Corresponds to the
    ``<p:pic>`` element.
    """
    def __init__(self, pic):
        super(Picture, self).__init__(pic)
    

class Shape(BaseShape):
    """
    A shape that can appear on a slide. Corresponds to the ``<p:sp>`` element
    that can appear in any of the slide-type parts (slide, slideLayout, 
    slideMaster, notesPage, notesMaster, handoutMaster).
    """
    def __init__(self, shape_element):
        super(Shape, self).__init__(shape_element)
    
    # @property
    # def __minimalelement(self):
    #     """
    #     Return an ElementTree element that contains all the elements and
    #     attributes of a shape required by the schema, initialized with default
    #     values where necessary or appropriate.
    #     """
    #     sp      = Element(qname('p', 'sp'))
    #     nvSpPr  = etree.SubElement(sp     , qname('p', 'nvSpPr'  ) )
    #     cNvPr   = etree.SubElement(nvSpPr , qname('p', 'cNvPr'   ) , id='None', name=self.name)
    #     cNvSpPr = etree.SubElement(nvSpPr , qname('p', 'cNvSpPr' ) )
    #     spPr    = etree.SubElement(sp     , qname('p', 'spPr'    ) )
    #     return sp
    # 


# ============================================================================
# Text-related classes
# ============================================================================

class TextFrame(object):
    """
    The part of a shape that contains its text. Not all shapes have a text
    frame. Corresponds to the ``<p:txBody>`` element that can appear as a
    chile element of ``<p:sp>``.
    """
    __nsmap = namespaces('a', 'r', 'p')
    
    def __init__(self, txBody):
        super(TextFrame, self).__init__()
        self.__txBody = txBody
    
    @property
    def paragraphs(self):
        """
        Immutable sequence of :class:`Paragraph` instances corresponding to
        the paragraphs in this text frame.
        """
        xpath = './a:p'
        p_elms = self.__txBody.xpath(xpath, namespaces=self.__nsmap)
        paragraphs = []
        for p in p_elms:
            paragraphs.append(Paragraph(p))
        return tuple(paragraphs)
    
    def clear(self):
        """
        Remove all paragraphs except one empty one.
        """
        p_list = self.__txBody.xpath('./a:p', namespaces=self.__nsmap)
        for p in p_list[1:]:
            self.__txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()
    
    def set_text(self, value):
        """Replace all text with single run containing *value*"""
        self.clear()
        self.paragraphs[0].text = value
    
    text = property(None, set_text)
    

class Paragraph(object):
    """
    Paragraph object.
    """
    __nsmap = namespaces('a', 'r', 'p')
    
    def __init__(self, p):
        super(Paragraph, self).__init__()
        self.__p = p
    
    def add_run(self):
        """Return a new run appended to the runs in this paragraph."""
        r = self.__new_run_element()
        # work out where to insert it, ahead of a:endParaRPr if there is one
        if self.__endParaRPr is not None:
            self.__endParaRPr.addprevious(r)
        else:
            self.__p.append(r)
        return Run(r)
    
    def clear(self):
        """Remove all runs from this paragraph."""
        # for now just remove all children; later might want to keep pPr
        self.__p.clear()
    
    def set_text(self, value):
        """Replace runs with single run containing *value*"""
        self.clear()
        r = self.add_run()
        r.text = value
    
    text = property(None, set_text)
    
    @property
    def runs(self):
        """
        Immutable sequence of :class:`Run` instances corresponding to the runs
        in this paragraph.
        """
        xpath = './a:r'
        r_elms = self.__p.xpath(xpath, namespaces=self.__nsmap)
        runs = []
        for r in r_elms:
            runs.append(Run(r))
        return tuple(runs)
    
    @property
    def __endParaRPr(self):
        """Bookmark to ``<a:endParaRPr>`` child element"""
        return _child(self.__p, 'a:endParaRPr', self.__nsmap)
    
    def __new_run_element(self):
        """Construct and return an empty run element"""
        # <a:r>
        r = etree.Element(qname('a', 'r'))
        #   <a:t>
        etree.SubElement(r, qname('a', 't'))
        # </a:r>
        return r
    

class Run(object):
    """
    Text run object. Corresponds to ``<a:r>`` child element in a paragraph.
    """
    __nsmap = namespaces('a', 'r', 'p')
    
    def __init__(self, r):
        super(Run, self).__init__()
        self.__r = r
        self.__t = r.xpath('./a:t', namespaces=self.__nsmap)[0]
    
    @property
    def text(self):
        """
        Text contained in this run. A regular text run is required to contain
        exactly one ``<a:t>`` (text) element.
        """
        return self.__t.text
    
    @text.setter
    def text(self, str):
        """Set the text of this run to *str*."""
        self.__t.text = str
    


