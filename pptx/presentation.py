# -*- coding: utf-8 -*-
#
# presentation.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

'''API classes for dealing with presentations and other objects one typically
encounters as an end-user of the PowerPoint user interface.'''

from lxml import etree

import pptx.packaging
import pptx.util as util

from pptx.exceptions import InvalidPackageError

from pptx.spec import namespaces, qname
from pptx.spec import (CT_PRESENTATION, CT_SLIDE, CT_SLIDELAYOUT,
    CT_SLIDEMASTER, CT_SLIDESHOW, CT_TEMPLATE)
from pptx.spec import (RT_HANDOUTMASTER, RT_NOTESMASTER, RT_OFFICEDOCUMENT,
    RT_PRESPROPS, RT_SLIDE, RT_SLIDELAYOUT, RT_SLIDEMASTER, RT_TABLESTYLES,
    RT_THEME, RT_VIEWPROPS)


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


class Package(object):
    """
    Root class of presentation object hierarchy.
    """
    def __init__(self):
        super(Package, self).__init__()
        self.__presentation = None
        self.__relationships = _RelationshipCollection()
    
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
    
    @property
    def partname(self):
        """Part name of this part, e.g. '/ppt/slides/slide1.xml'."""
        assert self.__partname, "BasePart.partname referenced before assigned"
        return self.__partname
    
    @partname.setter
    def partname(self, partname):
        self.__partname = partname
        self._notify_observers('partname', self.__partname)
    
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
        self.__slides = SlideCollection()
    
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
        nsmap = namespaces('a', 'r', 'p')
        sldIdLst = self._element.xpath('./p:sldIdLst', namespaces=nsmap)[0]
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
    


# ============================================================================
# Slide Parts
# ============================================================================

class SlideCollection(PartCollection):
    """
    Immutable sequence of slides, typically belonging to an instance of
    :class:`Presentation`, with model-domain methods for manipulating the
    slides in the presentation.
    """
    def __init__(self):
        super(SlideCollection, self).__init__()
    
    def add_slide(self, slidelayout):
        """Add a new slide that inherits layout from *slidelayout*."""
        slide = Slide(slidelayout)
        self._values.append(slide)
        return slide
    

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
        self.__shapes = None
    
    @property
    def name(self):
        """Internal name of this slide-like object."""
        root = self._element
        nsmap = namespaces('a', 'p', 'r')
        cSld = root.xpath('./p:cSld', namespaces=nsmap)[0]
        name = cSld.get('name', default='')
        return name
    
    @property
    def shapes(self):
        """Collection of shape objects belonging to this slide."""
        assert self.__shapes is not None, ("BaseSlide.shapes referenced"
                                                      " before assigned")
        return self.__shapes
    
    def _load(self, pkgpart, part_dict):
        """Handle aspects of loading that are general to slide types."""
        # call parent to do generic aspects of load
        super(BaseSlide, self)._load(pkgpart, part_dict)
        
        # unmarshal shapes
        nsmap = namespaces('a', 'p', 'r')
        spTree = self._element.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        self.__shapes = ShapeCollection(spTree)
        
        # return self-reference to allow generative calling
        return self
    

class Slide(BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    def __init__(self, slidelayout=None):
        super(Slide, self).__init__(CT_SLIDE)
        self.__slidelayout = slidelayout
    
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
    
    # @property
    # def minimalelement(self):
    #     """
    #     Return element containing the minimal XML for a slide, based on what
    #     is required by the XMLSchema. Note that in general, schema-minimal
    #     elements in the XML are not guaranteed to be loadable by PowerPoint,
    #     so test accordingly.
    #     """
    #     sld        = Element(qname('p', 'sld'), nsmap=self.nsmap)
    #     cSld       = SubElement(sld       , qname('p', 'cSld'       ) )
    #     spTree     = SubElement(cSld      , qname('p', 'spTree'     ) )
    #     nvGrpSpPr  = SubElement(spTree    , qname('p', 'nvGrpSpPr'  ) )
    #     cNvPr      = SubElement(nvGrpSpPr , qname('p', 'cNvPr'      ) ,
    #                               id=str(self._nextid), name=self.name)
    #     cNvGrpSpPr = SubElement(nvGrpSpPr , qname('p', 'cNvGrpSpPr' ) )
    #     nvPr       = SubElement(nvGrpSpPr , qname('p', 'nvPr'       ) )
    #     grpSpPr    = SubElement(spTree    , qname('p', 'grpSpPr'    ) )
    #     return sld
    # 

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
    __nsmap = namespaces('a', 'p', 'r')
    
    def __init__(self, shape_element):
        # log.debug('BaseShape.__init__() called w/element 0x%X',
        #            id(shape_element))
        super(BaseShape, self).__init__()
        self.__element = shape_element
        self.__cNvPr = shape_element.xpath('./*[1]/p:cNvPr',
                                           namespaces=self.__nsmap)[0]
    
    @property
    def has_textframe(self):
        """
        True if this shape has a txBody element and can support text.
        """
        xpath = './p:txBody'
        elements = self.__element.xpath(xpath, namespaces=self.__nsmap)
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
        ph_elms = self.__element.xpath(xpath, namespaces=self.__nsmap)
        return len(ph_elms) > 0
    
    @property
    def name(self):
        """Name of this shape."""
        return self.__cNvPr.get('name', default='')
    
    @property
    def textframe(self):
        """
        TextFrame instance for this shape. Raises :class:`ValueError` if shape
        has no text frame. Use :meth:`has_textframe` to check whether a shape
        has a text frame.
        """
        xpath = './p:txBody'
        elements = self.__element.xpath(xpath, namespaces=self.__nsmap)
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
        ph_elms = self.__element.xpath(xpath, namespaces=self.__nsmap)
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
    
    def __init__(self, spTree):
        # log.debug('ShapeCollect.__init__() called w/element 0x%X', id(spTree))
        super(ShapeCollection, self).__init__(spTree)
        # unmarshal shapes
        nsmap = namespaces('a', 'p', 'r')
        for elm in spTree:
            # log.debug('elm.tag == %s', elm.tag[60:])
            if elm.tag in (self.NVGRPSPPR, self.GRPSPPR, self.EXTLST):
                continue
            elif elm.tag == self.SP:
                shape = Shape(elm)
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
    def title(self):
        """The title shape in collection or None if no title placeholder."""
        for shape in self._values:
            if shape._is_title:
                return shape
        return None
    

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
    #     nvSpPr  = SubElement(sp     , qname('p', 'nvSpPr'  ) )
    #     cNvPr   = SubElement(nvSpPr , qname('p', 'cNvPr'   ) , id='None', name=self.name)
    #     cNvSpPr = SubElement(nvSpPr , qname('p', 'cNvSpPr' ) )
    #     spPr    = SubElement(sp     , qname('p', 'spPr'    ) )
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
    __nsmap = namespaces('a', 'p', 'r')
    
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
    

class Paragraph(object):
    """
    Paragraph object.
    """
    __nsmap = namespaces('a', 'p', 'r')
    
    def __init__(self, p):
        super(Paragraph, self).__init__()
        self.__p = p
    
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
    

class Run(object):
    """
    Text run object. Corresponds to ``<a:r>`` child element in a paragraph.
    """
    __nsmap = namespaces('a', 'p', 'r')
    
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
    


