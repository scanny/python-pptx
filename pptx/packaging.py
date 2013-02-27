# -*- coding: utf-8 -*-
#
# packaging.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

'''
The :mod:`pptx.packaging` module coheres around the concerns of reading and
writing presentations to and from a .pptx file. In doing so, it hides the
complexities of the package "directory" structure, reading and writing parts
to and from the package, zip file manipulation, and traversing relationship
items.

The main API class is :class:`pptx.packaging.Package` which provides the
methods :meth:`open`, :meth:`marshal`, and :meth:`save`.
'''

import os
import posixpath
import re

from StringIO import StringIO
from lxml import etree
from zipfile import ZipFile, is_zipfile, ZIP_DEFLATED

import pptx.spec

from pptx.exceptions import (CorruptedPackageError, DuplicateKeyError,
    NotXMLError, PackageNotFoundError)
from pptx.spec import qtag
from pptx.spec import PTS_HASRELS_NEVER, PTS_HASRELS_OPTIONAL

import logging
log = logging.getLogger('pptx.packaging')
log.setLevel(logging.DEBUG)
# log.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s'
                              ' - %(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)

PKG_BASE_URI = '/'


# ============================================================================
# API Classes
# ============================================================================

class Package(object):
    """
    Return a new package instance. Package is initially empty, call
    :meth:`open` to open an on-disk package or ``marshal()`` followed by
    ``save()`` to save an in-memory Office document.
    """
    
    PKG_RELSITEM_URI = '/_rels/.rels'
    
    def __init__(self):
        super(Package, self).__init__()
        self.__relationships = []
    
    @property
    def parts(self):
        """
        Return a list of :class:`pptx.packaging.Part` corresponding to the
        parts in this package.
        """
        return [part for part in self.__walkparts(self.relationships)]
    
    @property
    def relationships(self):
        """
        A tuple of :class:`pptx.packaging.Relationship` containing the package
        relationships for this package. Note these are not all the
        relationships in the package, just those from the package to top-level
        parts such as ``/ppt/presentation.xml`` and ``/docProps/core.xml``.
        These are useful primarily as the starting point to walk the part
        graph via its relationships.
        """
        return tuple(self.__relationships)
    
    def open(self, path):
        """Load the package located at *path*."""
        fs = FileSystem(path)
        cti = _ContentTypesItem().load(fs)
        self.__relationships = []  # discard any rels from prior load
        parts_dict = {}            # track loaded parts, graph is cyclic
        pkg_rel_elms = fs.getelement(Package.PKG_RELSITEM_URI)\
                         .findall(qtag('pr:Relationship'))
        for rel_elm in pkg_rel_elms:
            rId = rel_elm.get('Id')
            reltype = rel_elm.get('Type')
            partname = '/%s' % rel_elm.get('Target')
            part = Part()
            parts_dict[partname] = part
            part._load(fs, partname, cti, parts_dict)
            rel = Relationship(rId, self, reltype, part)
            self.__relationships.append(rel)
        fs.close()
        return self
    
    def marshal(self, model_pkg):
        """
        Load the contents of a model-side package such that it can be saved to
        a package file.
        """
        part_dict = {}  # keep track of marshaled parts, graph is cyclic
        for rel in model_pkg._relationships:
            # unpack working values for target part and relationship
            rId = rel._rId
            reltype = rel._reltype
            model_part = rel._target
            partname = model_part.partname
            # create package-part for target
            part = Part()
            part_dict[partname] = part
            part._marshal(model_part, part_dict)
            # create marshaled version of relationship
            marshaled_rel = Relationship(rId, self, reltype, part)
            self.__relationships.append(marshaled_rel)
        return self
    
    def save(self, path):
        """Save this package at *path*."""
        # open a zip filesystem for writing package
        path = self.__normalizedpath(path)
        zipfs = ZipFileSystem(path, 'w')
        # write [Content_Types].xml
        cti = _ContentTypesItem().compose(self.parts)
        zipfs.write_element(cti.element, '/[Content_Types].xml')
        # write pkg rels item
        zipfs.write_element(self.__relsitem_element, self.PKG_RELSITEM_URI)
        for part in self.parts:
            # write part item
            zipfs.write_blob(part.blob, part.partname)
            # write rels item if part has one
            if part.relationships:
                zipfs.write_element(part._relsitem_element, part._relsitemURI)
        zipfs.close()
    
    @property
    def __relsitem_element(self):
        nsmap = {None: pptx.spec.nsmap['pr']}
        element = etree.Element(qtag('pr:Relationships'), nsmap=nsmap)
        for rel in self.__relationships:
            element.append(rel._element)
        return element
    
    @staticmethod
    def __normalizedpath(path):
        """Add '.pptx' extension to path if not already there."""
        try:
            return path if path.endswith('.pptx') else '%s.pptx' % path
        except AttributeError, TypeEror:
            # path can be a file-like object
            return path

    
    @classmethod
    def __walkparts(cls, rels, parts=None):
        """
        Recursive generator method, walk relationships to iterate over all
        parts in this package. Leave out *parts* parameter in call to visit
        all parts.
        
        """
        # initial call can leave out parts parameter as a signal to initialize
        if parts is None:
            parts = []
        # log.debug("in __walkparts(), len(parts)==%d", len(parts))
        for rel in rels:
            # log.debug("rel.target.partname==%s", rel.target.partname)
            part = rel.target
            if part in parts:  # only visit each part once (graph is cyclic)
                continue
            parts.append(part)
            yield part
            for part in cls.__walkparts(part.relationships, parts):
                yield part
    

class Part(object):
    """
    Part instances are not intended to be constructed externally.
    :class:`pptx.packaging.Part` instances are constructed and initialized
    internally to the :meth:`Package.open` or :meth:`Package.marshal` methods.
    
    The following :class:`Part` instance attributes can be accessed once the
    part has been loaded as part of a package:
    
    .. attribute:: typespec
       
       An instance of :class:`PartTypeSpec` appropriate to the type of this
       part. The :class:`PartTypeSpec` instance provides attributes such as
       *content_type*, *baseURI*, etc. That are useful in several contexts.
    
    .. attribute:: blob
       
       The binary contents of this part contained in a byte string. For XML
       parts, this is simply the XML text. For binary parts such as an image,
       this is the string of bytes corresponding exactly to the bytes on disk
       for the binary object.
    
    """
    def __init__(self):
        super(Part, self).__init__()
        self.__partname = None
        self.__relationships = []
        self.typespec = None
        self.blob = None
    
    @property
    def content_type(self):
        """Content type of this part"""
        assert self.typespec, 'Part.content_type called before typespec set'
        return self.typespec.content_type
    
    @property
    def partname(self):
        """
        Package item URI for this part, commonly known as its part name,
        e.g. ``/ppt/slides/slide1.xml``
        """
        return self.__partname
    
    @property
    def relationships(self):
        """
        Tuple of :class:`Relationship` instances, each representing a
        relationship from this part to another part.
        """
        return tuple(self.__relationships)
    
    def _load(self, fs, partname, ct_dict, parts_dict):
        """
        Load part identified as *partname* from filesystem *fs* and propagate
        the load to related parts.
        """
        # log.debug("loading %s", partname)
        
        # calculate working values
        baseURI = os.path.split(partname)[0]
        content_type = ct_dict[partname]
        
        # set persisted attributes
        self.__partname = partname
        self.blob = fs.getblob(partname)
        self.typespec = PartTypeSpec(content_type)
        
        # load relationships and propagate load to target parts
        self.__relationships = []  # discard any rels from prior load
        rel_elms = self.__get_rel_elms(fs)
        for rel_elm in rel_elms:
            rId = rel_elm.get('Id')
            reltype = rel_elm.get('Type')
            target_relpath = rel_elm.get('Target')
            target_partname = posixpath.abspath(posixpath.join(baseURI,
                                                               target_relpath))
            if target_partname in parts_dict:
                target_part = parts_dict[target_partname]
            else:
                target_part = Part()
                parts_dict[target_partname] = target_part
                target_part._load(fs, target_partname, ct_dict, parts_dict)
            
            # create relationship to target_part
            rel = Relationship(rId, self, reltype, target_part)
            self.__relationships.append(rel)
        
        return self
    
    def _marshal(self, model_part, part_dict):
        """
        Load the contents of model-side part such that it can be saved to
        disk. Propagate marshalling to related parts.
        """
        # log.debug("marshalling %s", model_part.partname)
        
        # unpack working values
        content_type = model_part._content_type
        baseURI = os.path.split(model_part.partname)[0]
        # assign persisted attributes from model part
        self.__partname = model_part.partname
        self.blob = model_part._blob
        self.typespec = PartTypeSpec(content_type)
        
        # load relationships and propagate marshal to target parts
        for rel in model_part._relationships:
            # unpack working values for target part and relationship
            rId = rel._rId
            reltype = rel._reltype
            model_target_part = rel._target
            partname = model_target_part.partname
            # create package-part for target
            if partname in part_dict:
                part = part_dict[partname]
            else:
                part = Part()
                part_dict[partname] = part
                part._marshal(model_target_part, part_dict)
            # create marshalled version of relationship
            marshalled_rel = Relationship(rId, self, reltype, part)
            self.__relationships.append(marshalled_rel)
    
    @property
    def _relsitem_element(self):
        nsmap = {None: pptx.spec.nsmap['pr']}
        element = etree.Element(qtag('pr:Relationships'), nsmap=nsmap)
        for rel in self.__relationships:
            element.append(rel._element)
        return element
    
    @property
    def _relsitemURI(self):
        """
        Return theoretical package URI for this part's relationships item,
        without regard to whether this part actually has a relationships item.
        """
        head, tail = os.path.split(self.__partname)
        return '%s/_rels/%s.rels' % (head, tail)
    
    def __get_rel_elms(self, fs):
        """
        Helper method for _load(). Return list of this relationship elements
        for this part from *fs*. Returns empty list if there are no
        relationships for this part, either because parts of this type never
        have relationships or its relationships are optional and none exist in
        this filesystem (package).
        """
        relsitemURI = self.__relsitemURI(self.typespec, self.__partname, fs)
        if relsitemURI is None:
            return []
        if relsitemURI not in fs:
            tmpl = "required relationships item '%s' not found in package"
            raise CorruptedPackageError(tmpl % relsitemURI)
        baseURI = os.path.split(self.__partname)[0]
        root_elm = fs.getelement(relsitemURI)
        return root_elm.findall(qtag('pr:Relationship'))
    
    @staticmethod
    def __relsitemURI(typespec, partname, fs):
        """
        REFACTOR: Combine this logic into __get_rel_elms, it's the only caller
        and logic is partially redundant.
        
        Return package URI for this part's relationships item. Returns None if
        a part of this type never has relationships. Also returns None if a
        part of this type has only optional relationships and the package
        contains no rels item for this part.
        """
        if typespec.has_rels == PTS_HASRELS_NEVER:
            return None
        head, tail = os.path.split(partname)
        relsitemURI = '%s/_rels/%s.rels' % (head, tail)
        if typespec.has_rels == PTS_HASRELS_OPTIONAL:
            return relsitemURI if relsitemURI in fs else None
        return relsitemURI
    

class Relationship(object):
    """
    Return a new :class:`Relationship` instance with local identifier *rId*
    that associates *source* with *target*. *source* is an instance of either
    :class:`Package` or :class:`Part`. *target* is always an instance of
    :class:`Part`. Note that *rId* is only unique within the scope of
    *source*. Relationships do not have a globally unique identifier.
    
    The following attributes are available from :class:`Relationship`
    instances:
    
    .. attribute:: rId
       
       The source-local identifier for this relationship.
    
    .. attribute:: reltype
       
       The relationship type URI for this relationship. These are defined in
       the ECMA spec and look something like:
       'http://schemas.openxmlformats.org/.../relationships/slide'
    
    .. attribute:: target
       
       The target :class:`pptx.packaging.Part` instance in this relationship.
    
    """
    def __init__(self, rId, source, reltype, target):
        super(Relationship, self).__init__()
        self.__source = source
        self.rId = rId
        self.reltype = reltype
        self.target = target
    
    @property
    def _element(self):
        """
        The :class:`ElementTree._Element` instance containing the XML
        representation of this Relationship.
        """
        element = etree.Element('Relationship')
        element.set('Id', self.rId)
        element.set('Type', self.reltype)
        element.set('Target', self.__target_relpath)
        return element
    
    @property
    def __baseURI(self):
        """Return the directory part of the source itemURI."""
        if isinstance(self.__source, Part):
            return os.path.split(self.__source.partname)[0]
        return PKG_BASE_URI
    
    @property
    def __target_relpath(self):
        # workaround for posixpath bug in 2.6, doesn't generate correct
        # relative path when *start* (second) parameter is root ('/')
        if self.__baseURI == '/':
            relpath = self.target.partname[1:]
        else:
            relpath = posixpath.relpath(self.target.partname, self.__baseURI)
        return relpath
    

class PartTypeSpec(object):
    """
    Return an instance of :class:`PartTypeSpec` containing metadata for parts
    of type *content_type*. Instances are cached, so no more than one instance
    for a particular content type is in memory.
    
    Instances provide the following attributes:
    
    .. attribute:: content_type
       
       MIME type-like string that identifies how content is encoded for parts
       of this type. In most cases it corresponds to a particular XML
       sub-schema, although binary parts have other encoding schemes. As an
       example, the content type for a theme part is
       ``application/vnd.openxmlformats-officedocument.theme+xml``. Each
       part's content type is indicated in the content types item
       (``[Content_Types].xml``) located in the package root.
    
    .. attribute:: basename
       
       The root of the partname "filename" segment for parts of this type. For
       example, *basename* for slide layout parts is ``slideLayout`` and the
       partname for a slide layout matches the regular expression
       ``/ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``, e.g.
       ``/ppt/slideLayouts/slideLayout1.xml``.
       
       Note that while *basename* also usually corresponds to the base of the
       immediate parent "directory" name for tuple parts, this is not
       guaranteed. One example is theme parts, with partnames like
       ``/ppt/theme/theme1.xml``. Use *baseURI* to determine the "directory"
       portion of a partname.
    
    .. attribute:: ext
       
       The extension of the partname "filename" segment for parts of this
       type. For example, *ext* for a presentation part
       (``/ppt/presentation.xml``) is ``.xml``. Note that the leading period
       is included in the extension, consistent with the behavior of
       :func:`os.path.split`.
    
    .. attribute:: cardinality
       
       One of :attr:`pptx.spec.PTS_CARDINALITY_SINGLETON` or
       :attr:`pptx.spec.PTS_CARDINALITY_TUPLE`, corresponding to whether at
       most one or multiple parts of this type may appear in the package.
       ``/ppt/presentation.xml`` is an example of a singleton part.
       ``/ppt/slideLayouts/slideLayout4.xml`` is an example of a tuple part.
       The term *tuple* in this context is drawn from set theory in math and
       has no direct relationship to the Python tuple class.
    
    .. attribute:: required
       
       Boolean expressing whether at least one instance of this part type must
       appear in the package. ``presentation`` is an example of a required
       part type. ``notesMaster`` is an example of a optional part type.
    
    .. attribute:: baseURI
       
       The "directory" portion of the partname for parts of this type. The
       term *URI* is used because although part names (and other package item
       URIs) strongly resemble filesystem paths, and are readily operated on
       with functions from :mod:`os.path`, they have no direct correspondence
       to location in a file system (otherwise all packages would overwrite
       each other in the root directory :).
       
       For example, *baseURI* for slide layout parts is ``/ppt/slideLayouts``.
    
    .. attribute:: has_rels
       
       One of ``pptx.spec.PTS_HASRELS_ALWAYS``,
       ``pptx.spec.PTS_HASRELS_NEVER``, or ``pptx.spec.PTS_HASRELS_OPTIONAL``,
       indicating whether parts of this type always, never, or sometimes have
       relationships, respectively.
    
    .. attribute:: reltype
       
       The string used in the ``Type`` attribute of a ``Relationship`` XML
       element where a part of this content type is the target of the
       relationship. A relationship type is a URI string of the same form as a
       web page URL. For example, *reltype* for a part named
       ``/ppt/slides/slide1.xml`` would look something like
       ``http://schemas.openxmlformats.org/.../relationships/slide``.
    
    """
    __instances = {}
    
    def __new__(cls, content_type):
        """
        Only create new instance on first call for content_type. After that,
        use cached instance.
        """
        # if there's not an matching instance in the cache, create one
        if content_type not in cls.__instances:
            inst = super(PartTypeSpec, cls).__new__(cls)
            cls.__instances[content_type] = inst
        # return the instance; note that __init__() gets called either way
        return cls.__instances[content_type]
    
    def __init__(self, content_type):
        """Initialize spec attributes from constant values in pptx.spec."""
        # skip loading if this instance is from the cache
        if hasattr(self, '_loaded'):
            return
        # otherwise initialize new instance
        self._loaded = True
        if content_type not in pptx.spec.pml_parttypes:
            tmpl = "no content type '%s' in pptx.spec.pml_parttypes"
            raise KeyError(tmpl % content_type)
        ptsdict = pptx.spec.pml_parttypes[content_type]
        # load attributes from spec constants dictionary
        self.content_type = content_type            # e.g. 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
        self.basename     = ptsdict['basename']     # e.g. 'slideMaster'
        self.ext          = ptsdict['ext']          # e.g. '.xml'
        self.cardinality  = ptsdict['cardinality']  # e.g. PTS_CARDINALITY_SINGLETON or PTS_CARDINALITY_TUPLE
        self.required     = ptsdict['required']     # e.g. False
        self.baseURI      = ptsdict['baseURI']      # e.g. '/ppt/slideMasters'
        self.has_rels     = ptsdict['has_rels']     # e.g. PTS_HASRELS_ALWAYS, PTS_HASRELS_NEVER, or PTS_HASRELS_OPTIONAL
        self.reltype      = ptsdict['reltype']      # e.g. 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties'
    
    @property
    def format(self):
        """One of ``'xml'`` or ``'binary'``."""
        return 'xml' if self.ext == '.xml' else 'binary'
    


# ============================================================================
# Support Classes
# ============================================================================

class _ContentTypesItem(object):
    """
    Lookup content type by part name using dictionary syntax, e.g.
    ``content_type = cti['/ppt/presentation.xml']``.
    """
    def __init__(self):
        super(_ContentTypesItem, self).__init__()
        self.__defaults = None
        self.__overrides = None
    
    def __getitem__(self, partname):
        """
        Return the content type for the part with *partname*.
        
        """
        # raise exception if called before load()
        if self.__defaults is None or self.__overrides is None:
            tmpl = "lookup _ContentTypesItem['%s'] attempted before load"
            raise ValueError(tmpl % partname)
        # first look for an explicit content type
        if partname in self.__overrides:
            return self.__overrides[partname]
        # if not, look for a default based on the extension
        ext = os.path.splitext(partname)[1]            # get extension of partname
        ext = ext[1:] if ext.startswith('.') else ext  # with leading dot trimmed off
        if ext in self.__defaults:
            return self.__defaults[ext]
        # if neither of those work, raise an exception
        tmpl = "no content type for part '%s' in [Content_Types].xml"
        raise LookupError(tmpl % partname)
    
    def __len__(self):
        """
        Return sum count of Default and Override elements.
        
        """
        count = len(self.__defaults) if self.__defaults is not None else 0
        count += len(self.__overrides) if self.__overrides is not None else 0
        return count
    
    def compose(self, parts):
        """
        Assemble a [Content_Types].xml item based on the contents of *parts*.
        """
        # extensions in this dict include leading '.'
        def_cts = pptx.spec.default_content_types
        # initialize working dictionaries for defaults and overrides
        self.__defaults = dict((ext[1:], def_cts[ext]) for ext in ('.rels', '.xml'))
        self.__overrides = {}
        # compose appropriate element for each part
        for part in parts:
            ext = os.path.splitext(part.partname)[1]
            # if extension is '.xml', assume an override. There might be a
            # fancier way to do this, otherwise I don't know what 'xml'
            # Default entry is for.
            if ext == '.xml':
                self.__overrides[part.partname] = part.content_type
            elif ext in def_cts:
                self.__defaults[ext[1:]] = def_cts[ext]
            else:
                tmpl = "extension '%s' not found in default_content_types"
                raise LookupError(tmpl % (ext))
        return self
    
    @property
    def element(self):
        nsmap = {None: pptx.spec.nsmap['ct']}
        element = etree.Element(qtag('ct:Types'), nsmap=nsmap)
        if self.__defaults:
            for ext in sorted(self.__defaults.keys()):
                subelm = etree.SubElement(element, qtag('ct:Default'))
                subelm.set('Extension', ext)
                subelm.set('ContentType', self.__defaults[ext])
        if self.__overrides:
            for partname in sorted(self.__overrides.keys()):
                subelm = etree.SubElement(element, qtag('ct:Override'))
                subelm.set('PartName', partname)
                subelm.set('ContentType', self.__overrides[partname])
        return element
    
    def load(self, fs):
        """
        Retrieve [Content_Types].xml from specified file system and load it.
        Returns a reference to this _ContentTypesItem instance to allow
        generative call, e.g. ``cti = _ContentTypesItem().load(fs)``.
        
        """
        element = fs.getelement('/[Content_Types].xml')
        defaults = element.findall(qtag('ct:Default'))
        overrides = element.findall(qtag('ct:Override'))
        self.__defaults = dict((d.get('Extension'), d.get('ContentType'))
                               for d in defaults)
        self.__overrides = dict((o.get('PartName'), o.get('ContentType'))
                                for o in overrides)
        return self
    


# ============================================================================
# FileSystem Classes
# ============================================================================

class FileSystem(object):
    """
    Factory for filesystem interface instances.
    
    A FileSystem object provides access to on-disk package items via their URI
    (e.g. ``/_rels/.rels`` or ``/ppt/presentation.xml``). This allows parts to
    be accessed directly by part name, which for a part is identical to its
    item URI. The complexities of translating URIs into file paths or zip item
    names, and file and zip file access specifics are all hidden by the
    filesystem class. :class:`FileSystem` acts as the Factory, returning the
    appropriate concrete filesystem class depending on what it finds at *path*.
    """
    def __new__(cls, path):
        try:
            # if path is a path to a zip file or a valid file-like object,
            # return instance of ZipFileSystem
            if is_zipfile(path):
                instance = ZipFileSystem(path)
            # if path is a directory, return instance of DirectoryFileSystem
            elif os.path.isdir(path):
                instance = DirectoryFileSystem(path)
            # otherwise, something's not right, raise exception
            else:
                raise PackageNotFoundError("Package not found at '%s'" % (path))
        except TypeError:
            # in case we're dealing with a file-like object
            raise PackageNotFoundError("Package not found")
        return instance


class BaseFileSystem(object):
    """
    Base class for FileSystem classes, providing common methods.
    """
    def __init__(self, path):
        try:
            self.__path = os.path.abspath(path)
        except AttributeError:
            self.__path = path

    def __contains__(self, itemURI):
        """
        Allows use of 'in' operator to test whether an item with the specified
        URI exists in this filesystem.
        """
        return itemURI in self.itemURIs

    def getblob(self, itemURI):
        """Return byte string of item identified by *itemURI*."""
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        stream = self.getstream(itemURI)
        blob = stream.read()
        stream.close()
        return blob

    def getelement(self, itemURI):
        """
        Return ElementTree element of XML item identified by *itemURI*.
        """
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        stream = self.getstream(itemURI)
        try:
            parser = etree.XMLParser(remove_blank_text=True)
            element = etree.parse(stream, parser).getroot()
        except etree.XMLSyntaxError:
            raise NotXMLError("package item %s is not XML" % itemURI)
        stream.close()
        return element

    @property
    def path(self):
        """
        Path property is read-only. Need to instantiate a new FileSystem
        object to access a different package.
        """
        return self.__path


class DirectoryFileSystem(BaseFileSystem):
    """
    Provides access to package members that have been expanded into an on-disk
    directory structure.
    
    Inherits __contains__(), getelement(), and path from BaseFileSystem.
    
    """
    def __init__(self, path):
        """
        *path* is the path to a directory containing an expanded package.
        """
        super(DirectoryFileSystem, self).__init__(path)
    
    def close(self):
        """
        Provides interface consistency with |ZipFileSystem|, but does nothing,
        a directory file system doesn't need closing.
        """
        pass
    
    def getstream(self, itemURI):
        """
        Return file-like object containing package item identified by
        *itemURI*. Remember to call close() on the stream when you're done
        with it to free up the memory it uses.
        """
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        path = os.path.join(self.path, itemURI[1:])
        with open(path, 'rb') as f:
            stream = StringIO(f.read())
        return stream
    
    @property
    def itemURIs(self):
        """
        Return list of all filenames under filesystem root directory,
        formatted as item URIs. Each URI is the relative path of that file
        with a leading slash added, e.g. '/ppt/slides/slide1.xml'. Although
        not strictly necessary, the results are sorted for neatness' sake.
        
        """
        itemURIs = []
        for dirpath, dirnames, filenames in os.walk(self.path):
            for filename in filenames:
                item_path = os.path.join(dirpath, filename)
                itemURI = item_path[len(self.path):]  # leaves a leading slash on
                itemURIs.append(itemURI.replace(os.sep, '/'))
        return sorted(itemURIs)
    

class ZipFileSystem(BaseFileSystem):
    """
    Return new instance providing access to zip-format OPC package at *path*.

    *path* can be either a path to a file (a string) or a file-like object.

    If mode is 'w', creates a new zip archive at *path*. If a file by that name
    already exists, it is truncated.

    Inherits :meth:`__contains__`, :meth:`getelement`, and :attr:`path` from
    BaseFileSystem.
    """
    def __init__(self, path, mode='r'):
        super(ZipFileSystem, self).__init__(path)
        if 'w' in mode:
            self.zipf = ZipFile(self.path, 'w', compression=ZIP_DEFLATED)
        else:
            self.zipf = ZipFile(self.path, 'r')

    def close(self):
        """Close the zip file"""
        self.zipf.close()

    def getstream(self, itemURI):
        """
        Return file-like object containing package item identified by
        *itemURI*. Remember to call close() on the stream when you're done
        with it to free up the memory it uses.
        """
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        membername = itemURI[1:]  # trim off leading slash
        stream = StringIO(self.zipf.read(membername))
        return stream

    @property
    def itemURIs(self):
        """
        Return list of archive members formatted as item URIs. Each member
        name is the archive-relative path of that file. A forward-slash is
        prepended to form the URI, e.g. '/ppt/slides/slide1.xml'. Although
        not strictly necessary, the results are sorted for neatness' sake.
        """
        names = self.zipf.namelist()
        # zip archive can contain entries for directories, so get rid of those
        itemURIs = [('/%s' % name) for name in names if not name.endswith('/')]
        return sorted(itemURIs)

    def write_blob(self, blob, itemURI):
        """
        Write *blob* to zip file as binary stream named *itemURI*.
        """
        if itemURI in self:
            tmpl = "Item with URI '%s' already in package"
            raise DuplicateKeyError(tmpl % itemURI)
        membername = itemURI[1:]  # trim off leading slash
        self.zipf.writestr(membername, blob)

    def write_element(self, element, itemURI):
        """
        Write *element* to zip file as an XML document named *itemURI*.
        """
        if itemURI in self:
            tmpl = "Item with URI '%s' already in package"
            raise DuplicateKeyError(tmpl % itemURI)
        membername = itemURI[1:]  # trim off leading slash
        xml = etree.tostring(element, encoding='UTF-8', pretty_print=True,
                             standalone=True)
        xml = prettify_nsdecls(xml)
        self.zipf.writestr(membername, xml)



# ============================================================================
# Utility functions
# ============================================================================

def prettify_nsdecls(xml):
    """
    Wrap and indent second and later attributes on the root element so
    namespace declarations don't run off the page in the text editor and can
    be more easily inspected.
    """
    lines = xml.splitlines()
    # if entire XML document is all on one line, don't mess with it
    if len(lines) < 2:
        return xml
    # if don't find xml declaration on first line, bail
    if not lines[0].startswith('<?xml'):
        return xml
    # if don't find an unindented opening element on line 2, bail
    if not lines[1].startswith('<'):
        return xml
    rootline = lines[1]
    # split rootline into element tag part and attributes parts
    attrib_re = re.compile(r'([-a-zA-Z0-9_:.]+="[^"]*" *>?)')
    substrings = [substring.strip() for substring in attrib_re.split(rootline)
                                     if substring]
    # substrings look something like:
    # ['<p:sld', 'xmlns:p="html://..."', 'name="Office Theme>"']
    # if there's only one attribute there's no need to wrap
    if len(substrings) < 3:
        return xml
    indent = ' ' * (len(substrings[0])+1)
    # join element tag and first attribute onto same line
    newrootline = ' '.join(substrings[:2])
    # indent remaining attributes on following lines
    for substring in substrings[2:]:
        newrootline += '\n%s%s' % (indent, substring)
    lines[1] = newrootline
    return '\n'.join(lines)


