# -*- coding: utf-8 -*-
#
# packaging.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

'''Code that deals with packaging the various parts of the PowerPoint
presentation into a .pptx file.'''

import os
import re
import zipfile

from lxml import etree

import pptx.spec

from pptx            import util
from pptx.exceptions import CorruptedTemplateError

# ============================================================================
# Class catalog
# ============================================================================

# API Classes
# ----------------------------------------------------------------------------
# Package
# PartType
# PartTypes
# TemplatePackage

# Package Items
# ----------------------------------------------------------------------------
# PackageItem
# ContentTypesItem
# RelationshipItem
# PackageRelationshipItem
# PartRelationshipItem

# Part Collections
# ----------------------------------------------------------------------------
# PartCollection
# ImageParts
# SlideLayoutParts
# SlideMasterParts
# SlideParts
# ThemeParts

# Parts
# ----------------------------------------------------------------------------
# Part
# ImagePart
# PresentationPart
# PresPropsPart
# SlideLayoutPart
# SlideMasterPart
# SlidePart
# TableStylesPart
# ThemePart
# ViewPropsPart

# Relationships
# ----------------------------------------------------------------------------
# PartRelationship
# RelationshipElement
 
# Files
# ----------------------------------------------------------------------------
# TemplateFiles
# PackageItemFile
# PackageItemFileCollection
# PartFile
# PartFiles
# PartRelsFiles


# ============================================================================
# Package
# ============================================================================
# Start with media parts (they don't have a relationship item)
# -------------------------------------------------------------------
# (/) image parts  #TECHDEBT: Need to distinguish images from other media by extension in ext2mime_map
# audio parts
# video parts

# Then template parts
# -------------------------------------------------------------------
# (/) theme parts
# (/) slide master parts
# (/) slide layout parts
# (/) table styles part
# (/) view properties part
# handout master part
# notes master part
# theme override parts

# Then document property parts
# -------------------------------------------------------------------
# file property parts
# thumbnail part?

# Then presentation parts
# -------------------------------------------------------------------
# (/) slide parts
# notes slide parts
# (/) presentation properties part
# (/) presentation part
# ============================================================================

class Package(object):
    
    def __init__(self, presentation):
        
        # initialize part collections
        self.imageparts       = ImageParts       (self)
        self.themeparts       = ThemeParts       (self)
        self.slidemasterparts = SlideMasterParts (self)
        self.slidelayoutparts = SlideLayoutParts (self)
        self.slideparts       = SlideParts       (self)
        
        # initialize relationship item
        self.relationshipitem = PackageRelationshipItem(self)
        self.contenttypesitem = ContentTypesItem(self)
        
#TECHDEBT: Add images using their references from other parts so that only
#          referenced ones are written to the package
        # load collections from presentation
        for item in presentation.images       : self.imageparts       .additem(item)
        for item in presentation.themes       : self.themeparts       .additem(item)
        for item in presentation.slidemasters : self.slidemasterparts .additem(item)
        for item in presentation.slidelayouts : self.slidelayoutparts .additem(item)
        for item in presentation.slides       : self.slideparts       .additem(item)
        # load non-collection parts from presentation
        self.presentationpart  = PresentationPart (self, presentation)
#TECHDEBT: Need to make these two parts work when no file exists for them in the template
        # self.handoutmasterpart = HandoutMasterPart(self, presentation.presprops)
        # self.notesmasterpart   = NotesMasterPart(self, presentation.presprops)
        self.prespropspart     = PresPropsPart    (self, presentation.presprops)
        self.tablestylespart   = TableStylesPart  (self, presentation.tablestyles)
        self.viewpropspart     = ViewPropsPart    (self, presentation.viewprops)
        
        # compile list of all the package parts
        self.parts = []
        # singleton parts
#TECHDEBT: Need to make these two parts work when no file exists for them in the template
        # self.parts.append(self.handoutmasterpart)
        # self.parts.append(self.notesmasterpart)
        self.parts.append(self.presentationpart)
        self.parts.append(self.prespropspart)
        self.parts.append(self.tablestylespart)
        self.parts.append(self.viewpropspart)
        # collection parts
        self.parts.extend(self.imageparts)
        self.parts.extend(self.slidemasterparts)
        self.parts.extend(self.slidelayoutparts)
        self.parts.extend(self.slideparts)
        self.parts.extend(self.themeparts)
        
        # relationships are implemented by individual Part classes
        
    
    def __normalizedfilename(self, filename):
        # add .pptx extension to filename if it doesn't have one
        filename = filename if filename[-5:] == '.pptx' else filename + '.pptx'
        return filename
    
    @property
    def items(self):
        items = []
        items.extend(self.parts)              # parts (every part is a package item)
        for part in self.parts:               # Part relationship items
            if part.relationshipitem:
                # print "%s for %s" % (part.relationshipitem.__class__.__name__, part.filename)
                items.append(part.relationshipitem)
            else:
                pass  # just need this when logging line below is commented out
                # print "No relationship item for %s" % part.filename
        items.append(self.relationshipitem)   # package relationship item
        items.append(self.contenttypesitem)   # content types item
        return items
    
    @property
    def relatedparts(self):
        relatedparts = []
        # relatedparts.append(self.coredocpropspart)
        # relatedparts.append(self.appdocpropspart)
        # relatedparts.append(self.thumbnailpart)
        relatedparts.append(self.presentationpart)
        return relatedparts
    
    #REFACTOR: Consider renaming this property to relationshipelements to
    #          make clear it's a list of etree.Element
    @property
    def relationships(self):
        relationships = []
        for idx, relatedpart in enumerate(self.relatedparts):
            rId              = idx+1
            relationshiptype = relatedpart.parttype.relationshiptype
            targetdir        = relatedpart.parttype.pkgreldir
            targetfilename   = relatedpart.filename
            targetpath       = os.path.join(targetdir, targetfilename)
            relationships.append(RelationshipElement(rId, relationshiptype, targetpath))
        return relationships
    
    def save(self, filename):
        items = self.items
        # create the zip file to hold the presentation package
        pptxfile = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)
        # print 'Package contains %d items.' % len(items)
        for package_item in items:
            # print package_item.__class__
            package_item.write(pptxfile)
        pptxfile.close()
    


# ============================================================================
# PartType
# ============================================================================
#REFACTOR: Refactor this and spec.parttypes so spec doesn't call PartType
#          (it leads to a circular import problem). Make spec.parttypes a list
#          of dictionary, and have PartType read it into a class variable when
#          the class is initialized (on module load).

class PartType(object): 
    
    def __init__(self, rootname, file_ext, name, cardinality, required,
                 location, has_rel_item, rels_from, content_type,
                 namespace, relationship):
        self.rootname         = rootname       # e.g. 'core'
        self.file_ext         = file_ext       # e.g. 'xml'
        self.name             = name           # e.g. 'Core File Properties Part'
        self.cardinality      = cardinality    # e.g. 'single' or 'multiple'
        self.required         = required       # e.g. False
        self.location         = location       # e.g. '/docProps'
        self.has_rel_item     = has_rel_item   # e.g. 'always', 'never', or 'optional'
        self.rels_from        = rels_from      # e.g. ['package']
        self.content_type     = content_type   # e.g. 'application/vnd.openxmlformats-package.core-properties+xml'
        self.namespace        = namespace      # e.g. 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        self.relationshiptype = relationship   # e.g. 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties'
    
    @property
    def key(self):
        return self.rootname
    
    @property
    def pkgreldir(self):
        return self.location[1:] if self.location.startswith('/') else self.location
    
    # NOTE: This doesn't work for the package relationships item, only for
    # part relationship items!
    # Return the part path format used in part relationship items (files),
    # a path that is relative to the directory containing the frompart.
    def reltargetdir(self, fromlocation):
        if fromlocation == '/ppt':  # that means this is the presentation part, the only part with that path that has a relationship item
            if self.location == '/ppt':              # means topart is a singleton part
                return ''
            elif self.location.startswith('/ppt/'):  # means topart is a collection part
                return self.location[5:]
            else:
                raise NotImplementedError('''Unrecognized fromlocation '%s' received by PartType.reltargetdir()''' % fromlocation)
        if self.location.startswith('/ppt/'):
            return '../%s' % self.location[5:]
        raise NotImplementedError("""PartType.reltargetdir() is not yet implemented for part type '%s'""" % self.rootname)
    
    @property
    def zipdir(self):
        return self.location[1:] if self.location.startswith('/') else self.location
    


# ============================================================================
# PartTypes
# ============================================================================
#REFACTOR: See if PartTypes and PartType can be combined into a single class.

class PartTypes(object):
    
    def __init__(self):
        raise NotImplementedError('''The PartTypes class is not meant to be instantiated, only used for class methods''')
    
    @classmethod
    def classparttype(klass, classname):
        map = { 'Image'        : 'image'
              , 'Presentation' : 'presentation'
              , 'PresProps'    : 'presProps'
              , 'Slide'        : 'slide'
              , 'SlideLayout'  : 'slideLayout'
              , 'SlideMaster'  : 'slideMaster'
              , 'TableStyles'  : 'tableStyles'
              , 'Theme'        : 'theme'
              , 'ViewProps'    : 'viewProps'
              }
        if classname not in map:
            raise KeyError("""PartTypes.classparttype() lookup failed with key '%s'""" % (classname))
        return PartTypes.parttype(map[classname])
    
    @classmethod
    def parttype(klass, parttypekeyname):
        for parttype in pptx.spec.parttypes:
            if parttype.rootname == parttypekeyname:
                return parttype
        raise KeyError("""PartTypes.parttype() lookup failed with key '%s'""" % (parttypekeyname))
    


# ============================================================================
# TemplatePackage
# ============================================================================
# #TODO: Enhance this so it can read a template from a zip archive without
#        requiring it to be expanded out into a directory.
# ============================================================================

class TemplatePackage(Package):
    
    # NOTE: Not sure why this doesn't call Package.__init__(), might need to
    #       distinguish between PresentationPackage and TemplatePackage, with
    #       Package being the superclass for each of them.
    def __init__(self, templatedir):
        self.templatedir   = templatedir
        self.templatefiles = TemplateFiles(templatedir)
    
    @property
    def relationships(self):
        relationships = []
        for relsfile in self.templatefiles.relsfiles:
            relationships.extend(self.__parserelsfile(relsfile))
        return relationships
    
    def __parserelsfile(self, relsfile):
        relationships = []
        tree = etree.parse(relsfile.path)
        relationshipelements = tree.getroot().findall(pptx.spec.qname('pr','Relationship'))
        for relationshipelement in relationshipelements:
            relationships.append(PartRelationship.parse(relsfile.partfile, relationshipelement))
        return relationships
    


# ============================================================================
# PackageItem
# ============================================================================

class PackageItem(object):
    
    # must be overridden by items that are not XML-based
    def write(self, pptx_zipfile):
        pptx_zipfile.writestr(self.zipfilepath, self.xmlstring)
    
    @property
    def xmlstring(self):
#TECHDEBT: Maybe should check to see if item content is XML or not and throw exception if not XML
        element = self.element
        if element is None:
            return None
        return util.prettify_nsdecls(etree.tostring(element, encoding='UTF-8', pretty_print=True, standalone=True))
    
    @property
    def zipfilepath(self):
        return os.path.join(self.zipdir, self.filename)
    


# ============================================================================
# ContentTypesItem
# ============================================================================

class ContentTypesItem(PackageItem):
    
    def __init__(self, package):
        PackageItem.__init__(self)
        self.package  = package
        self.filename = '[Content_Types].xml'
        self.zipdir   = ''
    
    @property
    def element(self):
        #TECHDEBT: Look this up from lookup table in pptx.spec
        element = etree.fromstring('''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''')
        element.extend(self.__defaultcontenttypes)
        element.extend(self.__overrides)
        return element
    
#TODO: account for possible thumbnail.jpeg image in /docProps
#TODO: account for possible printerSettings[1-9][0-9]*.bin files in ppt/printerSettings/
#TODO: account for possible fntdata parts (handled by fntdata Default element) in /ppt/fonts directory
    @property
    def __defaultcontenttypes(self):
        defaults = []
        defaults.extend(self.__mediatypedefaultelements)
        defaults.append(self.__defaultelement('rels'))
        defaults.append(self.__defaultelement('xml'))
        return defaults
    
    def __defaultelement(self, ext):
        #REFACTOR: Work out a more elegant access method to this ext2mime_map.
        mimetype = pptx.spec.ext2mime_map[ext]
        element = etree.Element('Default')
        element.set('Extension'   , ext)
        element.set('ContentType' , mimetype)
        return element
    
#TECHDEBT: Need to accomodate audio and video content types as well as images
    @property
    def __mediatypedefaultelements(self):
        defaults = []
        exts = {}
        for imagepart in self.package.imageparts:
            ext = os.path.splitext(imagepart.filename)[1][1:]  # extension is second item in returned tuple and need to strip off '.' from the front
            exts[ext] = ''
        return [self.__defaultelement(ext) for ext in sorted(exts.keys())]
    
    def __override_element(self, part):
        partname = os.path.join(part.parttype.location, part.filename)
        element = etree.Element('Override')
        element.set('PartName'    , partname)
        element.set('ContentType' , part.parttype.content_type)
        return element
    
    @property
    def __overrides(self):
        return [self.__override_element(part) for part in self.package.parts if part.filename.endswith('.xml')]
    


# ============================================================================
# RelationshipItem
# ============================================================================

class RelationshipItem(PackageItem):
    
    def __init__(self, fromitem):
        PackageItem.__init__(self)
        self.fromitem = fromitem
    
    @property
    def element(self):
        relationships = self.fromitem.relationships
        if not relationships:
            return None
#TECHDEBT: Root element and namespace should be lookups in ItemTypes (ItemTypes.itemtype('relationship'))
        element = etree.fromstring('''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>''')
        for relationship in relationships:
            element.append(relationship.element)
        return element
    


# ============================================================================z
# PackageRelationshipItem
# ============================================================================

class PackageRelationshipItem(RelationshipItem):
    
    def __init__(self, package):
        RelationshipItem.__init__(self, package)
        self.filename = '.rels'
        self.zipdir   = '_rels'
    


# ============================================================================z
# PartRelationshipItem
# ============================================================================

class PartRelationshipItem(RelationshipItem):
    
    def __init__(self, part):
        RelationshipItem.__init__(self, part)
        self.part = part
    
    @property
    def filename(self):
        return '%s.rels' % self.part.filename
    
    @property
    def zipdir(self):
        return os.path.join(self.part.parttype.zipdir, '_rels')
    


# ============================================================================
# PartCollection
# ============================================================================

class PartCollection(list):
    
    def __init__(self, package):
        list.__init__(self)
        self.package = package
        self.__dict = {}
    
    def additem(self, item):
        if item.key in self.__dict:
            return self.__dict[item.key]
        part = self.memberclass(self, item)
        self.__dict[item.key] = part
        self.append(part)
        return part
    
    # Can implement a __contains__(self, value) method if anyone needs to test
    # whether something is already in the collection.
    def getitem(self, key):
        if key in self.__dict:
            return self.__dict[key]
        raise KeyError("""getitem() lookup in %s failed with key '%s'""" % (self.__class__.__name__, key))
    


# ============================================================================
# ImageParts
# ============================================================================

class ImageParts(PartCollection):
    
    def __init__(self, package):
        PartCollection.__init__(self, package)
        self.memberclass = ImagePart
    


# ============================================================================
# SlideLayoutParts
# ============================================================================

class SlideLayoutParts(PartCollection):
    
    def __init__(self, package):
        PartCollection.__init__(self, package)
        self.memberclass = SlideLayoutPart
    


# ============================================================================
# SlideMasterParts
# ============================================================================

class SlideMasterParts(PartCollection):
    
    def __init__(self, package):
        PartCollection.__init__(self, package)
        self.memberclass = SlideMasterPart
    


# ============================================================================
# SlideParts
# ============================================================================

class SlideParts(PartCollection):
    
    def __init__(self, package):
        PartCollection.__init__(self, package)
        self.memberclass = SlidePart
    


# ============================================================================
# ThemeParts
# ============================================================================

class ThemeParts(PartCollection):
    
    def __init__(self, package):
        PartCollection.__init__(self, package)
        self.memberclass = ThemePart
    


# ============================================================================
# Part
# ============================================================================
#REFACTOR: Also some inconsistency here between (plain-old) parts and
#          collection parts, in particular having parent being assigned by the
#          Part super class. [Update]: I'm not so sure, I'm thinking all Parts
#          need access to the package in order to look up their possible
#          relationships. idx method might be the only difference between the
#          po parts and collection parts.

class Part(PackageItem):
    
    def __init__(self, parent, item):
        self.parttype = PartTypes.classparttype(item.__class__.__name__)
        self.parent   = parent
        self.package  = parent.package
        self.item     = item
        self.path     = item.path
    
    @property
    def filename(self):
        cardinality = self.parttype.cardinality
        if cardinality == 'single':
            return "%s.%s" % (self.parttype.rootname, self.parttype.file_ext)
        elif cardinality == 'multiple':
            return "%s%d.%s" % (self.parttype.rootname, self.idx+1, self.parttype.file_ext)
        else:
            raise ValueError('''Invalid value for PartType.cardinality in Part.filename call. Expected 'single' or 'multiple', got %s''' % cardinality)
    
    @property
    def idx(self):
        if self.parttype.cardinality == 'multiple':
            return self.parent.index(self)
        else:
            raise NotImplementedError('''Part.idx is not defined for parts with cardinality!='multiple'. Called on class '%s'.''' % self.__class__.__name__)
    
    @property
    def imageparts(self):
        return [self.package.imageparts.getitem(key=image.key) for image in self.item.images]
    
    # default is the part has no outward relationships
    @property
    def relatedparts(self):
        return []
    
    @property
    def relationshipitem(self):
        return PartRelationshipItem(self) if self.relationships else None
    
    #REFACTOR: Consider renaming this property to relationshipelements to
    #          make clear it's a list of etree.Element
    #REFACTOR: See if this can be factored up to a higher-level class. There's
    #          a very similar property method by the same name in Package used
    #          for the package relationship item.
    @property
    def relationships(self):
        relationships = []
        for idx, relatedpart in enumerate(self.relatedparts):
            rId              = idx+1
            relationshiptype = relatedpart.parttype.relationshiptype
            targetdir        = relatedpart.parttype.reltargetdir(self.parttype.location)
            targetfilename   = relatedpart.filename
            targetpath       = os.path.join(targetdir, targetfilename)
            relationships.append(RelationshipElement(rId, relationshiptype, targetpath))
        return relationships
    
    @property
    def handoutmasterparts(self):
        try:
            handoutmaster = self.item.handoutmaster
        except AttributeError:
            return []
        return [self.package.handoutmasterparts.getitem(key=handoutmaster.key)] if handoutmaster else []
    
    @property
    def notesmasterparts(self):
        try:
            notesmaster = self.item.notesmaster
        except AttributeError:
            return []
        return [self.package.notesmasterparts.getitem(key=notesmaster.key)] if notesmaster else []
    
    # presProps.xml is not optional, so leaving this to throw an exception if
    # for some reason it's not found.
    @property
    def prespropspart(self):
        return [self.package.prespropspart]
    
    # NOTE: Careful, the singular and plural form of this part are the same
    #       i.e. printersettings refers to both a printersettings part as well
    #       as a list of printersettings parts. Hoping this doesn't screw us
    #       up later but waiting for it to become a problem before inventing a
    #       solution.
    @property
    def printersettingsparts(self):
        try:
            printersettings = self.item.printersettings
        except AttributeError:
            return []
        return [self.package.printersettingsparts.getitem(key=printersettings.key) for printersettings in self.item.printersettings]
    
    @property
    def slidelayoutpart(self):
        return [self.package.slidelayoutparts.getitem(key=self.item.slidelayout.key)] if self.item.slidelayout else []
    
    @property
    def slidelayoutparts(self):
        return [self.package.slidelayoutparts.getitem(key=slidelayout.key) for slidelayout in self.item.slidelayouts]
    
    @property
    def slidemasterpart(self):
        return [self.package.slidemasterparts.getitem(key=self.item.slidemaster.key)] if self.item.slidemaster else []
    
    @property
    def slidemasterparts(self):
        return [self.package.slidemasterparts.getitem(key=slidemaster.key) for slidemaster in self.item.slidemasters]
    
    @property
    def slideparts(self):
        try:
            slides = self.item.slides
        except AttributeError:
            return []
        return [self.package.slideparts.getitem(key=slide.key) for slide in self.item.slides]
    
    @property
    def tablestylespart(self):
        return [self.package.tablestylespart] if self.item.tablestyles else []
    
    @property
    def themepart(self):
        return [self.package.themeparts.getitem(key=self.slidemaster.theme.key)] if self.item.theme else []
    
    @property
    def themeparts(self):
        try:
            themes = self.item.themes
        except AttributeError:
            return []
        return [self.package.themeparts.getitem(key=theme.key) for theme in self.item.themes]
    
    @property
    def viewpropspart(self):
        return [self.package.viewpropspart] if self.item.viewprops else []
    
#TECHDEBT: Not sure the writing bit is all squared away, check it out.
    # must be overridden by XML-based items that are not part of the template
    # and so therefore need to be written from the XML provided by the
    # presentation
    def write(self, pptx_zipfile):
        pptx_zipfile.write(self.path, self.zipfilepath)
    
    @property
    def zipfilepath(self):
        return os.path.join(self.parttype.zipdir, self.filename)
    


# ============================================================================
# ImagePart
# ============================================================================

class ImagePart(Part):
    
    def __init__(self, parent, item):
        Part.__init__(self, parent, item)
#TECHDEBT: These aren't populated yet, not sure they need to be unless we want
#          to cull image files that don't have any references from other
#          parts. In any case, if they are implemented it needs to be with
#          property methods so any possible infinite recursion is avoided on
#          package load (all other parts need to be loaded before setting
#          these relationships).
        self.themes       = []
        self.slidemasters = []
        self.slidelayouts = []
        self.slides       = []
    
    @property
    def filename(self):
        ext = os.path.splitext(os.path.basename(self.path))[1]
        return "%s%d%s" % (self.parttype.rootname, self.idx+1, ext)
    
    def write(self, pptx_zipfile):
        pptx_zipfile.write(self.path, self.zipfilepath)
    


# ============================================================================
# PresentationPart
# ============================================================================

class PresentationPart(Part):
    
    # NOTE: Inherits from Part but doesn't call Part.__init__(), pending refactoring of Part into Part and CollectionPart
    def __init__(self, parent, item):
        # Part.__init__(self, parent, item)  # superclass needs refactoring because it only understands collection parts
        self.parttype = PartTypes.classparttype(item.__class__.__name__)
        # self.parent       = parent      # don't think we have any takers for this property since it's not a collection part
        self.package      = parent
        self.item         = item
        self.presentation = item
    
    @property
    def element(self):
        return self.presentation.element
    
    @property
    def relatedparts(self):
        relatedparts = []
        relatedparts.extend(self.slidemasterparts)     # --+-- being first and maintaining sequence of these
        relatedparts.extend(self.notesmasterparts)     #   |   four part collections is critical to making
        relatedparts.extend(self.handoutmasterparts)   #   |   rIds in presentation.xml sync with those
        relatedparts.extend(self.slideparts)           # --+   in presentation.xml.rels
        relatedparts.extend(self.printersettingsparts)
        relatedparts.extend(self.prespropspart)
        relatedparts.extend(self.viewpropspart)
        relatedparts.extend(self.themeparts)
        relatedparts.extend(self.tablestylespart)
        return relatedparts
    
#TECHDEBT: Check this out to make sure, was very sleepy when I did this ...
    # overridden to write source XML
    def write(self, pptx_zipfile):
        pptx_zipfile.writestr(self.zipfilepath, self.xmlstring)
    


# ============================================================================
# PresPropsPart
# ============================================================================

class PresPropsPart(Part):
    
    # NOTE: Inherits from Part but doesn't call Part.__init__(), pending refactoring of Part into Part and CollectionPart
    def __init__(self, parent, item):
        # Part.__init__(self, parent, item)  # superclass needs refactoring because it only understands collection parts
        self.parttype = PartTypes.classparttype(item.__class__.__name__)
        # self.parent   = parent
        self.item      = item
        self.presprops = item
        self.package   = parent
        self.path      = item.path
    


# ============================================================================
# SlideLayoutPart
# ============================================================================

class SlideLayoutPart(Part):
    
    def __init__(self, parent, item):
        Part.__init__(self, parent, item)
        self.slidelayout = item
    
    @property
    def relatedparts(self):
        relatedparts = []
        relatedparts.extend(self.slidemasterpart)
        relatedparts.extend(self.imageparts)
        return relatedparts
    


# ============================================================================
# SlideMasterPart
# ============================================================================

class SlideMasterPart(Part):
    
    def __init__(self, parent, item):
        Part.__init__(self, parent, item)
        self.slidemaster = item
    
#REFACTOR: make relatedparts a list of properties and let high-level code work
#          out what's a list and what's a scalar value or None.
#          e.g. return [self.imageparts, self.slidelayoutparts, self.themepart]
    @property
    def relatedparts(self):
        relatedparts = []
        relatedparts.extend(self.imageparts)
        relatedparts.extend(self.slidelayoutparts)
        relatedparts.extend(self.themepart)
        return relatedparts
    


# ============================================================================
# SlidePart
# ============================================================================

# ----------------------------------------------------------------------------
#REFACTOR: Might want to create a superclass PresentationPart for this to
#          inherit from to give it some of the automatic behaviors that
#          TemplatePart subclasses have. Will be more useful when additional
#          presentation parts such as notesSlide and perhaps comment are
#          added.
# ----------------------------------------------------------------------------

class SlidePart(Part):
    
    # NOTE: SlidePart inherits from Part, but does not call __init__() on it.
    def __init__(self, parent, item):
        # Part.__init__(self, parent, item)  # superclass needs refactoring because it only understands collection parts
        self.parttype = PartTypes.parttype('slide')
        self.parent   = parent
        self.package  = parent.package
        self.item     = item
        self.slide    = item
    
    @property
    def element(self):
        return self.slide.element
    
    @property
    def relatedparts(self):
        relatedparts = []
        relatedparts.extend(self.imageparts)
        relatedparts.extend(self.slidelayoutpart)
        return relatedparts
    
#TECHDEBT: Check this out to make sure, was very sleepy when I did this ...
    # overridden to write source XML
    def write(self, pptx_zipfile):
        pptx_zipfile.writestr(self.zipfilepath, self.xmlstring)
    


# ============================================================================
# TableStylesPart
# ============================================================================

class TableStylesPart(Part):
    
    # NOTE: Inherits from Part but doesn't call Part.__init__(), pending refactoring of Part into Part and CollectionPart
    def __init__(self, parent, item):
        # Part.__init__(self, parent, item)  # superclass needs refactoring because it only understands collection parts
        self.parttype = PartTypes.classparttype(item.__class__.__name__)
        # self.parent   = parent
        self.tablestyles = item
        self.package     = parent
        self.path        = item.path
    


# ============================================================================
# ThemePart
# ============================================================================

class ThemePart(Part):
    
    def __init__(self, parent, item):
        Part.__init__(self, parent, item)
        self.theme = item
    
    @property
    def relatedparts(self):
        return self.imageparts
    
    def write(self, pptx_zipfile):
        pptx_zipfile.write(self.path, self.zipfilepath)
    


# ============================================================================
# ViewPropsPart
# ============================================================================

class ViewPropsPart(Part):
    
    # NOTE: Inherits from Part but doesn't call Part.__init__(), pending refactoring of Part into Part and CollectionPart
    def __init__(self, parent, item):
        # Part.__init__(self, parent, item)  # superclass needs refactoring because it only understands collection parts
        self.parttype = PartTypes.classparttype(item.__class__.__name__)
        # self.parent   = parent
        self.viewprops = item
        self.package   = parent
        self.path      = item.path
    


# ============================================================================
# PartRelationship
# ============================================================================
# Relationship between two package parts
# ============================================================================

class PartRelationship(object):
    
    def __init__(self, frompartfile, relationshiptype, target):
        self.fromtype = frompartfile.parttype.key
        self.frompath = frompartfile.path
        self.totype   = relationshiptype.split('/')[-1]
        self.topath   = os.path.normpath(os.path.join(frompartfile.dirpath, target))
        
        # self.__frompartfile     = frompartfile
        # self.__relationshiptype = relationshiptype
        # self.__target           = target
        # self.fromparttype       = frompartfile.parttype
        # self.toparttype         = PartTypes.parttype(self.totype)
        # print 'Relationship:  %-12s ===>  %-12s  %s' % (self.fromtype, self.totype, self.topath)
    
    @classmethod
    def parse(klass, frompartfile, relationshipelement):
        reltype = relationshipelement.get('Type'  )
        target  = relationshipelement.get('Target')
        return PartRelationship(frompartfile, reltype, target)
    


# ============================================================================
# RelationshipElement
# ============================================================================
# Relationship element found in package and part relationship items.
# ============================================================================

class RelationshipElement(object):
    
    def __init__(self, idnumber, relationshiptype, target):
        self.idnumber         = idnumber
        self.relationshiptype = relationshiptype
        self.target           = target
    
    @property
    def element(self):
        element = etree.Element('Relationship')
        element.set('Id'     , 'rId%d' % self.idnumber )
        element.set('Type'   , self.relationshiptype   )
        element.set('Target' , self.target             )
        return element
    


# ============================================================================
# TemplateFiles
# ============================================================================
# Discover all the template files in the specified package directory tree
# and make them available with a high-level interface. Essential role is to
# hide complexities of the package directory layout and file system access.
# ============================================================================

class TemplateFiles(object):
    
    def __init__(self, templatedir):
        self.templatedir = templatedir
        
        self.imagefiles       = PartFiles(templatedir, 'image'       )
        self.themefiles       = PartFiles(templatedir, 'theme'       )
        self.slidemasterfiles = PartFiles(templatedir, 'slideMaster' )
        self.slidelayoutfiles = PartFiles(templatedir, 'slideLayout' )
        
#TECHDEBT: Need a way to make these two Part files optional
        # self.handoutmasterfile = SinglePartFile(templatedir, 'handoutMaster')
        # self.notesmasterfile   = SinglePartFile(templatedir, 'notesMaster')
        self.presentationfile  = SinglePartFile(templatedir, 'presentation')
        self.prespropsfile     = SinglePartFile(templatedir, 'presProps')
        self.tablestylesfile   = SinglePartFile(templatedir, 'tableStyles')
        self.viewpropsfile     = SinglePartFile(templatedir, 'viewProps')
        
    @property
    def relsfiles(self):
        return [partfile.relsfile for partfile in self.partfiles if partfile.relsfile]
    
    @property
    def partfiles(self):
        partfiles = []
        partfiles.extend(self.imagefiles)
        partfiles.extend(self.themefiles)
        partfiles.extend(self.slidemasterfiles)
        partfiles.extend(self.slidelayoutfiles)
        partfiles.append(self.prespropsfile)
        return partfiles
    


# ============================================================================
# PackageItemFile
# ============================================================================

class PackageItemFile(object):
    pass
    


# ============================================================================
# PackageItemFileCollection
# ============================================================================

class PackageItemFileCollection(list):
    
    def __init__(self):
        list.__init__(self)
    


# ============================================================================
# PartFile
# ============================================================================

class PartFile(PackageItemFile):
    
    def __init__(self, partfiles, path):
        PackageItemFile.__init__(self)
        self.parttype  = partfiles.parttype
        self.partfiles = partfiles
        self.parent    = partfiles
        self.path      = path
        self.dirpath   = os.path.split(path)[0]
        self.relsfile  = self.loadrelsfile()
    
    def loadrelsfile(self):
        if self.parttype.has_rel_item == 'never':
            return None
        dirpath, filename = os.path.split(self.path)
        rootname, ext     = os.path.splitext(filename)
        relsdirpath       = os.path.join(dirpath, '_rels')
        relsfilename      = '%s.rels' % filename
        relsfilepath      = os.path.join(relsdirpath, relsfilename)
        if not os.path.isfile(relsfilepath):
            if self.parttype.has_rel_item == 'optional':
                return None
            else:
                raise CorruptedTemplateError('''No relationships file '%s' found for part '%s' in template %s''' % (relsfilename, filename, self.partfiles.templatedir))
        return PartRelsFile(self, relsfilepath)
    


# ============================================================================
# SinglePartFile
# ============================================================================
#REFACTOR: Either rename this to SingletonPartFile or probably better just
#          make it PartFile and create a subclass CollectionPartFile.
#          Actually, should probably just refactor loading functionality into
#          Parts and get rid of separate PartFile class hierarchy.

class SinglePartFile(PartFile):
    
    # NOTE: Although PresPropsPartFile inherits from PartFile, it doesn't call
    #       PartFile.__init__(). Probably their roles should be reversed, like
    #       PartFile becomes GroupPartFile, this becomes PartFile, and
    #       GroupPartFile inherits from PartFile. Then PartFiles should become
    #       PartFileGroup.
    def __init__(self, templatedir, parttypekey):
        PackageItemFile.__init__(self)
        self.parttype = PartTypes.parttype(parttypekey)
        self.filename = "%s.%s" % (self.parttype.rootname, self.parttype.file_ext)
        self.path     = os.path.join(templatedir, self.parttype.pkgreldir, self.filename)
        # self.dirpath     = os.path.split(path)[0]
        self.relsfile    = self.loadrelsfile()
    
#HACKALERT: This doesn't work properly, need separate part file for handoutMaster
        self.images = []
        self.theme  = []
        self.relateditems = []
        
    # # This is now inherited from PartFile, but later the roles will be
    # # reversed. Either way, only one copy needed.
    # def __loadrelsfile(self):
    #     if self.parttype.has_rel_item == 'never':
    #         return None
    #     dirpath, filename = os.path.split(self.path)
    #     rootname, ext     = os.path.splitext(filename)
    #     relsdirpath       = os.path.join(dirpath, '_rels')
    #     relsfilename      = '%s.rels' % filename
    #     relsfilepath      = os.path.join(relsdirpath, relsfilename)
    #     if not os.path.isfile(relsfilepath):
    #         if self.parttype.has_rel_item == 'optional':
    #             return None
    #         else:
    #             raise CorruptedTemplateError('''No relationships file '%s' found for part '%s' in template %s''' % (relsfilename, filename, self.partfiles.templatedir))
    #     return PartRelsFile(self, relsfilepath)
    


# ============================================================================
# PartFiles
# ============================================================================
#TECHDEBT: Refactor 'ext' throughout to include the leading period. That's
#          consistent with ext returned from os.path.splitext(), and also
#          allows for empty extensions (as in 'python'), which I'm sure is
#          why they designed splitext that way :).

class PartFiles(PackageItemFileCollection):
    
    def __init__(self, templatedir=None, parttypekey=None):
        PackageItemFileCollection.__init__(self)
        self.templatedir = templatedir
        self.parttypekey = parttypekey
        self.parttype    = PartTypes.parttype(parttypekey) if parttypekey else None
        if templatedir and parttypekey:
            self.load(templatedir, parttypekey)
    
    def load(self, templatedir, parttypekey):
        self.templatedir = templatedir
        self.parttypekey = parttypekey
        self.parttype    = PartTypes.parttype(parttypekey)
        del self[0:]
        for partfilepath in self.__partfilepaths:
            self.append(PartFile(self, partfilepath))
    
    # Return list of fully qualified part file paths, sorted in numerical
    # order (not lexicographic order), e.g. [slideLayout1, slideLayout2 ...]
    # rather than [slideLayout1, slideLayout10 ...]
    # -----------------------------------------------------------------------
    @property
    def __partfilepaths(self):
        pkgreldir       = self.parttype.pkgreldir
        filenameroot    = self.parttype.rootname
        ext             = self.parttype.file_ext if self.parttype.file_ext else '.+'  # empty extension means extension can vary, like .jpg, .png, .wmf
        searchdirpath   = os.path.join(self.templatedir, pkgreldir)
        filenamepattern = r'%s([1-9][0-9]*)?\.(%s)' % (filenameroot, ext)
        regexp          = re.compile(filenamepattern)
        partfilepaths   = {}
        if not os.path.isdir(searchdirpath):
            return []
        for name in os.listdir(searchdirpath):
            fqname = os.path.join(searchdirpath, name)
            if not os.path.isfile(fqname):
                continue
            match = regexp.match(name)
            if not match:
#TECHDEBT: Should probably remove this after this class is debugged because it's not really an error to find an unexpected file, it just gets skipped.
                raise CorruptedTemplateError('''Unexpected file '%s' found in template''' % os.path.join(self.parttype.location, name))
                continue
            filenamenumber = match.group(1)
            sortkey = int(filenamenumber) if filenamenumber else 0
            partfilepaths[sortkey] = fqname
        # return file paths sorted in numerical order (not lexicographic order)
        return [partfilepaths[key] for key in sorted(partfilepaths.keys())]
    


# ============================================================================
# PartRelsFiles
# ============================================================================

class PartRelsFile(PackageItemFile):
    
    def __init__(self, partfile, path):
        PackageItemFile.__init__(self)
        self.partfile = partfile
        self.path     = path
        self.filename = os.path.split(self.path)[1]
    


