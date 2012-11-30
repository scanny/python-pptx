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

import os
import zipfile

from StringIO   import StringIO
from lxml       import etree
from lxml.etree import ElementTree, Element, SubElement

from pptx           import spec
from pptx           import util
from pptx.packaging import Package, TemplatePackage
from pptx.spec      import qname


# ============================================================================
# Class catalog
# ============================================================================

# API Classes
# ----------------------------------------------------------------------------
# Presentation
# Template

# Base Classes
# ----------------------------------------------------------------------------
# Shape
#REFACTOR: See if there's an opportunity for Part, perhaps Master, also
#          Collection.

# Presentation Parts
# ----------------------------------------------------------------------------
# Slide
# Slides

# Slide Elements
# ----------------------------------------------------------------------------
# Shapes
#REFACTOR: This code needs to find a home in another slide element
# # PlaceholderText

# Template Part Collections
# ----------------------------------------------------------------------------
# TemplatePartCollection
# Images
# SlideLayouts
# SlideMasters
# Themes

# Template Parts
# ----------------------------------------------------------------------------
# TemplatePart
# HandoutMaster
# Image
# NotesMaster
# PresProps
# SlideLayout
# SlideMaster
# TableStyles
# Theme
# ViewProps

#TODO: To be implemented
# ----------------------------------------------------------------------------
# HandoutMaster
# NotesMaster
# SlideShow


# ============================================================================
# Presentation
# ============================================================================
# Top level class in object model, represents the contents of the /ppt
# directory of a .pptx file.
#     I'm thinking there's a higher-level construct, perhaps an OfficeDocument
# class that holds things like the document properties and thumbnail image.
# ============================================================================
# CAUTION: Currently __applytemplate assumes all the presentation
# collections are empty, so it should only be called from the __init__()
# method. It might make sense later to add it to the API and make it smart
# enough to be applied to a presentation that already has some content in it,
# but that's for a future effort.
# ============================================================================

class Presentation(object):
    
    def __init__(self, templatedir=None):
        self.templatedir  = self.__normalizedtemplatedir(templatedir)
        self._element     = None
        # initialize collections
        self.images       = Images       ()
        self.themes       = Themes       ()
        self.slidemasters = SlideMasters ()
        self.slidelayouts = SlideLayouts ()
        self.slides       = Slides       (self)
        # apply template
        template = Template(self.templatedir)
        self.__applytemplate(template)
    
    def __applytemplate(self, template):
        # add parts
#REFACTOR: Seems like the template's part collections should be added directly
#          rather than reconstructing them from scratch. Maybe even do it
#          in __init__() and then initializing the presentation-specific
#          collections like slides. E.g.:
#          self.images = template.images if template.images else ImageCollection()
        for image       in template.images       : self.images       .additem(image       .path)
        for theme       in template.themes       : self.themes       .additem(theme       .path)
        for slidemaster in template.slidemasters : self.slidemasters .additem(slidemaster .path)
        for slidelayout in template.slidelayouts : self.slidelayouts .additem(slidelayout .path)
        # set singleton parts
        self.element       = template.element      # this part is different because it's the presentation root item rather than a part that belongs to the presentation
        self.presprops     = template.presprops
        self.tablestyles   = template.tablestyles
        self.viewprops     = template.viewprops
#TECHDEBT: Need to make this work when either of these two files don't exist in the template
        # self.handoutmaster = template.handoutmaster
        # self.notesmaster   = template.viewprops
        
        # add relationships
        for relationship in template.relationships:
            self.__applyrelationship(relationship)
    
    def __applyrelationship(self, relationship):
        # print 'Relationship %s  ==>  %s' % (relationship.fromtype, relationship.totype)
        
        collections = {'image': self.images, 'slideLayout': self.slidelayouts, 'slideMaster': self.slidemasters, 'theme': self.themes}
        fromtype    = relationship.fromtype
        frompath    = relationship.frompath
        fromitem    = collections[fromtype].getitem(path=frompath)
        if not fromitem: raise Exception("""Lookup of %s failed on relationship %s ==> %s""" % (fromtype, fromtype, totype))   # remove after debugging
        totype      = relationship.totype
        topath      = relationship.topath
        toitem      = collections[totype].getitem(path=topath)
        if not toitem: raise Exception("""Lookup of %s failed on relationship %s ==> %s""" % (totype, fromtype, totype))       # remove after debugging
        
        if fromtype == 'slideMaster':
            if totype == 'slideLayout':
                fromitem.slidelayouts.append(toitem) # slideLayout provides backreference
            elif totype == 'theme':
                fromitem.theme = toitem              # no backreference required
            elif totype == 'image':
                fromitem.images.append(toitem)       # add image reference to slideMaster
                toitem.slidemasters.append(fromitem) # add backreference
            else:
                raise NotImplementedError("""Relationship from %s to %s not implemented.""" % (fromtype, totype))
        elif fromtype == 'slideLayout':
            if totype == 'slideMaster':
                fromitem.slidemaster = toitem        # slideMaster provides backreference
            elif totype == 'image':
                fromitem.images.append(toitem)       # add image reference to slideLayout
                toitem.slidelayouts.append(fromitem) # and add backreference from image
            # elif totype == 'theme':
            #     fromitem.theme = toitem              # no backreference required
            else:
                raise NotImplementedError("""Relationship from %s to %s not implemented.""" % (fromtype, totype))
        elif fromtype == 'theme':
            if totype == 'image':
                raise NotImplementedError("""Relationship from %s to %s not implemented.""" % (fromtype, totype))
                # lookup the image in question
                # add a reference to the image to the theme
                # add a reference to the theme to the image
            else:
                raise NotImplementedError("""Relationship from %s to %s not implemented.""" % (fromtype, totype))
        else:
            raise NotImplementedError("""Relationship from %s to %s not implemented.""" % (fromtype, totype))
            
        # print "collections[fromtype] is class '%s'" % collections[fromtype].__class__.__name__
        # print "fromitem is class '%s'" % fromitem.__class__.__name__
        # print 'Presentation contains %d slidemaster(s).' % len(self.slidemasters)
        # print
    
    def __normalizedtemplatedir(self, templatedir):
        if not templatedir:  # if templatedir is not provided, the default template is used
            return os.path.join(os.path.dirname(__file__),'pptx_template')
        assert os.path.isdir(templatedir)
        return os.path.abspath(templatedir)
    
    @property
    def element(self):
        rId_iter     = util.intsequence(start=1)
        presentation = self._element  # NOTE: this is presentation element, not presentation instance
        
        # rewrite the sldMasterIdLst subelement ______________________________
        sldMasterId_iter = util.intsequence(start=2147483660)   # This starting number is a bit of a hack based on sldMasterId values found in files saved by PowerPoint
        
        # create the new sldMasterIdLst element
        sldMasterIdLst = Element(qname('p','sldMasterIdLst'))
        for slidemaster in self.slidemasters:
            id  = str(sldMasterId_iter.next())
            rId = 'rId%d' % rId_iter.next()
            sldMasterId = SubElement(sldMasterIdLst, qname('p','sldMasterId'))
            sldMasterId.set('id', id)
            sldMasterId.set(qname('r','id'), rId)
            sldMasterIdLst.append(sldMasterId)
        
        # get the old sldMasterIdLst element
        old_sldMasterIdLst = presentation.find(qname('p','sldMasterIdLst'))
        # put new sldMasterIdLst element in place
        if old_sldMasterIdLst is None:
            presentation.insert(0, sldMasterIdLst)
        else:
            presentation.replace(old_sldMasterIdLst, sldMasterIdLst)
        
        # rewrite the sldIdLst subelement ____________________________________
        sldId_iter = util.intsequence(start=256)   # This starting number is also based on observed values in PowerPoint files, no idea where it comes from.
        
        # create the new sldIdLst element
        sldIdLst = Element(qname('p','sldIdLst'))
        for slide in self.slides:
            id = str(sldId_iter.next())
            rId = 'rId%d' % rId_iter.next()
            sldId = SubElement(sldIdLst, qname('p','sldId'))
            sldId.set('id', id)
            sldId.set(qname('r','id'), rId)
            sldIdLst.append(sldId)
        
        # get the old sldIdLst element
        old_sldIdLst = presentation.find(qname('p','sldIdLst'))
        # put new sldIdLst element in place
        if old_sldIdLst is None:
            #TECHDEBT: Won't need these next three lines once notesMasters and handoutMasters support is added
            idx = 1
            if presentation.find(qname('p','notesMasterIdLst'  )) : idx += 1
            if presentation.find(qname('p','handoutMasterIdLst')) : idx += 1
            presentation.insert(idx, sldIdLst)
        else:
            presentation.replace(old_sldIdLst, sldIdLst)
        
        return self._element
        
    
    @element.setter
    def element(self, element):
        self._element = element
    
    def save(self, filename):
        Package(self).save(filename)
    


# ============================================================================
# Template
# ============================================================================
# Assemble a presentation that contains only template parts (essentially a
# presentation without any slides or slide sub-parts like notesSlides or
# comments) from the unzipped presentation package in the specified directory.
# ============================================================================
# DEVNOTE: This is a bit overkill just for applying a template to a new
# presentation, but I'm thinking maybe later features will work on templates
# as a first-class citizen, so happy to leave this implemented this way for
# the sake of future enhancement opportunities.
# ============================================================================

class Template(object):
    
    def __init__(self, templatedir):
        # initialize collections
        self.images       = Images       ()
        self.themes       = Themes       ()
        self.slidemasters = SlideMasters ()
        self.slidelayouts = SlideLayouts ()
        # load template files from templatedir
        templatepackage = TemplatePackage(templatedir)
        templatefiles = templatepackage.templatefiles
        for imagefile       in templatefiles.imagefiles       : self.images       .additem(imagefile       .path)
        for themefile       in templatefiles.themefiles       : self.themes       .additem(themefile       .path)
        for slidemasterfile in templatefiles.slidemasterfiles : self.slidemasters .additem(slidemasterfile .path)
        for slidelayoutfile in templatefiles.slidelayoutfiles : self.slidelayouts .additem(slidelayoutfile .path)
        
        self.presprops     = PresProps     (templatefiles.prespropsfile.path)
        self.tablestyles   = TableStyles   (templatefiles.tablestylesfile.path)
        self.viewprops     = ViewProps     (templatefiles.viewpropsfile.path)
#TECHDEBT: Need to make these two files optional
        # self.handoutmaster = HandoutMaster (templatefiles.handoutmasterfile.path)
        # self.notesmaster   = NotesMaster   (templatefiles.viewpropsfile.path)
        
        # load presentation.xml into element property
        parser = etree.XMLParser(remove_blank_text=True)  # need to remove indentation whitespace so pretty_print will work later
        self.element = etree.parse(templatefiles.presentationfile.path, parser).getroot()
        
        # load package relationships (file-path oriented)
        self.relationships = templatepackage.relationships
    


# ============================================================================
# Shape
# ============================================================================
# Base class for the various types of shape (not including pic, which is a
# distinct element type).  This is actually an abstract class. Adding a Shape
# directly to a slide will likely cause a load error in PowerPoint. Use one
# of the fully fleshed-out subclasses such as PlaceholderText instead or
# modify the element directly to add the required bits for a complete shape.
# ============================================================================

#TECHDEBT: need to account for hierarchy of GroupShape, cSld/spTree is type
#          GroupShape and GroupShape can contain other GroupShape elements in
#          a hierarchy

class Shape(object):
    
    def __init__(self, name):
        self.name    = name
        self.element = self.minimalelement
    
    @property
    def type(self):
        raise NotImplementedError('''The *type* property needs to be implemented for all subclasses of Shape.''')
    
    @property
    def minimalelement(self):
        sp      = Element(qname('p', 'sp'))
        nvSpPr  = SubElement(sp     , qname('p', 'nvSpPr'  ) )
        cNvPr   = SubElement(nvSpPr , qname('p', 'cNvPr'   ) , id='None', name=self.name)
        cNvSpPr = SubElement(nvSpPr , qname('p', 'cNvSpPr' ) )
        spPr    = SubElement(sp     , qname('p', 'spPr'    ) )
        return sp
    


# ============================================================================
# Slide
# ============================================================================
# Superclass to take care of all base slide functionality.
# ============================================================================

class Slide(object):
    
    def __init__(self, id, slides, slidelayout, name):
        self.id          = id
        self.key         = id
        self.slides      = slides
        self.parent      = slides
        self.slidelayout = slidelayout
        self.name        = name
        self.nsmap       = spec.nsmap_subset(['a', 'p', 'r'])  # keep this here so remaining init can use it
        self.__id_iter   = util.intsequence()                  # this one too ...
        self.element     = self.minimalelement
        self.shapes      = Shapes(self)
    
    @property
    def idx(self):
        return self.slides.index(self)
    
    @property
    def images(self):
        return [shape.image for shape in self.shapes if shape.type == 'pic']
    
    #TECHDEBT: Not sure about this one, just trying it to get a first package
    #          build working
    def setxml(self, xml):
        self.element = etree.fromstring(xml)
    
    @property
    def _nextid(self):
        return self.__id_iter.next()
    
    # def _append_shape(self, shape):
    #     # assign the new shape an id that's unique within this Slide
    #     sp_cNvPr = shape.element.xpath('/p:sp/p:nvSpPr/p:cNvPr', namespaces=nsmap)[0]
    #     sp_cNvPr.attrib['id'] = str(self._nextid)
    #     # append shape to spTree element of Slide
    #     sld_spTree = self.element.xpath('/p:sld/p:cSld/p:spTree', namespaces=nsmap)[0]
    #     sld_spTree.append(shape.element)
    
    @property
    def minimalelement(self):
        sld        = Element(qname('p', 'sld'), nsmap=self.nsmap)
        cSld       = SubElement(sld       , qname('p', 'cSld'       ) )
        spTree     = SubElement(cSld      , qname('p', 'spTree'     ) )
        nvGrpSpPr  = SubElement(spTree    , qname('p', 'nvGrpSpPr'  ) )
        cNvPr      = SubElement(nvGrpSpPr , qname('p', 'cNvPr'      ) , id=str(self._nextid), name=self.name)
        cNvGrpSpPr = SubElement(nvGrpSpPr , qname('p', 'cNvGrpSpPr' ) )
        nvPr       = SubElement(nvGrpSpPr , qname('p', 'nvPr'       ) )
        grpSpPr    = SubElement(spTree    , qname('p', 'grpSpPr'    ) )
        return sld
    
    @property
    def xmlstring(self):
        return etree.tostring(self.element, pretty_print=True)
    


# ============================================================================
# Slides
# ============================================================================

class Slides(list):
    
    def __init__(self, presentation):
        list.__init__(self)
        self.presentation = presentation
        self.__id_iter    = util.intsequence()
        self.__dict       = {}
    
    def addslide(self, slidelayoutidx=0, name=None, xml=None):
        id          = self.__id_iter.next()
        slidelayout = self.presentation.slidelayouts.getitem(idx=slidelayoutidx)
# ----------------------------------------------------------------------------
#TECHDEBT: This method of assigning the name allows more than one slide to
#          potentially have the same name. Need something a bit more
#          sophisticated if it's important to avoid that scenario. Note that
#          it's perhaps somewhat unlikely for that to happen when the
#          presentation is being built by software rather than an end-user
#          who might rearrange and delete slides.
# ----------------------------------------------------------------------------
#REFACTOR: Refactor to use Slides.__nextname and let it use the next unused
#          name. or copy PowerPoint behavior, whatever that is when you delete
#          a slide then add a new one.
        name        = name if name else 'Slide %d' % len(self.presentation.slides)+1
        slide       = Slide(id, self, slidelayout, name)
# ----------------------------------------------------------------------------
#TECHDEBT: Not sure exactly how to implement a setxml() method of determining
#          a slide's content. Would cause a round-tripping challenge unless
#          the XML was parsed and all the properties of the slide set from
#          its contents. Or perhaps it could set a flag that the slide's
#          contents were static and no properties could be set. Or perhaps
#          the properties we knew how to change would be updated without
#          generating the slide from scratch. So the provided XML would be
#          used as "starting" XML instead of minimalxml(). Opens question
#          of whether XML is updated with each property set call or only
#          generated when the slide XML is rendered.
# ----------------------------------------------------------------------------
        if xml:
            slide.setxml(xml)
        self.__dict[slide.key] = slide
        self.append(slide)
        return slide
    


# ============================================================================
# Shapes
# ============================================================================
# Inherits from list so additional methods and behaviors can be added to
# shapes member of Slide.
# ============================================================================

class Shapes(list):
    
    def __init__(self, parent):
        list.__init__(self)
        self.parent = parent
    
    def append(self, shape):
        if not isinstance(shape, Shape):
            raise TypeError("Slide.shapes.append() called with incorrect type, (expected Shape, got %s)" % shape.__class__.__name__)
        list.append(self, shape)
        self.slide._append_shape(shape)
    


# ============================================================================
# TemplatePartCollection
# ============================================================================

class TemplatePartCollection(list):
    
    def __init__(self):
        list.__init__(self)
        self.__dict = {}
    
    def additem(self, path):
        key = path
        if key in self.__dict:
            return self.__dict[key]
        part = self.memberclass(self, path)
        self.__dict[key] = part
        self.append(part)
        return part
    
    # Can implement a __contains__(self, value) method if anyone needs to test
    # whether something is already in the collection.
    def getitem(self, path=None, idx=None):
        if path:
            key = path
            if key in self.__dict:
                return self.__dict[key]
            raise KeyError("""getitem() lookup in %s failed with key '%s'""" % (self.__class__.__name__, key))
        if idx:
            return self[idx]
        raise Exception('''%s.getitem() called without a lookup key (either path or idx)''' % self.__class__.__name__)
    


# ============================================================================
# Images
# ============================================================================
# Collection of all the images used in this presentation. Each image is
# stored only once, regardless of how often it is used.
# ============================================================================

class Images(TemplatePartCollection):
    
    def __init__(self):
        TemplatePartCollection.__init__(self)
        self.memberclass = Image
    


# ============================================================================
# SlideLayouts
# ============================================================================

class SlideLayouts(TemplatePartCollection):
    
    def __init__(self):
        TemplatePartCollection.__init__(self)
        self.memberclass = SlideLayout
    


# ============================================================================
# SlideMasters
# ============================================================================

class SlideMasters(TemplatePartCollection):
    
    def __init__(self):
        TemplatePartCollection.__init__(self)
        self.memberclass = SlideMaster
    


# ============================================================================
# Themes
# ============================================================================

class Themes(TemplatePartCollection):
    
    def __init__(self):
        TemplatePartCollection.__init__(self)
        self.memberclass = Theme
    


# ============================================================================
# TemplatePart
# ============================================================================
#REFACTOR: Needs some cleanup here to account for differences between single
#          parts and those that are part of a collection (group parts?). The
#          inheritance of properties is too distributed between super class
#          and subclass, and it's weird that they have to keep track of their
#          own cardinality. Using the modifiers "scalar" vs. "vector" might be
#          something to use in elaborating the class hierarchy by one level
#          e.g. ScalarTemplatePart and VectorTemplatePart. Also consider
#          single vs group, scalar vs tuple, scalar vs list, scalar vs array,
#          (plain old) part vs collectionPart. That last one is interesting
#          because collection parts need everything a regular part has and
#          just need to add a bit more. Also 'collection' is what they're a
#          part of. SOLD!

class TemplatePart(object):
    
    def __init__(self, path):
        self.path        = path
        self.key         = path
        self.cardinality = 'multiple'
    
    @property
    def idx(self):
        if self.cardinality == 'multiple':
            return self.parent.index(self)
        else:
            raise NotImplementedError('''TemplatePart.idx is not defined for parts with cardinality!='multiple'. Called on class '%s'.''' % self.__class__.__name__)
    


# ============================================================================
# Image
# ============================================================================

class Image(TemplatePart):
    
    def __init__(self, parent, path):
        TemplatePart.__init__(self, path)
        self.parent       = parent
        self.slidemasters = []
        self.slidelayouts = []
        self.slides       = []
        self.themes       = []
    


# ============================================================================
# HandoutMaster
# ============================================================================

class HandoutMaster(TemplatePart):
    
    def __init__(self, path):
        TemplatePart.__init__(self, path)
        self.cardinality = 'single'
    


# ============================================================================
# NotesMaster
# ============================================================================

class NotesMaster(TemplatePart):
    
    def __init__(self, path):
        TemplatePart.__init__(self, path)
        self.cardinality = 'single'
    


# ============================================================================
# PresProps
# ============================================================================

class PresProps(TemplatePart):
    
    def __init__(self, path):
        TemplatePart.__init__(self, path)
        self.cardinality = 'single'
    


# ============================================================================
# SlideLayout
# ============================================================================

class SlideLayout(TemplatePart):
    
    def __init__(self, parent, path):
        TemplatePart.__init__(self, path)
        self.parent      = parent
        self.slidemaster = None
        self.images      = []
    


# ============================================================================
# SlideMaster
# ============================================================================
# TECHNOTE: In the Microsoft API, Master is a general type that all of
# SlideMaster, SlideLayout (CustomLayout), HandoutMaster, and NotesMaster
# inherit from. So might look into why that is and consider refactoring the
# various masters a bit later.
# ============================================================================

class SlideMaster(TemplatePart):
    
    def __init__(self, parent, path):
        TemplatePart.__init__(self, path)
        self.parent       = parent
        self.element      = etree.parse(path).getroot()
        self.theme        = None
        self.slidelayouts = []
        self.images       = []
    


# ============================================================================
# TableStyles
# ============================================================================

class TableStyles(TemplatePart):
    
    def __init__(self, path):
        TemplatePart.__init__(self, path)
        self.cardinality = 'single'
    


# ============================================================================
# Theme
# ============================================================================

class Theme(TemplatePart):
    
    def __init__(self, parent, path):
        TemplatePart.__init__(self, path)
        self.parent = parent
        self.images = []
    


# ============================================================================
# ViewProps
# ============================================================================

class ViewProps(TemplatePart):
    
    def __init__(self, path):
        TemplatePart.__init__(self, path)
        self.cardinality = 'single'
    


# # ============================================================================
# # PlaceholderText
# # ============================================================================
# # Shape that holds text to be displayed in a placeholder specified by the
# # slide layout.
# #     The placeholder_type provided must be one of ST_PlaceholderType defined
# # in the PresentationML schema.
# #     The string supplied as the text parameter may contain carriage returns
# # ('\r') and or line feeds ('\n'). Carriage returns will cause a new run to
# # be created with a break (<a:br/>) intervening between it and the prior run.
# # Line feeds will cause a new paragraph to be created to contain the text
# # that follows the line feed.
# #     The index parameter is important for matching up the text with the
# # right placeholder. I don't fully understand the logic used, but in one
# # pptx file I've seen the <p:ph> idx attribute matched with that of the
# # proper placeholder. The default is 0 and that seems to match the title
# # placeholder.
# # ============================================================================
#
# class PlaceholderText(Shape):
#
#     def __init__(self, name, placeholder_type, text, placeholder_index=None):
#         if placeholder_type not in placeholder_types:
#             raise TypeError("placeholder_type must be one of %s, got '%s'." % (placeholder_types, placeholder_type))
#         Shape.__init__(self, name)
#         self.placeholder_type  = placeholder_type
#         self.placeholder_index = placeholder_index
#         self.text = text
#
#         # locate nodes we need to have handy (these are all required elements so we're assured they're present)
#         sp      = self.element
#         nvSpPr  = sp.find(qname('p', 'nvSpPr'))
#         spPr    = sp.find(qname('p', 'spPr'))
#         cNvSpPr = nvSpPr.find(qname('p', 'cNvSpPr'))
#
#         # add shape lock to prevent grouping the placeholder element with others
#         node = cNvSpPr.find(qname('a', 'spLocks'))
#         spLocks = node if node else SubElement(cNvSpPr, qname('a', 'spLocks'))
#         spLocks.attrib['noGrp'] = 'true'
#
#         # add placeholder type so text can find its home on slide layout
#         node = nvSpPr.find(qname('p', 'nvPr'))
#         nvPr = node if node else SubElement(nvSpPr, qname('p', 'nvPr'))
#         node = nvPr.find(qname('p', 'ph'))
#         ph   = node if node else SubElement(nvPr, qname('p', 'ph'))
#         ph.attrib['type'] = placeholder_type
#         if placeholder_index:
#             ph.attrib['idx'] = str(placeholder_index)
#
#         # add txBody to hold text
#         txBody = sp.find(qname('p', 'txBody'))
#         if not node:
#             txBody = Element(qname('p', 'txBody'))
#             sp.insert(sp.index(spPr)+1, txBody)
#         bodyPr = txBody.find(qname('a', 'bodyPr'))
#         if not bodyPr:
#             bodyPr = Element(qname('a', 'bodyPr'))
#             txBody.insert(0, bodyPr)
#
#         # get rid of any text that might be there already
#         del txBody[1:]
#
#         # insert the provided text
#         lines = text.split('\n')
#         for line in lines:
#             p = SubElement(txBody, qname('a', 'p'))
#             runs = line.split('\r')
#             for run in runs:
#                 br = SubElement(p, qname('a', 'br')) if runs.index(run) > 0 else None
#                 r = SubElement(p, qname('a', 'r'))
#                 t = SubElement(r, qname('a', 't'))
#                 t.text = run
#
#
#     @property
#     def minimalelement(self):
#         sp      = Element(qname('p', 'sp'))
#         nvSpPr  = SubElement(sp     , qname('p', 'nvSpPr'  ) )
#         cNvPr   = SubElement(nvSpPr , qname('p', 'cNvPr'   ) , id='None', name=self.name)
#         cNvSpPr = SubElement(nvSpPr , qname('p', 'cNvSpPr' ) )
#         spPr    = SubElement(sp     , qname('p', 'spPr'    ) )
#         return sp


