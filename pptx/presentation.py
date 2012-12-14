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

from lxml       import etree
from lxml.etree import ElementTree, Element, SubElement

from pptx           import util
from pptx.packaging import Package
# from pptx.spec      import PartType, nsmap_subset, qname
from pptx.spec      import nsmap_subset, qname

# ============================================================================
# API Classes
# ============================================================================

class Presentation(object):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file. I'm thinking there's a higher-level construct,
    perhaps an OfficeDocument class that holds things like the document
    properties and thumbnail image.
    
    REFACTOR: Inherit from Part.
    
    """
    def __init__(self, templatedir=None):
        self.templatedir  = self.__normalizedtemplatedir(templatedir)
        self.parttypekey  = 'presentation'
        self._element     = None
        self.path         = None  # need this default value until inherit from Part
        # initialize collections
        self.images       = ImageCollection       (self)
        self.themes       = ThemeCollection       (self)
        self.slidemasters = SlideMasterCollection (self)
        self.slidelayouts = SlideLayoutCollection (self)
        self.slides       = SlideCollection       (self)
        # load template parts
        self.__loadtemplate(templatedir)
    
    def __loadtemplate(self, templatedir):
        """
        Load the template parts of the package unzipped in *templatedir*.
        
        """
        # tp = TemplatePackage(templatedir)
        for part in tp.imageparts:
            self.images.additem(part)
        for part in tp.slidelayoutparts:
            self.slidelayouts .additem(part)
        for part in tp.themeparts:
            self.themes.additem(part)
        for part in tp.slidemasterparts:
            self.slidemasters.additem(part)
        self.presprops = PresProps(self).load(tp.prespropspart)
        self.tablestyles = TableStyles(self).load(tp.tablestylespart)
        self.viewprops = ViewProps(self).load(tp.viewpropspart)
        # load presentation-level element and relationships
        self.element = tp.presentationpart.element
        self.relationships = tp.presentationpart.relationships
    
    def __normalizedtemplatedir(self, templatedir):
        """
        Return an absolute path to the directory containing the specified
        template. If *templatedir* is ``None``, return an absolute path to the
        default template.
        """
        if not templatedir:  # if templatedir is not provided, the default template is used
            return os.path.join(os.path.dirname(__file__),'pptx_template')
        assert os.path.isdir(templatedir)
        return os.path.abspath(templatedir)
    
    @property
    def element(self):
        """
        Return XML for the presentation part as an ElementTree element.
        """
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
        """
        Save this presentation as a .pptx file. If *filename* does not end in
        '.pptx', that extension will be appended.
        
        """
        Package().save(self, filename)
    


# ============================================================================
# Part Collections
# ============================================================================

class PartCollection(list):
    """
    Base class for collections of presentation parts.
    
    Subclasses of :class:`PartCollection` must implement :attr:`parttype` to
    provide access to metadata about the member parts of the collection.
    
    """
    def __init__(self, presentation):
        super(PartCollection, self).__init__()
        self.presentation = presentation
        self._dict = {}
    
    def additem(self, partitem):
        key = partitem.key
        if key in self._dict:
            return self._dict[key]
        part = self.memberclass(self).load(partitem)
        self._dict[key] = part
        self.append(part)
        return part
    
    # Can implement a __contains__(self, value) method if anyone needs to test
    # whether something is already in the collection.
    def getitem(self, key=None, idx=None):
        if key:
            if key in self._dict:
                return self._dict[key]
            raise KeyError("""getitem() lookup in %s failed with key '%s'""" % (self.__class__.__name__, key))
        if idx:
            return self[idx]
        raise Exception('''%s.getitem() called without a lookup key (either path or idx)''' % self.__class__.__name__)
    

class ImageCollection(PartCollection):
    """
    Collection of all the images used in this presentation. Each image is
    stored only once, regardless of how often it is used.
    
    """
    def __init__(self, presentation):
        super(ImageCollection, self).__init__(presentation)
        self.parttype    = PartType.lookup('image')
        self.memberclass = Image
    

class SlideCollection(PartCollection):
    """
    Collection of slides in a presentation.
    
    """
    def __init__(self, presentation):
        super(SlideCollection, self).__init__(presentation)
        self.parttype  = PartType.lookup('slide')
        self.__id_iter = util.intsequence()
        self._dict     = {}
    
    def addslide(self, slidelayoutidx=0, name=None, xml=None):
        """
        Add a new slide to the presentation and return the new slide.
        
        **TODO:** Add an optional *idx=None* parameter that allows the
        position of the new slide within the collection to be specified.
        Default behavior will be to append the new slide to the end of the
        presentation's slide collection.
        
        **REFACTOR:** Add *element* property with a corresponding setter and
        allow element of slide to be set using API. Then set that element if
        XML is passed in to the constructor, like when a slide is being formed
        from a template. Note that any images referred to in the XML would
        need to get synced up somehow.
        """
        id          = self.__id_iter.next()
#TECHDEBT: Quick hack on next line to get this to work while I'm debugging elsewhere
        print "SlideCollection.presentation ==> %s" % self.presentation
        slidelayout = self.presentation.slidelayouts[slidelayoutidx]
        name        = name if name else 'Slide %d' % len(self.presentation.slides)+1  # Note: more than one slide could end up with same name
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
        self._dict[slide.key] = slide
        self.append(slide)
        return slide
    

class SlideLayoutCollection(PartCollection):
    
    def __init__(self, presentation):
        super(SlideLayoutCollection, self).__init__(presentation)
        self.parttype    = PartType.lookup('slideLayout')
        self.memberclass = SlideLayout
    

class SlideMasterCollection(PartCollection):
    
    def __init__(self, presentation):
        super(SlideMasterCollection, self).__init__(presentation)
        self.parttype    = PartType.lookup('slideMaster')
        self.memberclass = SlideMaster
    

class ThemeCollection(PartCollection):
    
    def __init__(self, presentation):
        super(ThemeCollection, self).__init__(presentation)
        self.parttype    = PartType.lookup('theme')
        self.memberclass = Theme
    


# ============================================================================
# Parts
# ============================================================================

class Part(object):
    """
    Base class that presentation parts inherit from. Its primary purpose is to
    contain code common to those parts.
    
    REFACTOR: Check construction signature, doesn't seem quite right.
    
    REFACTOR: Load and carry parttype for the type, not just the parttypekey.
    
    .. attribute:: parent
       
       PartCollection instance containing this part.
    
    .. attribute:: parttype
       
       :class:`pptx.spec.PartType` instance containing metadata about the
       part type contained in this collection.
    
    """
    def __init__(self, presentation, parttype):
        self.presentation = presentation
        self.parttype     = parttype
        self._element     = None
    
    def load(self, part):
        """
        Load this part from an object that implements the attributes *key*,
        *element*, *path*, and *relationships*. *key* is an identifier for
        this part that is unique within its presentation context. *element* is
        an ElementTree element containing the XML for this part. If
        ``parttype.format`` for this part is 'binary', *element* will be None
        and *path* should be used to store a reference to the file containing
        the part. If ``parttype.format`` for this part is 'xml', *path* may be
        None. *relationships* is a list of :class:`Relationship` instances
        corresponding to the relationships this part maintains with other
        parts in the package. Returns a reference to self so a generative call
        is supported (e.g. ``newpart = Part(partcollection).load(partitem)``.
        
        """
        self.key           = part.key
        self.element       = part.element
        self.path          = part.path
        self.relationships = part.relationships
        return self
    

class CollectionPart(Part):
    """
    Parts that appear in groups, e.g. slides, slideLayouts, and others,
    require certain attributes that singleton parts do not. Those collection
    part specific attributes are implemented by this class.
    
    ... access to the presentation is required in order to reference the other
    parts in the presentation this part has relationships to.
    
    Inherits:
       :meth:`Part.load()`
    
    REFACTOR: Check construction signature, doesn't seem quite right.
    
    REFACTOR: Load and carry parttype for the type, not just the parttypekey.
    
    """
    def __init__(self, parent):
        super(CollectionPart, self).__init__(parent.presentation, parent.parttype)
        self.parent       = parent
        self.presentation = parent.presentation
        self.parttype     = parent.parttype
        self.path         = None  # this default will be overwritten on load by template parts
    
    @property
    def idx(self):
        """Return the index of this part in its parent collection."""
        return self.parent.index(self)
    

class HandoutMaster(CollectionPart):
    """**NOT YET FULLY IMPLEMENTED**. This class is currently a stub."""
    def __init__(self, parent):
        super(HandoutMaster, self).__init__(parent)
    

class Image(CollectionPart):
    """
    An image part. Note that an image part is distinct from a picture element
    that appears on a slide or other part. An image part corresponds to a
    distinct image file such as image1.png. A picture element is a placement
    of that image on one or more visual parts such as a slide, slide master,
    or slide layout. A theme can also include a picture (although I've yet to
    discover just how :).
    
    Also note that a single image part can appear on multiple visual parts
    while the image part (and therefore its image file) is stored only once in
    the package.
    
    **REFACTOR:** An image is not a TemplatePart strictly speaking. Images may
    be imported from the template, but they can also be added to slides using
    the API, so not all images are static and the images collection certainly
    needs to be dynamic. I'm thinking what this points up is that there's a
    distinction between *binary* and *XML* as well as between *static* and
    *dynamic*. Binary parts cannot be manipulated while static parts are
    simply not manipulated as a matter of convenience. All dynamic parts must
    be XML, so perhaps only three distinctions are strictly required. But I'm
    thinking it makes better sense for each part to have two new properties,
    format ('binary' | 'xml') and static (boolean).
    
    printerSettings1.bin is another example of a binary part.
    """
    def __init__(self, parent):
        super(Image, self).__init__(parent)
    

class NotesMaster(CollectionPart):
    """**NOT YET FULLY IMPLEMENTED**. This class is currently a stub."""
    def __init__(self, parent):
        super(NotesMaster, self).__init__(parent)
    

class PresProps(Part):
    """
    Presentation Properties part. Corresponds to package file
    ppt/presProps.xml.
    """
    def __init__(self, presentation):
        super(PresProps, self).__init__(presentation, PartType.lookup('presProps'))
    

class Slide(CollectionPart):
    """
    Slide part. Provides base slide functionality. Should not be constructed
    externally, use ``presentation.add_slide()`` to create new slides.
    
    :param id:          a unique id for the new slide
    :param slides:      parent collection
    :param slidelayout: instance of slide layout to inherit properties from
    :type  slidelayout: :class:`SlideLayout`
    :param name:        name property for slide, not used as a key
    :type  name:        string
    :rtype:             :class:`Slide`
    
    **REFACTOR:** See about inheriting from :class:`Part` rather than
    :class:`object`.
    """
    def __init__(self, id, slides, slidelayout, name):
        super(Slide, self).__init__(slides)
        self.id          = id
        self.key         = id
        self.slides      = slides
        self.parent      = slides
        self.slidelayout = slidelayout
        self.name        = name
        self.nsmap       = nsmap_subset(['a', 'p', 'r'])  # keep this here so remaining init can use it
        self.__id_iter   = util.intsequence()             # this one too ...
        self.element     = self.minimalelement
        self.shapes      = ShapeCollection(self)
        self.images      = []
    
    def setxml(self, xml):
        """
        **REFACTOR:** Not sure about this one, just trying it to get the first
        clean package build working. This should be replace with
        slide.element = ..., also add slide.xml read/write property to do the
        same thing but with the XML as a string instead of an ElementTree
        element.
        """
        self.element = etree.fromstring(xml)
    
    @property
    def _nextid(self):
        """
        ... iterator used for generating unique ids for the slide's shapes
        """
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
        """
        Return element containing the minimal XML for a slide, based on what
        is required by the XMLSchema. Note that in general, schema-minimal
        elements in the XML are not guaranteed to be loadable by PowerPoint,
        so test accordingly.
        """
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
        """
        Return the XML for the slide as a string. Used by the package to get
        the contents of the partfile for the slide during ``Package.save()``.
        """
        return etree.tostring(self.element, pretty_print=True)
    

class SlideLayout(CollectionPart):
    """
    Slide layout part. Corresponds to package files
    ppt/slideLayouts/slideLayout[1-9][0-9]*.xml.
    """
    def __init__(self, parent):
        super(SlideLayout, self).__init__(parent)
    

class SlideMaster(CollectionPart):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    
    TECHNOTE: In the Microsoft API, Master is a general type that all of
    SlideMaster, SlideLayout (CustomLayout), HandoutMaster, and NotesMaster
    inherit from. So might look into why that is and consider refactoring the
    various masters a bit later.
    """
    def __init__(self, parent):
        super(SlideMaster, self).__init__(parent)
    

class TableStyles(Part):
    """
    Optional table styles part. Corresponds to package file
    ppt/tableStyles.xml.
    """
    def __init__(self, presentation):
        super(TableStyles, self).__init__(presentation, PartType.lookup('tableStyles'))
    

class Theme(CollectionPart):
    """
    Theme part. Corresponds to package files ppt/theme/theme[1-9][0-9]*.xml.
    """
    def __init__(self, parent):
        super(Theme, self).__init__(parent)
    

class ViewProps(Part):
    """
    Optional view properties part. Corresponds to package file
    ppt/viewProps.xml.
    """
    def __init__(self, presentation):
        super(ViewProps, self).__init__(presentation, PartType.lookup('viewProps'))
    


# ============================================================================
# Shapes
# ============================================================================

class ShapeCollection(list):
    """
    Collection of all the Shapes in a slide. It can also serve for the shapes
    in SlideMaster and SlideLayout, and perhaps HandoutMaster, although
    dynamic behavior for those parts is not yet implemented so there's no need
    yet.
    
    **REFACTOR:** This should actually become GroupShape. A slide's shape tree
    (<spTree> element) is actually of type ``CT_GroupShape`` which can contain
    shapes and pictures, but can also contain other group shapes. So it needs
    to be modeled as an object hierarcy, a list won't do the trick by itself.
    Good news is that the GroupShape class will be able to be reused a fair
    amount, so should be worth the implementation effort.
    
    **REFACTOR:** Create Collection base class and inherit from that. Also
    have PartCollection and any other collections inherit from it. Place all
    code that's general to all collections up there.
    """
    def __init__(self, parent):
        super(ShapeCollection, self).__init__()
        self.parent = parent
    
    def append(self, shape):
        if not isinstance(shape, Shape):
            raise TypeError("Slide.shapes.append() called with incorrect type, (expected Shape, got %s)" % shape.__class__.__name__)
        list.append(self, shape)
        self.slide._append_shape(shape)
    

class Shape(object):
    """
    Base class for the various permutations of shape (not including pic, which
    is a distinct element type). This is actually an abstract class. Adding a
    raw Shape directly to a slide will likely cause a load error in
    PowerPoint. Use one of the fully fleshed-out subclasses such as
    PlaceholderText instead or modify the element directly to add the required
    bits for a complete shape.
    
    **TECHDEBT:** need to account for hierarchy of GroupShape, cSld/spTree is
    type GroupShape and GroupShape can contain other GroupShape elements in a
    hierarchy
    """
    def __init__(self, name):
        self.name    = name
        self.element = self.__minimalelement
    
    @property
    def __minimalelement(self):
        """
        Return an ElementTree element that contains all the elements and
        attributes of a shape required by the schema, initialized with default
        values where necessary or appropriate.
        """
        sp      = Element(qname('p', 'sp'))
        nvSpPr  = SubElement(sp     , qname('p', 'nvSpPr'  ) )
        cNvPr   = SubElement(nvSpPr , qname('p', 'cNvPr'   ) , id='None', name=self.name)
        cNvSpPr = SubElement(nvSpPr , qname('p', 'cNvSpPr' ) )
        spPr    = SubElement(sp     , qname('p', 'spPr'    ) )
        return sp
    

class PlaceholderText(Shape):
    """
    REFACTOR: This is just scrap code, this class needs to be re-implemented
    or at least reconsidered from first principles, it was initially written
    before any of the scaffolding class hierarchies were set up.
    
    Shape that holds text to be displayed in a placeholder specified by the
    slide layout.
    
    The placeholder_type provided must be one of ST_PlaceholderType defined in
    the PresentationML schema.
    
    The string supplied as the text parameter may contain carriage returns
    ('\\r') and or line feeds ('\\n'). Carriage returns will cause a new run
    to be created with a break (<a:br/>) intervening between it and the prior
    run. Line feeds will cause a new paragraph to be created to contain the
    text that follows the line feed.
    
    The index parameter is important for matching up the text with the right
    placeholder. I don't fully understand the logic used, but in one pptx file
    I've seen the <p:ph> idx attribute matched with that of the proper
    placeholder. The default is 0 and that seems to match the title
    placeholder.
    
    """
    pass
    
    # def __init__(self, name, placeholder_type, text, placeholder_index=None):
    #     if placeholder_type not in placeholder_types:
    #         raise TypeError("placeholder_type must be one of %s, got '%s'." % (placeholder_types, placeholder_type))
    #     super(PlaceholderText, self).__init__(name)
    #     self.placeholder_type  = placeholder_type
    #     self.placeholder_index = placeholder_index
    #     self.text = text
    #     # locate nodes we need to have handy (these are all required elements so we're assured they're present)
    #     sp      = self.element
    #     nvSpPr  = sp.find(qname('p', 'nvSpPr'))
    #     spPr    = sp.find(qname('p', 'spPr'))
    #     cNvSpPr = nvSpPr.find(qname('p', 'cNvSpPr'))
    #     # add shape lock to prevent grouping the placeholder element with others
    #     node = cNvSpPr.find(qname('a', 'spLocks'))
    #     spLocks = node if node else SubElement(cNvSpPr, qname('a', 'spLocks'))
    #     spLocks.attrib['noGrp'] = 'true'
    #     # add placeholder type so text can find its home on slide layout
    #     node = nvSpPr.find(qname('p', 'nvPr'))
    #     nvPr = node if node else SubElement(nvSpPr, qname('p', 'nvPr'))
    #     node = nvPr.find(qname('p', 'ph'))
    #     ph   = node if node else SubElement(nvPr, qname('p', 'ph'))
    #     ph.attrib['type'] = placeholder_type
    #     if placeholder_index:
    #         ph.attrib['idx'] = str(placeholder_index)
    #     # add txBody to hold text
    #     txBody = sp.find(qname('p', 'txBody'))
    #     if not node:
    #         txBody = Element(qname('p', 'txBody'))
    #         sp.insert(sp.index(spPr)+1, txBody)
    #     bodyPr = txBody.find(qname('a', 'bodyPr'))
    #     if not bodyPr:
    #         bodyPr = Element(qname('a', 'bodyPr'))
    #         txBody.insert(0, bodyPr)
    #     # get rid of any text that might be there already
    #     del txBody[1:]
    #     # insert the provided text
    #     lines = text.split('\n')
    #     for line in lines:
    #         p = SubElement(txBody, qname('a', 'p'))
    #         runs = line.split('\r')
    #         for run in runs:
    #             br = SubElement(p, qname('a', 'br')) if runs.index(run) > 0 else None
    #             r = SubElement(p, qname('a', 'r'))
    #             t = SubElement(r, qname('a', 't'))
    #             t.text = run
    # 
    # @property
    # def minimalelement(self):
    #     sp      = Element(qname('p', 'sp'))
    #     nvSpPr  = SubElement(sp     , qname('p', 'nvSpPr'  ) )
    #     cNvPr   = SubElement(nvSpPr , qname('p', 'cNvPr'   ) , id='None', name=self.name)
    #     cNvSpPr = SubElement(nvSpPr , qname('p', 'cNvSpPr' ) )
    #     spPr    = SubElement(sp     , qname('p', 'spPr'    ) )
    #     return sp
    # 


