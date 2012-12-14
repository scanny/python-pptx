# -*- coding: utf-8 -*-
#
# spec.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

'''Constant values from the ECMA-376 spec that are needed for XML generation
and packaging, and a utility function or two for accessing some of them.'''


# ============================================================================
# PresentationML Part Type specs
# ============================================================================
# Keyed by content type
# Not yet included:
# * Font Part (font1.fntdata) 15.2.13
# * chart         : 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
# * themeOverride : 'application/vnd.openxmlformats-officedocument.themeOverride+xml'
# * several others, especially DrawingML parts ...
#
# TODO: Also check out other shared parts in section 15.
# ============================================================================

pml_parttypes =\
    { 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml':
       # ECMA-376-1 13.3.1
       { 'basename'     : 'commentAuthors'
       , 'ext'          : '.xml'
       , 'name'         : 'Comment Authors Part'
       , 'cardinality'  : 'single'
       , 'required'     : False
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['presentation']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml':
       # ECMA-376-1 13.3.2
       { 'basename'     : 'comment'
       , 'ext'          : '.xml'
       , 'name'         : 'Comments Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : False
       , 'baseURI'      : '/ppt/comments'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['slide']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml':
       # ECMA-376-1 13.3.3
       { 'basename'     : 'handoutMaster'
       , 'ext'          : '.xml'
       , 'name'         : 'Handout Master Part'
       , 'cardinality'  : 'multiple'  # actually can only be one according to spec, but behaves like part collection (handoutMasters folder, handoutMaster1.xml, etc.)
       , 'required'     : False
       , 'baseURI'      : '/ppt/handoutMasters'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['presentation']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml':
       # ECMA-376-1 13.3.4
       { 'basename'     : 'notesMaster'
       , 'ext'          : '.xml'
       , 'name'         : 'Notes Master Part'
       , 'cardinality'  : 'multiple'  # actually can only be one according to spec, but behaves like part collection (notesMasters folder, notesMaster1.xml, etc.)
       , 'required'     : False
       , 'baseURI'      : '/ppt/notesMasters'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['presentation', 'notesSlide']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml':
       # ECMA-376-1 13.3.5
       { 'basename'     : 'notesSlide'
       , 'ext'          : '.xml'
       , 'name'         : 'Notes Slide Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : False
       , 'baseURI'      : '/ppt/notesSlides'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['slide']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml':
       # ECMA-376-1 13.3.6
       # one of three possible Content Type values for presentation part
       { 'basename'     : 'presentation'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Part'
       , 'cardinality'  : 'single'
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['package']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml':
       # ECMA-376-1 13.3.6
       # one of three possible Content Type values for presentation part
       { 'basename'     : 'presentation'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Part'
       , 'cardinality'  : 'single'
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['package']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml':
       # ECMA-376-1 13.3.6
       # one of three possible Content Type values for presentation part
       { 'basename'     : 'presentation'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Part'
       , 'cardinality'  : 'single'
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['package']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml':
       # ECMA-376-1 13.3.7
       { 'basename'     : 'presProps'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Properties Part'
       , 'cardinality'  : 'single'
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['presentation']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
       # ECMA-376-1 13.3.8
       { 'basename'     : 'slide'
       , 'ext'          : '.xml'
       , 'name'         : 'Slide Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : True  # spec is ambiguous, should check to see if a deck with no slides will load.
       , 'baseURI'      : '/ppt/slides'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['presentation', 'notesSlide']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
       # ECMA-376-1 13.3.9
       { 'basename'     : 'slideLayout'
       , 'ext'          : '.xml'
       , 'name'         : 'Slide Layout Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : True
       , 'baseURI'      : '/ppt/slideLayouts'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['slide', 'slideMaster']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml':
       # ECMA-376-1 13.3.10
       { 'basename'     : 'slideMaster'
       , 'ext'          : '.xml'
       , 'name'         : 'Slide Master Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : True
       , 'baseURI'      : '/ppt/slideMasters'
       , 'has_rels'     : 'always'
       , 'rels_from'    : ['presentation', 'slideLayout']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml':
       # ECMA-376-1 13.3.12
       { 'basename'     : 'tag'
       , 'ext'          : '.xml'
       , 'name'         : 'User-Defined Tags Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : False
       , 'baseURI'      : '/ppt/tags'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['presentation', 'slide']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml':
       # ECMA-376-1 13.3.13
       { 'basename'     : 'viewProps'
       , 'ext'          : '.xml'
       , 'name'         : 'View Properties Part'
       , 'cardinality'  : 'single'
       , 'required'     : False
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['presentation']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps'
       }
    , 'application/vnd.openxmlformats-officedocument.theme+xml':
       # ECMA-376-1 14.2.7
       { 'basename'     : 'theme'
       , 'ext'          : '.xml'
       , 'name'         : 'Theme Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : True        # spec indicates theme part is optional, but I've never seen a .pptx without one
       , 'baseURI'      : '/ppt/theme'
       , 'has_rels'     : 'optional'  # can have _rels items, but only if the theme contains one or more images
       , 'rels_from'    : ['presentation', 'handoutMaster', 'notesMaster', 'slideMaster']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml':
       # ECMA-376-1 14.2.9
       { 'basename'     : 'tableStyles'
       , 'ext'          : '.xml'
       , 'name'         : 'Table Styles Part'
       , 'cardinality'  : 'single'
       , 'required'     : False
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['presentation']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'
       }
    , 'application/vnd.openxmlformats-package.core-properties+xml':
       # ECMA-376-1 15.2.12.1
       { 'basename'     : 'core'
       , 'ext'          : '.xml'
       , 'name'         : 'Core File Properties Part'  # 'Core' as in Dublin Core
       , 'cardinality'  : 'single'
       , 'required'     : False
       , 'baseURI'      : '/docProps'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['package']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties'
       }
    , 'application/vnd.openxmlformats-officedocument.custom-properties+xml':
       # ECMA-376-1 15.2.12.2
       { 'basename'     : 'custom'
       , 'ext'          : '.xml'
       , 'name'         : 'Custom File Properties Part'
       , 'cardinality'  : 'single'
       , 'required'     : False
       , 'baseURI'      : '/docProps'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['package']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperties'
       }
    , 'application/vnd.openxmlformats-officedocument.extended-properties+xml':
       # ECMA-376-1 15.2.12.3 (Extended File Properties Part)
       { 'basename'     : 'app'
       , 'ext'          : '.xml'
       , 'name'         : 'Application-Defined File Properties Part'
       , 'cardinality'  : 'single'
       , 'required'     : False
       , 'baseURI'      : '/docProps'
       , 'has_rels'     : 'never'      # has_rel_item should be construed as "can have a relationship item" rather than "always has a relationship item"
       , 'rels_from'    : ['package']
       , 'content_type' : 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extendedProperties'
       }
    ,  'image/gif':
       # ECMA-376-1 15.2.14
       { 'basename'     : 'image'
       , 'ext'          : '.gif'
       , 'name'         : 'Image Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : False
       , 'baseURI'      : '/ppt/media'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
       }
    ,  'image/jpeg':
       # ECMA-376-1 15.2.14
       { 'basename'     : 'image'
       , 'ext'          : '.jpeg'
       , 'name'         : 'Image Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : False
       , 'baseURI'      : '/ppt/media'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
       }
    ,  'image/png':
       # ECMA-376-1 15.2.14
       { 'basename'     : 'image'
       , 'ext'          : '.png'
       , 'name'         : 'Image Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : False
       , 'baseURI'      : '/ppt/media'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
       }
    ,  'application/vnd.openxmlformats-officedocument.presentationml.printerSettings':
       # ECMA-376-1 15.2.15
       { 'basename'     : 'printerSettings'
       , 'ext'          : '.bin'
       , 'name'         : 'Printer Settings Part'
       , 'cardinality'  : 'multiple'
       , 'required'     : False
       , 'baseURI'      : '/ppt/printerSettings'
       , 'has_rels'     : 'never'
       , 'rels_from'    : ['presentation']
       , 'rel_type'     : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
       }
    }


# class PartType(object):
#     """
#     Reference to the characteristics of the various package parts, as defined
#     in ECMA-376.
#     
#     .. attribute:: rootname
#        
#        The root of the part's filename within the package. For example,
#        rootname for slideLayout1.xml is 'slideLayout'. Note that the part's
#        rootname is also used as its key value.
#     
#     .. attribute:: file_ext
#        
#        The extension of the part's filename within the package. For example,
#        file_ext for the presentation part (presentation.xml) is 'xml'.
#     
#     .. attribute:: name
#        
#        Full name of the part type, as inferred from prose in ECMA-376. This
#        attribute is not formally specified, so there may be some variation in
#        actual usage.
#     
#     .. attribute:: cardinality
#        
#        One of 'single' or 'multiple', specifying whether the part is a
#        singleton or tuple within the package. ``presentation.xml`` is an
#        example of a singleton part. ``slideLayout4.xml`` is an example of a
#        tuple part. The term *tuple* in this context is drawn from set theory
#        in math and has no direct relationship to the Python tuple class.
#     
#     .. attribute:: required
#        
#        Boolean expressing whether at least one instance of this part type must
#        appear in the package. ``presentation`` is an example of a required
#        part type. ``notesMaster`` is an example of a part type that is
#        optional.
#     
#     .. attribute:: location
#        
#        The package-relative path of the directory in which part files for this
#        type are stored. For example, location for ``slideLayout`` is
#        '/ppt/slideLayout'. The leading slash corresponds to the root of the
#        package (zip file). Note that directories in the actual package zip
#        file do not contain this leading slash (otherwise they would be
#        placed in the root directory when the zip file was expanded).
#     
#     .. attribute:: has_rel_item
#        
#        One of 'always', 'never', or 'optional', indicating whether parts of
#        this type have a corresponding relationship item, or "rels file".
#     
#     .. attribute:: rels_from
#        
#        List of part type keys for parts that may have relationships to this
#        part type.
#     
#     .. attribute:: content_type
#        
#        MIME-type-like string that distinguishes the content of parts of this
#        type from simple XML. For example, the content_type of a theme part is
#        ``application/vnd.openxmlformats-officedocument.theme+xml``. Each
#        part's content type is written in the content types item located in the
#        root directory of the package ([Content_Types].xml).
#     
#     .. attribute:: namespace
#        
#        The XML namespace of the root element for this part.
#     
#     .. attribute:: relationshiptype
#        
#        A URL that identifies this part type in rels files. For example,
#        relationshiptype for ``slides/slide1.xml`` is
#        ``http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide``
#     
#     .. attribute:: format
#        
#        One of 'xml' or 'binary'.
#     """
#     
#     _dict = {}
#     
#     def __init__(self, key, ptdict):
#         self.rootname         = key                     # e.g. 'slideMaster'
#         self.ext              = ptdict['ext'         ]  # e.g. 'xml'
#         self.name             = ptdict['name'        ]  # e.g. 'Core File Properties Part'
#         self.cardinality      = ptdict['cardinality' ]  # e.g. 'single' or 'multiple'
#         self.required         = ptdict['required'    ]  # e.g. False
#         self.baseURI          = ptdict['baseURI'     ]  # e.g. '/ppt/slideMasters'
#         self.has_rel_item     = ptdict['has_rel_item']  # e.g. 'always', 'never', or 'optional'
#         self.rels_from        = ptdict['rels_from'   ]  # e.g. ['package']
#         self.namespace        = ptdict['namespace'   ]  # e.g. 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
#         self.relationshiptype = ptdict['relationship']  # e.g. 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties'
#         self.format           = 'xml' if self.file_ext == 'xml' else 'binary'
#     
#     @property
#     def key(self):
#         return self.rootname
#     
#     @classmethod
#     def lookup(cls, parttypekeyname):
#         if parttypekeyname not in cls._dict:
#             raise KeyError("""PartType.lookup() lookup failed with key '%s'""" % (parttypekeyname))
#         return cls._dict[parttypekeyname]
#     
#     @property
#     def pkgreldir(self):
#         return self.location[1:] if self.location.startswith('/') else self.location
#     
#     # NOTE: This doesn't work for the package relationships item, only for
#     # part relationship items!
#     # Return the part path format used in part relationship items (files),
#     # a path that is relative to the directory containing the frompart.
#     def reltargetdir(self, fromlocation):
#         if fromlocation == '/ppt':  # that means this is the presentation part, the only part with that path that has a relationship item
#             if self.location == '/ppt':              # means topart is a singleton part
#                 return ''
#             elif self.location.startswith('/ppt/'):  # means topart is a collection part
#                 return self.location[5:]
#             else:
#                 raise NotImplementedError('''Unrecognized fromlocation '%s' received by PartType.reltargetdir()''' % fromlocation)
#         if self.location.startswith('/ppt/'):
#             return '../%s' % self.location[5:]
#         raise NotImplementedError("""PartType.reltargetdir() is not yet implemented for part type '%s'""" % self.rootname)
#     
#     @property
#     def zipdir(self):
#         return self.location[1:] if self.location.startswith('/') else self.location
#     
# 
# # This short code passage initializes PartType with all the defined part types
# for key in iter(pml_parttypes):
#     PartType._dict[key] = PartType(key, pml_parttypes[key])
# del key


# ============================================================================
# ext2mime_map
# ============================================================================
# Default file extension to MIME type mapping. This is used as a reference for
# adding <Default> elements to [Content_Types].xml for parts like media files.
#     TODO: I've seen .wmv elements in the media folder of at least one
# presentation, might need to add an entry for that and perhaps other rich
# media PowerPoint allows to be embedded (e.g. audio, movie, object, ...).
# ============================================================================

ext2mime_map =\
    { 'bin'     : 'application/vnd.openxmlformats-officedocument.presentationml.printerSettings'
    , 'emf'     : 'image/x-emf'
    , 'fntdata' : 'application/x-fontdata'
    , 'gif'     : 'image/gif'
    , 'jpe'     : 'image/jpeg'
    , 'jpeg'    : 'image/jpeg'
    , 'jpg'     : 'image/jpeg'
    , 'png'     : 'image/png'
    , 'rels'    : 'application/vnd.openxmlformats-package.relationships+xml'
    , 'tif'     : 'image/tiff'
    , 'tiff'    : 'image/tiff'
    , 'wmf'     : 'image/x-wmf'
    , 'xlsx'    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    , 'xml'     : 'application/xml'
    }


# ============================================================================
# nsmap
# ============================================================================
# namespace prefix to namespace name map
# ============================================================================

nsmap = {}
# Text Content
nsmap['mv'      ] = 'urn:schemas-microsoft-com:mac:vml'
nsmap['mo'      ] = 'http://schemas.microsoft.com/office/mac/office/2008/main'
nsmap['ve'      ] = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
nsmap['o'       ] = 'urn:schemas-microsoft-com:office:office'
nsmap['r'       ] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
nsmap['m'       ] = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
nsmap['v'       ] = 'urn:schemas-microsoft-com:vml'
nsmap['w'       ] = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
nsmap['p'       ] = 'http://schemas.openxmlformats.org/presentationml/2006/main'
nsmap['sl'      ] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
nsmap['w10'     ] = 'urn:schemas-microsoft-com:office:word'
nsmap['wne'     ] = 'http://schemas.microsoft.com/office/word/2006/wordml'
nsmap['i'       ] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
# Drawing
nsmap['wp'      ] = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
nsmap['pd'      ] = 'http://schemas.openxmlformats.org/drawingml/2006/presentationDrawing'
nsmap['a'       ] = 'http://schemas.openxmlformats.org/drawingml/2006/main'
nsmap['pic'     ] = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
# Properties (core and extended)
nsmap['cp'      ] = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
nsmap['dc'      ] = 'http://purl.org/dc/elements/1.1/'
nsmap['dcterms' ] = 'http://purl.org/dc/terms/'
nsmap['dcmitype'] = 'http://purl.org/dc/dcmitype/'
nsmap['xsi'     ] = 'http://www.w3.org/2001/XMLSchema-instance'
nsmap['ep'      ] = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
# Content Types (Note: the ct prefix is not actually used anywhere in .pptx files)
nsmap['ct'      ] = 'http://schemas.openxmlformats.org/package/2006/content-types'
# Package Relationships (Note: the pr prefix is not actually used anywhere in .pptx files)
nsmap['pr'      ] = 'http://schemas.openxmlformats.org/package/2006/relationships'


# generate a subset of the complete namespace map suitable for using in an
# Element call (to specify which namespaces will be used in that XML document)
def nsmap_subset(prefixes):
    nsmap_subset = {}
    for prefix in prefixes:
        nsmap_subset[prefix] = nsmap[prefix]
    return nsmap_subset


# Return a qualified name (QName) for an XML element or attribute in Clark
# notation (e.g. "{http://www.w3.org/1999/xhtml}body instead of "html:body")
# by looking up the specified namespace prefix in the overall namespace map
# (nsmap) above. Google on "xml clark notation" for more on that topic.
def qname(nsprefix, localtagname):
    uri = nsmap[nsprefix]
    return '{%s}%s' % (uri, localtagname)


# # valid values for <p:ph> type attribute (ST_PlaceholderType)
# placeholder_types = ('title', 'body', 'ctrTitle', 'subTitle', 'dt', 'sldNum',
#                      'ftr', 'hdr', 'obj', 'chart', 'tbl', 'clipArt', 'dgm',
#                      'media', 'sldImg', 'pic')


