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

PTS_CARDINALITY_SINGLETON = 'singleton'
PTS_CARDINALITY_TUPLE     = 'tuple'
PTS_HASRELS_ALWAYS        = 'always'
PTS_HASRELS_NEVER         = 'never'
PTS_HASRELS_OPTIONAL      = 'optional'

CT_PRESENTATION = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
CT_SLIDE        = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
CT_SLIDELAYOUT  = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
CT_SLIDEMASTER  = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
CT_SLIDESHOW    = 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml'
CT_TEMPLATE     = 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'

RT_HANDOUTMASTER  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster'
RT_NOTESMASTER    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'
RT_OFFICEDOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
RT_PRESPROPS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
RT_SLIDE          = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
RT_SLIDELAYOUT    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
RT_SLIDEMASTER    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
RT_TABLESTYLES    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'
RT_THEME          = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
RT_VIEWPROPS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps'

pml_parttypes =\
    { 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml':
       # ECMA-376-1 13.3.1
       { 'basename'     : 'commentAuthors'
       , 'ext'          : '.xml'
       , 'name'         : 'Comment Authors Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : False
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['presentation']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml':
       # ECMA-376-1 13.3.2
       { 'basename'     : 'comment'
       , 'ext'          : '.xml'
       , 'name'         : 'Comments Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/comments'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['slide']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml':
       # ECMA-376-1 13.3.3
       { 'basename'     : 'handoutMaster'
       , 'ext'          : '.xml'
       , 'name'         : 'Handout Master Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE  # actually can only be one according to spec, but behaves like part collection (handoutMasters folder, handoutMaster1.xml, etc.)
       , 'required'     : False
       , 'baseURI'      : '/ppt/handoutMasters'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['presentation']
       , 'reltype'      : RT_HANDOUTMASTER
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml':
       # ECMA-376-1 13.3.4
       { 'basename'     : 'notesMaster'
       , 'ext'          : '.xml'
       , 'name'         : 'Notes Master Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE  # actually can only be one according to spec, but behaves like part collection (notesMasters folder, notesMaster1.xml, etc.)
       , 'required'     : False
       , 'baseURI'      : '/ppt/notesMasters'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['presentation', 'notesSlide']
       , 'reltype'      : RT_NOTESMASTER
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml':
       # ECMA-376-1 13.3.5
       { 'basename'     : 'notesSlide'
       , 'ext'          : '.xml'
       , 'name'         : 'Notes Slide Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/notesSlides'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['slide']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
       }
    , CT_PRESENTATION:
       # ECMA-376-1 13.3.6
       # one of three possible Content Type values for presentation part
       { 'basename'     : 'presentation'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['package']
       , 'reltype'      : RT_OFFICEDOCUMENT
       }
    , CT_TEMPLATE:
       # ECMA-376-1 13.3.6
       # one of three possible Content Type values for presentation part
       { 'basename'     : 'presentation'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['package']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
       }
    , CT_SLIDESHOW:
       # ECMA-376-1 13.3.6
       # one of three possible Content Type values for presentation part
       { 'basename'     : 'presentation'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['package']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml':
       # ECMA-376-1 13.3.7
       { 'basename'     : 'presProps'
       , 'ext'          : '.xml'
       , 'name'         : 'Presentation Properties Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : True
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['presentation']
       , 'reltype'      : RT_PRESPROPS
       }
    , CT_SLIDE:
       # ECMA-376-1 13.3.8
       { 'basename'     : 'slide'
       , 'ext'          : '.xml'
       , 'name'         : 'Slide Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/slides'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['presentation', 'notesSlide']
       , 'reltype'      : RT_SLIDE
       }
    , CT_SLIDELAYOUT:
       # ECMA-376-1 13.3.9
       { 'basename'     : 'slideLayout'
       , 'ext'          : '.xml'
       , 'name'         : 'Slide Layout Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : True
       , 'baseURI'      : '/ppt/slideLayouts'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['slide', 'slideMaster']
       , 'reltype'      : RT_SLIDELAYOUT
       }
    , CT_SLIDEMASTER:
       # ECMA-376-1 13.3.10
       { 'basename'     : 'slideMaster'
       , 'ext'          : '.xml'
       , 'name'         : 'Slide Master Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : True
       , 'baseURI'      : '/ppt/slideMasters'
       , 'has_rels'     : PTS_HASRELS_ALWAYS
       , 'rels_from'    : ['presentation', 'slideLayout']
       , 'reltype'      : RT_SLIDEMASTER
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml':
       # ECMA-376-1 13.3.12
       { 'basename'     : 'tag'
       , 'ext'          : '.xml'
       , 'name'         : 'User-Defined Tags Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/tags'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['presentation', 'slide']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml':
       # ECMA-376-1 13.3.13
       { 'basename'     : 'viewProps'
       , 'ext'          : '.xml'
       , 'name'         : 'View Properties Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : False
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['presentation']
       , 'reltype'      : RT_VIEWPROPS
       }
    , 'application/vnd.openxmlformats-officedocument.theme+xml':
       # ECMA-376-1 14.2.7
       { 'basename'     : 'theme'
       , 'ext'          : '.xml'
       , 'name'         : 'Theme Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : True  # spec indicates theme part is optional, but I've never seen a .pptx without one
       , 'baseURI'      : '/ppt/theme'
       , 'has_rels'     : PTS_HASRELS_OPTIONAL  # can have _rels items, but only if the theme contains one or more images
       , 'rels_from'    : ['presentation', 'handoutMaster', 'notesMaster', 'slideMaster']
       , 'reltype'      : RT_THEME
       }
    , 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml':
       # ECMA-376-1 14.2.9
       { 'basename'     : 'tableStyles'
       , 'ext'          : '.xml'
       , 'name'         : 'Table Styles Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : False
       , 'baseURI'      : '/ppt'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['presentation']
       , 'reltype'      : RT_TABLESTYLES
       }
    , 'application/vnd.openxmlformats-package.core-properties+xml':
       # ECMA-376-1 15.2.12.1
       { 'basename'     : 'core'
       , 'ext'          : '.xml'
       , 'name'         : 'Core File Properties Part'  # 'Core' as in Dublin Core
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : False
       , 'baseURI'      : '/docProps'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['package']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties'
       }
    , 'application/vnd.openxmlformats-officedocument.custom-properties+xml':
       # ECMA-376-1 15.2.12.2
       { 'basename'     : 'custom'
       , 'ext'          : '.xml'
       , 'name'         : 'Custom File Properties Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : False
       , 'baseURI'      : '/docProps'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['package']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperties'
       }
    , 'application/vnd.openxmlformats-officedocument.extended-properties+xml':
       # ECMA-376-1 15.2.12.3 (Extended File Properties Part)
       { 'basename'     : 'app'
       , 'ext'          : '.xml'
       , 'name'         : 'Application-Defined File Properties Part'
       , 'cardinality'  : PTS_CARDINALITY_SINGLETON
       , 'required'     : False
       , 'baseURI'      : '/docProps'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['package']
       , 'content_type' : 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extendedProperties'
       }
    ,  'image/gif':
       # ECMA-376-1 15.2.14
       { 'basename'     : 'image'
       , 'ext'          : '.gif'
       , 'name'         : 'Image Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/media'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
       }
    ,  'image/jpeg':
       # ECMA-376-1 15.2.14
       { 'basename'     : 'image'
       , 'ext'          : '.jpeg'
       , 'name'         : 'Image Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/media'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
       }
    ,  'image/png':
       # ECMA-376-1 15.2.14
       { 'basename'     : 'image'
       , 'ext'          : '.png'
       , 'name'         : 'Image Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/media'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
       }
    ,  'image/x-emf':
       # ECMA-376-1 15.2.14
       { 'basename'     : 'image'
       , 'ext'          : '.emf'
       , 'name'         : 'Image Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/media'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
       }
    ,  'application/vnd.openxmlformats-officedocument.presentationml.printerSettings':
       # ECMA-376-1 15.2.15
       { 'basename'     : 'printerSettings'
       , 'ext'          : '.bin'
       , 'name'         : 'Printer Settings Part'
       , 'cardinality'  : PTS_CARDINALITY_TUPLE
       , 'required'     : False
       , 'baseURI'      : '/ppt/printerSettings'
       , 'has_rels'     : PTS_HASRELS_NEVER
       , 'rels_from'    : ['presentation']
       , 'reltype'      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
       }
    }


# ============================================================================
# default_content_types
# ============================================================================
# Default file extension to MIME type mapping. This is used as a reference for
# adding <Default> elements to [Content_Types].xml for parts like media files.
#     TODO: I've seen .wmv elements in the media folder of at least one
# presentation, might need to add an entry for that and perhaps other rich
# media PowerPoint allows to be embedded (e.g. audio, movie, object, ...).
# ============================================================================

default_content_types =\
    { '.bin'     : 'application/vnd.openxmlformats-officedocument.presentationml.printerSettings'
    , '.emf'     : 'image/x-emf'
    , '.fntdata' : 'application/x-fontdata'
    , '.gif'     : 'image/gif'
    , '.jpe'     : 'image/jpeg'
    , '.jpeg'    : 'image/jpeg'
    , '.jpg'     : 'image/jpeg'
    , '.png'     : 'image/png'
    , '.rels'    : 'application/vnd.openxmlformats-package.relationships+xml'
    , '.tif'     : 'image/tiff'
    , '.tiff'    : 'image/tiff'
    , '.wmf'     : 'image/x-wmf'
    , '.xlsx'    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    , '.xml'     : 'application/xml'
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


def namespaces(*prefixes):
    """
    Return a dict containing the subset namespace prefix mappings specified by
    *prefixes*. Any number of namespace prefixes can be supplied, e.g.
    namespaces('a', 'r', 'p').
    """
    namespaces = {}
    for prefix in prefixes:
        namespaces[prefix] = nsmap[prefix]
    return namespaces


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


