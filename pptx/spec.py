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

from pptx.packaging import PartType


# ============================================================================
# Part Type specs
# ============================================================================
# Not yet included:
# * Font Part (font1.fntdata) 15.2.13
# * chart         : 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
# * themeOverride : 'application/vnd.openxmlformats-officedocument.themeOverride+xml'
# * several others, especially DrawingML parts ...
# ============================================================================

#TODO: Add Printer Settings Part (ECMA-376-1 Section 15.2.15). Also check out
#      other shared parts in section 15.

parttypes =\
    [ PartType( rootname     = 'app'                         # ECMA-376-1 15.2.12.3 (Extended File Properties Part)
              , file_ext     = 'xml'
              , name         = 'Application-Defined File Properties Part'
              , cardinality  = 'single'
              , required     = False                         # optional according to spec, don't know if package will load without one though
              , location     = '/docProps'
              , has_rel_item = 'never'                       # has_rel_item should be construed as "can have a relationship item" rather than "always has a relationship item"
              , rels_from    = ['package']
              , content_type = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
              , namespace    = 'http://schemas.openxmlformats.org/officeDocument/2006/extendedProperties'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extendedProperties'
              )
    , PartType( rootname     = 'core'                        # ECMA-376-1 15.2.12.1
              , file_ext     = 'xml'
              , name         = 'Core File Properties Part'   # 'Core' as in Dublin Core
              , cardinality  = 'single'
              , required     = False
              , location     = '/docProps'
              , has_rel_item = 'never'
              , rels_from    = ['package']
              , content_type = 'application/vnd.openxmlformats-package.core-properties+xml'
              , namespace    = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
              , relationship = 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties'
              )
    , PartType( rootname     = 'custom'                      # ECMA-376-1 15.2.12.2
              , file_ext     = 'xml'
              , name         = 'Custom File Properties Part'
              , cardinality  = 'single'
              , required     = False
              , location     = '/docProps'
              , has_rel_item = 'never'
              , rels_from    = ['package']
              , content_type = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
              , namespace    = 'http://schemas.openxmlformats.org/officeDocument/2006/customProperties'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperties'
              )
    , PartType( rootname     = 'commentAuthors'              # ECMA-376-1 13.3.1
              , file_ext     = 'xml'
              , name         = 'Comment Authors Part'
              , cardinality  = 'single'
              , required     = False
              , location     = '/ppt'
              , has_rel_item = 'never'
              , rels_from    = ['presentation']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors'
              )
    , PartType( rootname     = 'presentation'                # ECMA-376-1 13.3.6
              , file_ext     = 'xml'
              , name         = 'Presentation Part'
              , cardinality  = 'single'
              , required     = True
              , location     = '/ppt'
              , has_rel_item = 'always'
              , rels_from    = ['package']
              # there are three possible Content Type values for the
              # presentation part, this one below for presentations, then one
              # for slideshows (present-only decks) and a third one for
              # PowerPoint template files (.potx files) see spec for more
              # details if needed (ECMA-376 13.3.6).
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
              )
    , PartType( rootname     = 'presProps'                   # ECMA-376-1 13.3.7
              , file_ext     = 'xml'
              , name         = 'Presentation Properties Part'
              , cardinality  = 'single'
              , required     = True
              , location     = '/ppt'
              , has_rel_item = 'never'
              , rels_from    = ['presentation']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
              )
    , PartType( rootname     = 'tableStyles'                 # ECMA-376-1 14.2.9
              , file_ext     = 'xml'
              , name         = 'Table Styles Part'
              , cardinality  = 'single'
              , required     = False
              , location     = '/ppt'
              , has_rel_item = 'never'
              , rels_from    = ['presentation']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'
              , namespace    = 'http://schemas.openxmlformats.org/drawingml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'
              )
    , PartType( rootname     = 'viewProps'                   # ECMA-376-1 13.3.13
              , file_ext     = 'xml'
              , name         = 'View Properties Part'
              , cardinality  = 'single'
              , required     = False
              , location     = '/ppt'
              , has_rel_item = 'never'
              , rels_from    = ['presentation']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps'
              )

    # Part Collections
    # =======================================================================
    
    , PartType( rootname     = 'comment'                     # ECMA-376-1 13.3.2
              , file_ext     = 'xml'
              , name         = 'Comments Part'
              , cardinality  = 'multiple'
              , required     = False
              , location     = '/ppt/comments'
              , has_rel_item = 'never'
              , rels_from    = ['slide']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
              )
    , PartType( rootname     = 'handoutMaster'               # ECMA-376-1 13.3.3
              , file_ext     = 'xml'
              , name         = 'Handout Master Part'
              , cardinality  = 'multiple'                    # actually can only be one according to spec, but behaves like part collection (handoutMasters folder, handoutMaster1.xml, etc.)
              , required     = False
              , location     = '/ppt/handoutMasters'
              , has_rel_item = 'always'
              , rels_from    = ['presentation']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster'
              )
    , PartType( rootname     = 'image'                       # ECMA-376-1 15.2.14
              , file_ext     = ''                            # n/a, various extensions and they're handled with <Default> elements in [Content_Types].xml
              , name         = 'Image Part'
              , cardinality  = 'multiple'
              , required     = False
              , location     = '/ppt/media'
              , has_rel_item = 'never'
              # allowing relationships to image from notesSlide might be an error in the spec, UI doesn't allow it
              , rels_from    = ['handoutMaster', 'notesSlide', 'notesMaster', 'slide', 'slideLayout', 'slideMaster']
              , content_type = ''                            # n/a, various MIME types and they're handled with <Default> elements in [Content_Types].xml
              , namespace    = ''                            # n/a, not an XML format
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
              )
    , PartType( rootname     = 'notesMaster'                 # ECMA-376-1 13.3.4
              , file_ext     = 'xml'
              , name         = 'Notes Master Part'
              , cardinality  = 'multiple'                    # actually can only be one according to spec, but behaves like part collection (notesMasters folder, notesMaster1.xml, etc.)
              , required     = False
              , location     = '/ppt/notesMasters'
              , has_rel_item = 'always'
              , rels_from    = ['presentation', 'notesSlide']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'
              )
    , PartType( rootname     = 'notesSlide'                  # ECMA-376-1 13.3.5
              , file_ext     = 'xml'
              , name         = 'Notes Slide Part'
              , cardinality  = 'multiple'
              , required     = False
              , location     = '/ppt/notesSlides'
              , has_rel_item = 'always'
              , rels_from    = ['slide']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
              )
    , PartType( rootname     = 'slide'                       # ECMA-376-1 13.3.8
              , file_ext     = 'xml'
              , name         = 'Slide Part'
              , cardinality  = 'multiple'
              , required     = True                          # spec is ambiguous, should check to see if a deck with no slides will load.
              , location     = '/ppt/slides'
              , has_rel_item = 'always'
              , rels_from    = ['presentation', 'notesSlide']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
              )
    , PartType( rootname     = 'slideLayout'                 # ECMA-376-1 13.3.9
              , file_ext     = 'xml'
              , name         = 'Slide Layout Part'
              , cardinality  = 'multiple'
              , required     = True
              , location     = '/ppt/slideLayouts'
              , has_rel_item = 'always'
              , rels_from    = ['slide', 'slideMaster']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
              )
    , PartType( rootname     = 'slideMaster'                 # ECMA-376-1 13.3.10
              , file_ext     = 'xml'
              , name         = 'Slide Master Part'
              , cardinality  = 'multiple'
              , required     = True
              , location     = '/ppt/slideMasters'
              , has_rel_item = 'always'
              , rels_from    = ['presentation', 'slideLayout']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
              )
    , PartType( rootname     = 'tag'                         # ECMA-376-1 13.3.12
              , file_ext     = 'xml'
              , name         = 'User-Defined Tags Part'
              , cardinality  = 'multiple'
              , required     = False
              , location     = '/ppt/tags'
              , has_rel_item = 'never'
              , rels_from    = ['presentation', 'slide']
              , content_type = 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml'
              , namespace    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'
              )
    , PartType( rootname     = 'theme'                       # ECMA-376-1 14.2.7
              , file_ext     = 'xml'
              , name         = 'Theme Part'
              , cardinality  = 'multiple'
              , required     = True                          # spec indicates theme part is optional, but I've never seen a .pptx without one
              , location     = '/ppt/theme'
              , has_rel_item = 'optional'                    # can have _rels items, but only if the theme contains one or more images
              , rels_from    = ['presentation', 'handoutMaster', 'notesMaster', 'slideMaster']
              , content_type = 'application/vnd.openxmlformats-officedocument.theme+xml'
              , namespace    = 'http://schemas.openxmlformats.org/drawingml/2006/main'
              , relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
              )
    ]


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


