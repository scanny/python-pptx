# -*- coding: utf-8 -*-
#
# spec.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
Constant values from the ECMA-376 spec that are needed for XML generation and
packaging, and a utility function or two for accessing some of them.
"""


# ============================================================================
# Placeholder constants
# ============================================================================

# valid values for <p:ph> type attribute (ST_PlaceholderType)
# -----------------------------------------------------------
PH_TYPE_TITLE = 'title'
PH_TYPE_BODY = 'body'
PH_TYPE_CTRTITLE = 'ctrTitle'
PH_TYPE_SUBTITLE = 'subTitle'
PH_TYPE_DT = 'dt'
PH_TYPE_SLDNUM = 'sldNum'
PH_TYPE_FTR = 'ftr'
PH_TYPE_HDR = 'hdr'
PH_TYPE_OBJ = 'obj'
PH_TYPE_CHART = 'chart'
PH_TYPE_TBL = 'tbl'
PH_TYPE_CLIPART = 'clipArt'
PH_TYPE_DGM = 'dgm'
PH_TYPE_MEDIA = 'media'
PH_TYPE_SLDIMG = 'sldImg'
PH_TYPE_PIC = 'pic'

# valid values for <p:ph> orient attribute
# ----------------------------------------
PH_ORIENT_HORZ = 'horz'
PH_ORIENT_VERT = 'vert'

# valid values for <p:ph> sz (size) attribute
# -------------------------------------------
PH_SZ_FULL = 'full'
PH_SZ_HALF = 'half'
PH_SZ_QUARTER = 'quarter'

# mapping of type to rootname (part of auto-generated placeholder shape name)
slide_ph_basenames = {
    PH_TYPE_TITLE: 'Title',
    # this next one is named 'Notes Placeholder' in a notes master
    PH_TYPE_BODY:     'Text Placeholder',
    PH_TYPE_CTRTITLE: 'Title',
    PH_TYPE_SUBTITLE: 'Subtitle',
    PH_TYPE_DT:       'Date Placeholder',
    PH_TYPE_SLDNUM:   'Slide Number Placeholder',
    PH_TYPE_FTR:      'Footer Placeholder',
    PH_TYPE_HDR:      'Header Placeholder',
    PH_TYPE_OBJ:      'Content Placeholder',
    PH_TYPE_CHART:    'Chart Placeholder',
    PH_TYPE_TBL:      'Table Placeholder',
    PH_TYPE_CLIPART:  'ClipArt Placeholder',
    PH_TYPE_DGM:      'SmartArt Placeholder',
    PH_TYPE_MEDIA:    'Media Placeholder',
    PH_TYPE_SLDIMG:   'Slide Image Placeholder',
    PH_TYPE_PIC:      'Picture Placeholder'}

# ============================================================================
# PresentationML Part Type specs
# ============================================================================
# Keyed by content type
# Not yet included:
# * Font Part (font1.fntdata) 15.2.13
# * themeOverride : 'application/vnd.openxmlformats-officedocument.'
#                   'themeOverride+xml'
# * several others, especially DrawingML parts ...
#
# TODO: Also check out other shared parts in section 15.
# ============================================================================

PTS_CARDINALITY_SINGLETON = 'singleton'
PTS_CARDINALITY_TUPLE = 'tuple'
PTS_HASRELS_ALWAYS = 'always'
PTS_HASRELS_NEVER = 'never'
PTS_HASRELS_OPTIONAL = 'optional'

CT_CHART = (
    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
CT_COMMENT_AUTHORS = (
    'application/vnd.openxmlformats-officedocument.presentationml.commentAuth'
    'ors+xml')
CT_COMMENTS = (
    'application/vnd.openxmlformats-officedocument.presentationml.comments+xm'
    'l')
CT_CORE_PROPS = (
    'application/vnd.openxmlformats-package.core-properties+xml')
CT_CUSTOM_PROPS = (
    'application/vnd.openxmlformats-officedocument.custom-properties+xml')
CT_HANDOUT_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.handoutMast'
    'er+xml')
CT_EXTENDED_PROPS = (
    'application/vnd.openxmlformats-officedocument.extended-properties+xml')
CT_NOTES_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.notesMaster'
    '+xml')
CT_NOTES_SLIDE = (
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+'
    'xml')
CT_PRESENTATION = (
    'application/vnd.openxmlformats-officedocument.presentationml.presentatio'
    'n.main+xml')
CT_PRES_PROPS = (
    'application/vnd.openxmlformats-officedocument.presentationml.presProps+x'
    'ml')
CT_PRINTER_SETTINGS = (
    'application/vnd.openxmlformats-officedocument.presentationml.printerSett'
    'ings')
CT_SLIDE = (
    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
CT_SLIDE_LAYOUT = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideLayout'
    '+xml')
CT_SLIDE_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideMaster'
    '+xml')
CT_SLIDESHOW = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow.m'
    'ain+xml')
CT_TABLE_STYLES = (
    'application/vnd.openxmlformats-officedocument.presentationml.tableStyles'
    '+xml')
CT_TAGS = (
    'application/vnd.openxmlformats-officedocument.presentationml.tags+xml')
CT_TEMPLATE = (
    'application/vnd.openxmlformats-officedocument.presentationml.template.ma'
    'in+xml')
CT_THEME = (
    'application/vnd.openxmlformats-officedocument.theme+xml')
CT_VIEW_PROPS = (
    'application/vnd.openxmlformats-officedocument.presentationml.viewProps+x'
    'ml')
CT_WORKSHEET = (
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


RT_CHART = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/char'
    't')
RT_COMMENTS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comm'
    'ents')
RT_COMMENT_AUTHORS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comm'
    'entAuthors')
RT_CORE_PROPS = (
    'http://schemas.openxmlformats.org/officedocument/2006/relationships/meta'
    'data/core-properties')
RT_CUSTOM_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/cust'
    'omProperties')
RT_EXTENDED_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/exte'
    'ndedProperties')
RT_HANDOUT_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hand'
    'outMaster')
RT_IMAGE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/imag'
    'e')
RT_NOTES_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/note'
    'sMaster')
RT_NOTES_SLIDE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/note'
    'sSlide')
RT_OFFICE_DOCUMENT = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/offi'
    'ceDocument')
RT_PACKAGE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pack'
    'age')
RT_PRES_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pres'
    'Props')
RT_PRINTER_SETTINGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/prin'
    'terSettings')
RT_SLIDE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'e')
RT_SLIDE_LAYOUT = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'eLayout')
RT_SLIDE_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'eMaster')
RT_TABLE_STYLES = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tabl'
    'eStyles')
RT_TAGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags')
RT_THEME = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/them'
    'e')
RT_VIEW_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/view'
    'Props')


pml_parttypes = {
    CT_COMMENT_AUTHORS: {  # ECMA-376-1 13.3.1
        'basename':    'commentAuthors',
        'ext':         '.xml',
        'name':        'Comment Authors Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_COMMENT_AUTHORS},
    CT_COMMENTS: {  # ECMA-376-1 13.3.2
        'basename':    'comment',
        'ext':         '.xml',
        'name':        'Comments Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/comments',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['slide'],
        'reltype':     RT_COMMENTS},
    CT_HANDOUT_MASTER: {  # ECMA-376-1 13.3.3
        'basename':    'handoutMaster',
        'ext':         '.xml',
        'name':        'Handout Master Part',
        # actually can only be one according to spec, but behaves like part
        # collection (handoutMasters folder, handoutMaster1.xml, etc.)
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/handoutMasters',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation'],
        'reltype':     RT_HANDOUT_MASTER},
    CT_NOTES_MASTER: {  # ECMA-376-1 13.3.4
        'basename':    'notesMaster',
        'ext':         '.xml',
        'name':        'Notes Master Part',
        # actually can only be one according to spec, but behaves like part
        # collection (handoutMasters folder, handoutMaster1.xml, etc.)
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/notesMasters',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'notesSlide'],
        'reltype':     RT_NOTES_MASTER},
    CT_NOTES_SLIDE: {  # ECMA-376-1 13.3.5
        'basename':    'notesSlide',
        'ext':         '.xml',
        'name':        'Notes Slide Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/notesSlides',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['slide'],
        'reltype':     RT_NOTES_SLIDE},
    CT_PRESENTATION: {  # ECMA-376-1 13.3.6
        # one of three possible Content Type values for presentation part
        'basename':    'presentation',
        'ext':         '.xml',
        'name':        'Presentation Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['package'],
        'reltype':     RT_OFFICE_DOCUMENT},
    CT_TEMPLATE: {  # ECMA-376-1 13.3.6
        # one of three possible Content Type values for presentation part
        'basename':    'presentation',
        'ext':         '.xml',
        'name':        'Presentation Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['package'],
        'reltype':     RT_OFFICE_DOCUMENT},
    CT_SLIDESHOW: {  # ECMA-376-1 13.3.6
        # one of three possible Content Type values for presentation part
        'basename':    'presentation',
        'ext':         '.xml',
        'name':        'Presentation Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['package'],
        'reltype':     RT_OFFICE_DOCUMENT},
    CT_PRES_PROPS: {  # ECMA-376-1 13.3.7
        'basename':    'presProps',
        'ext':         '.xml',
        'name':        'Presentation Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_PRES_PROPS},
    CT_SLIDE: {  # ECMA-376-1 13.3.8
        'basename':    'slide',
        'ext':         '.xml',
        'name':        'Slide Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/slides',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'notesSlide'],
        'reltype':     RT_SLIDE},
    CT_SLIDE_LAYOUT: {  # ECMA-376-1 13.3.9
        'basename':    'slideLayout',
        'ext':         '.xml',
        'name':        'Slide Layout Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    True,
        'baseURI':     '/ppt/slideLayouts',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['slide', 'slideMaster'],
        'reltype':     RT_SLIDE_LAYOUT},
    CT_SLIDE_MASTER: {  # ECMA-376-1 13.3.10
        'basename':    'slideMaster',
        'ext':         '.xml',
        'name':        'Slide Master Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    True,
        'baseURI':     '/ppt/slideMasters',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'slideLayout'],
        'reltype':     RT_SLIDE_MASTER},
    CT_TAGS: {  # ECMA-376-1 13.3.12
        'basename':    'tag',
        'ext':         '.xml',
        'name':        'User-Defined Tags Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/tags',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation', 'slide'],
        'reltype':     RT_TAGS},
    CT_VIEW_PROPS: {  # ECMA-376-1 13.3.13
        'basename':    'viewProps',
        'ext':         '.xml',
        'name':        'View Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_VIEW_PROPS},
    CT_CHART: {  # ECMA-376-1 14.2.1
        'basename':    'chart',
        'ext':         '.xml',
        'name':        'Chart Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/charts',
        # Spec states rels are optional. I'm not sure what kind of chart
        # doesn't have a relationship to a dataset (embedded xlsx package)
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['handoutMaster', 'notesMaster', 'notesSlide', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_CHART},
    CT_THEME: {  # ECMA-376-1 14.2.7
        'basename':    'theme',
        'ext':         '.xml',
        'name':        'Theme Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        # spec indicates theme part is optional, but I've never seen a .pptx
        # without one
        'required':    True,
        'baseURI':     '/ppt/theme',
        # can have _rels items, but only if the theme contains one or more
        # images
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['presentation', 'handoutMaster', 'notesMaster',
                        'slideMaster'],
        'reltype':     RT_THEME},
    CT_TABLE_STYLES: {  # ECMA-376-1 14.2.9
        'basename':    'tableStyles',
        'ext':         '.xml',
        'name':        'Table Styles Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_TABLE_STYLES},
    CT_WORKSHEET: {  # ECMA-376-1 15.2.11
        'basename':    'Microsoft_Excel_Sheet',
        'ext':         '.xlsx',
        'name':        'Embedded Package Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/embeddings',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['chart', 'handoutMaster', 'notesSlide', 'notesMaster',
                        'slide', 'slideLayout', 'slideMaster'],
        'reltype':     RT_PACKAGE},
    CT_CORE_PROPS: {  # ECMA-376-1 15.2.12.1
        'basename':    'core',
        'ext':         '.xml',
        # 'Core' as in Dublin Core
        'name':        'Core File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_CORE_PROPS},
    CT_CUSTOM_PROPS: {  # ECMA-376-1 15.2.12.2
        'basename':    'custom',
        'ext':         '.xml',
        'name':        'Custom File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_CUSTOM_PROPS},
    CT_EXTENDED_PROPS: {  # ECMA-376-1 15.2.12.3
        'basename':    'app',
        'ext':         '.xml',
        'name':        'Extended File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_EXTENDED_PROPS},
    'image/gif': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.gif',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE},
    'image/jpeg': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.jpeg',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE},
    'image/png': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.png',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE},
    'image/x-emf': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.emf',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE},
    CT_PRINTER_SETTINGS: {  # ECMA-376-1 15.2.15
        'basename':    'printerSettings',
        'ext':         '.bin',
        'name':        'Printer Settings Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/printerSettings',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_PRINTER_SETTINGS}}


# ============================================================================
# default_content_types
# ============================================================================
# Default file extension to MIME type mapping. This is used as a reference for
# adding <Default> elements to [Content_Types].xml for parts like media files.
#     TODO: I've seen .wmv elements in the media folder of at least one
# presentation, might need to add an entry for that and perhaps other rich
# media PowerPoint allows to be embedded (e.g. audio, movie, object, ...).
# ============================================================================

default_content_types = {
    '.bin':     CT_PRINTER_SETTINGS,
    '.emf':     'image/x-emf',
    '.fntdata': 'application/x-fontdata',
    '.gif':     'image/gif',
    '.jpe':     'image/jpeg',
    '.jpeg':    'image/jpeg',
    '.jpg':     'image/jpeg',
    '.png':     'image/png',
    '.rels':    'application/vnd.openxmlformats-package.relationships+xml',
    '.tif':     'image/tiff',
    '.tiff':    'image/tiff',
    '.wmf':     'image/x-wmf',
    '.xlsx':    CT_WORKSHEET,
    '.xml':     'application/xml'}


# ============================================================================
# nsmap
# ============================================================================
# namespace prefix to namespace name map
# ============================================================================

nsmap = {
    'a':   ('http://schemas.openxmlformats.org/drawingml/2006/main'),
    'cp':  ('http://schemas.openxmlformats.org/package/2006/metadata/core-pro'
            'perties'),
    'ct':  ('http://schemas.openxmlformats.org/package/2006/content-types'),
    'dc':  ('http://purl.org/dc/elements/1.1/'),
    'dcmitype': ('http://purl.org/dc/dcmitype/'),
    'dcterms':  ('http://purl.org/dc/terms/'),
    'ep':  ('http://schemas.openxmlformats.org/officeDocument/2006/extended-p'
            'roperties'),
    'mv':  ('urn:schemas-microsoft-com:mac:vml'),
    'mo':  ('http://schemas.microsoft.com/office/mac/office/2008/main'),
    'm':   ('http://schemas.openxmlformats.org/officeDocument/2006/math'),
    'o':   ('urn:schemas-microsoft-com:office:office'),
    'p':   ('http://schemas.openxmlformats.org/presentationml/2006/main'),
    'pd':  ('http://schemas.openxmlformats.org/drawingml/2006/presentationDra'
            'wing'),
    'pic': ('http://schemas.openxmlformats.org/drawingml/2006/picture'),
    'pr':  ('http://schemas.openxmlformats.org/package/2006/relationships'),
    'r':   ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
            'ips'),
    'sl':  ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
            'ips/slideLayout'),
    'v':   ('urn:schemas-microsoft-com:vml'),
    've':  ('http://schemas.openxmlformats.org/markup-compatibility/2006'),
    'w':   ('http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
    'w10': ('urn:schemas-microsoft-com:office:word'),
    'wne': ('http://schemas.microsoft.com/office/word/2006/wordml'),
    'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingD'
            'rawing'),
    'xsi': ('http://www.w3.org/2001/XMLSchema-instance')}


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


def qtag(tag):
    """
    Return a qualified name (QName) for an XML element or attribute in Clark
    notation, e.g. ``'{http://www.w3.org/1999/xhtml}body'`` instead of
    ``'html:body'``, by looking up the specified namespace prefix in the
    overall namespace map (nsmap) above. Google on "xml clark notation" for
    more on Clark notation. *tag* is a namespace-prefixed tagname, e.g.
    ``'p:cSld'``.
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{%s}%s' % (uri, tagroot)
