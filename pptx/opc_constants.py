# -*- coding: utf-8 -*-
#
# opc_constants.py
#
# Copyright (C) 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under the MIT License:
# http://www.opensource.org/licenses/mit-license.php

"""
Constant values related to the Open Packaging Convention, in particular,
content types and relationship types.
"""


class CONTENT_TYPE(object):
    """
    Content type URIs (like MIME-types) that specify a part's format
    """
    BMP = (
        'image/bmp'
    )
    DML_CHART = (
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    )
    DML_CHARTSHAPES = (
        'application/vnd.openxmlformats-officedocument.drawingml.chartshapes'
        '+xml'
    )
    DML_DIAGRAM_COLORS = (
        'application/vnd.openxmlformats-officedocument.drawingml.diagramColo'
        'rs+xml'
    )
    DML_DIAGRAM_DATA = (
        'application/vnd.openxmlformats-officedocument.drawingml.diagramData'
        '+xml'
    )
    DML_DIAGRAM_LAYOUT = (
        'application/vnd.openxmlformats-officedocument.drawingml.diagramLayo'
        'ut+xml'
    )
    DML_DIAGRAM_STYLE = (
        'application/vnd.openxmlformats-officedocument.drawingml.diagramStyl'
        'e+xml'
    )
    GIF = (
        'image/gif'
    )
    JPEG = (
        'image/jpeg'
    )
    MS_PHOTO = (
        'image/vnd.ms-photo'
    )
    OFC_CUSTOM_PROPERTIES = (
        'application/vnd.openxmlformats-officedocument.custom-properties+xml'
    )
    OFC_CUSTOM_XML_PROPERTIES = (
        'application/vnd.openxmlformats-officedocument.customXmlProperties+x'
        'ml'
    )
    OFC_DRAWING = (
        'application/vnd.openxmlformats-officedocument.drawing+xml'
    )
    OFC_EXTENDED_PROPERTIES = (
        'application/vnd.openxmlformats-officedocument.extended-properties+x'
        'ml'
    )
    OFC_OLE_OBJECT = (
        'application/vnd.openxmlformats-officedocument.oleObject'
    )
    OFC_PACKAGE = (
        'application/vnd.openxmlformats-officedocument.package'
    )
    OFC_THEME = (
        'application/vnd.openxmlformats-officedocument.theme+xml'
    )
    OFC_THEME_OVERRIDE = (
        'application/vnd.openxmlformats-officedocument.themeOverride+xml'
    )
    OFC_VML_DRAWING = (
        'application/vnd.openxmlformats-officedocument.vmlDrawing'
    )
    OPC_CORE_PROPERTIES = (
        'application/vnd.openxmlformats-package.core-properties+xml'
    )
    OPC_DIGITAL_SIGNATURE_CERTIFICATE = (
        'application/vnd.openxmlformats-package.digital-signature-certificat'
        'e'
    )
    OPC_DIGITAL_SIGNATURE_ORIGIN = (
        'application/vnd.openxmlformats-package.digital-signature-origin'
    )
    OPC_DIGITAL_SIGNATURE_XMLSIGNATURE = (
        'application/vnd.openxmlformats-package.digital-signature-xmlsignatu'
        're+xml'
    )
    OPC_RELATIONSHIPS = (
        'application/vnd.openxmlformats-package.relationships+xml'
    )
    PML_COMMENTS = (
        'application/vnd.openxmlformats-officedocument.presentationml.commen'
        'ts+xml'
    )
    PML_COMMENT_AUTHORS = (
        'application/vnd.openxmlformats-officedocument.presentationml.commen'
        'tAuthors+xml'
    )
    PML_HANDOUT_MASTER = (
        'application/vnd.openxmlformats-officedocument.presentationml.handou'
        'tMaster+xml'
    )
    PML_NOTES_MASTER = (
        'application/vnd.openxmlformats-officedocument.presentationml.notesM'
        'aster+xml'
    )
    PML_NOTES_SLIDE = (
        'application/vnd.openxmlformats-officedocument.presentationml.notesS'
        'lide+xml'
    )
    PML_PRESENTATION_MAIN = (
        'application/vnd.openxmlformats-officedocument.presentationml.presen'
        'tation.main+xml'
    )
    PML_PRES_PROPS = (
        'application/vnd.openxmlformats-officedocument.presentationml.presPr'
        'ops+xml'
    )
    PML_PRINTER_SETTINGS = (
        'application/vnd.openxmlformats-officedocument.presentationml.printe'
        'rSettings'
    )
    PML_SLIDE = (
        'application/vnd.openxmlformats-officedocument.presentationml.slide+'
        'xml'
    )
    PML_SLIDESHOW_MAIN = (
        'application/vnd.openxmlformats-officedocument.presentationml.slides'
        'how.main+xml'
    )
    PML_SLIDE_LAYOUT = (
        'application/vnd.openxmlformats-officedocument.presentationml.slideL'
        'ayout+xml'
    )
    PML_SLIDE_MASTER = (
        'application/vnd.openxmlformats-officedocument.presentationml.slideM'
        'aster+xml'
    )
    PML_SLIDE_UPDATE_INFO = (
        'application/vnd.openxmlformats-officedocument.presentationml.slideU'
        'pdateInfo+xml'
    )
    PML_TABLE_STYLES = (
        'application/vnd.openxmlformats-officedocument.presentationml.tableS'
        'tyles+xml'
    )
    PML_TAGS = (
        'application/vnd.openxmlformats-officedocument.presentationml.tags+x'
        'ml'
    )
    PML_TEMPLATE_MAIN = (
        'application/vnd.openxmlformats-officedocument.presentationml.templa'
        'te.main+xml'
    )
    PML_VIEW_PROPS = (
        'application/vnd.openxmlformats-officedocument.presentationml.viewPr'
        'ops+xml'
    )
    PNG = (
        'image/png'
    )
    SML_CALC_CHAIN = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.calcCha'
        'in+xml'
    )
    SML_CHARTSHEET = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsh'
        'eet+xml'
    )
    SML_COMMENTS = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.comment'
        's+xml'
    )
    SML_CONNECTIONS = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.connect'
        'ions+xml'
    )
    SML_CUSTOM_PROPERTY = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.customP'
        'roperty'
    )
    SML_DIALOGSHEET = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.dialogs'
        'heet+xml'
    )
    SML_EXTERNAL_LINK = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.externa'
        'lLink+xml'
    )
    SML_PIVOT_CACHE_DEFINITION = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCa'
        'cheDefinition+xml'
    )
    SML_PIVOT_CACHE_RECORDS = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCa'
        'cheRecords+xml'
    )
    SML_PIVOT_TABLE = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTa'
        'ble+xml'
    )
    SML_PRINTER_SETTINGS = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.printer'
        'Settings'
    )
    SML_QUERY_TABLE = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTa'
        'ble+xml'
    )
    SML_REVISION_HEADERS = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.revisio'
        'nHeaders+xml'
    )
    SML_REVISION_LOG = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.revisio'
        'nLog+xml'
    )
    SML_SHARED_STRINGS = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedS'
        'trings+xml'
    )
    SML_SHEET = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    SML_SHEET_MAIN = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.m'
        'ain+xml'
    )
    SML_SHEET_METADATA = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMe'
        'tadata+xml'
    )
    SML_STYLES = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+'
        'xml'
    )
    SML_TABLE = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.table+x'
        'ml'
    )
    SML_TABLE_SINGLE_CELLS = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.tableSi'
        'ngleCells+xml'
    )
    SML_TEMPLATE_MAIN = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.templat'
        'e.main+xml'
    )
    SML_USER_NAMES = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.userNam'
        'es+xml'
    )
    SML_VOLATILE_DEPENDENCIES = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.volatil'
        'eDependencies+xml'
    )
    SML_WORKSHEET = (
        'application/vnd.openxmlformats-officedocument.spreadsheetml.workshe'
        'et+xml'
    )
    TIFF = (
        'image/tiff'
    )
    WML_COMMENTS = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.comm'
        'ents+xml'
    )
    WML_DOCUMENT_GLOSSARY = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.docu'
        'ment.glossary+xml'
    )
    WML_DOCUMENT_MAIN = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.docu'
        'ment.main+xml'
    )
    WML_ENDNOTES = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.endn'
        'otes+xml'
    )
    WML_FONT_TABLE = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.font'
        'Table+xml'
    )
    WML_FOOTER = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.foot'
        'er+xml'
    )
    WML_FOOTNOTES = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.foot'
        'notes+xml'
    )
    WML_HEADER = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.head'
        'er+xml'
    )
    WML_NUMBERING = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.numb'
        'ering+xml'
    )
    WML_PRINTER_SETTINGS = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.prin'
        'terSettings'
    )
    WML_SETTINGS = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.sett'
        'ings+xml'
    )
    WML_STYLES = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.styl'
        'es+xml'
    )
    WML_WEB_SETTINGS = (
        'application/vnd.openxmlformats-officedocument.wordprocessingml.webS'
        'ettings+xml'
    )
    XML = (
        'application/xml'
    )
    X_EMF = (
        'image/x-emf'
    )
    X_FONTDATA = (
        'application/x-fontdata'
    )
    X_FONT_TTF = (
        'application/x-font-ttf'
    )
    X_WMF = (
        'image/x-wmf'
    )
