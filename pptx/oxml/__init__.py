# encoding: utf-8

"""
Initializes lxml parser and makes available a handful of functions that wrap
its typical uses.
"""

from __future__ import absolute_import

from lxml import etree, objectify

from pptx.oxml.ns import NamespacePrefixedTag


# oxml-specific constants
XSD_TRUE = '1'


# configure objectified XML parser
_fallback_lookup = objectify.ObjectifyElementClassLookup()
_element_class_lookup = etree.ElementNamespaceClassLookup(_fallback_lookup)
oxml_parser = etree.XMLParser(remove_blank_text=True)
oxml_parser.set_element_class_lookup(_element_class_lookup)


def parse_xml_bytes(xml_bytes):
    """
    Return root lxml element obtained by parsing XML contained in *xml_bytes*.
    """
    root_element = objectify.fromstring(xml_bytes, oxml_parser)
    return root_element


def register_custom_element_class(nsptag_str, cls):
    """
    Register the lxml custom element class *cls* with the parser to be used
    for XML elements with tag matching *nsptag_str*.
    """
    nsptag = NamespacePrefixedTag(nsptag_str)
    namespace = _element_class_lookup.get_namespace(nsptag.nsuri)
    namespace[nsptag.local_part] = cls


from pptx.oxml.autoshape import (
    CT_PresetGeometry2D, CT_Shape, CT_ShapeProperties
)
register_custom_element_class('a:prstGeom', CT_PresetGeometry2D)
register_custom_element_class('p:sp',       CT_Shape)
register_custom_element_class('p:spPr',     CT_ShapeProperties)


from pptx.oxml.dml.color import (
    CT_HslColor, CT_Percentage, CT_PresetColor, CT_SchemeColor,
    CT_ScRgbColor, CT_SRgbColor, CT_SystemColor
)
register_custom_element_class('a:hslClr',    CT_HslColor)
register_custom_element_class('a:lumMod',    CT_Percentage)
register_custom_element_class('a:lumOff',    CT_Percentage)
register_custom_element_class('a:prstClr',   CT_PresetColor)
register_custom_element_class('a:schemeClr', CT_SchemeColor)
register_custom_element_class('a:scrgbClr',  CT_ScRgbColor)
register_custom_element_class('a:srgbClr',   CT_SRgbColor)
register_custom_element_class('a:sysClr',    CT_SystemColor)


from pptx.oxml.dml.fill import (
    CT_BlipFillProperties, CT_GradientFillProperties, CT_GroupFillProperties,
    CT_NoFillProperties, CT_PatternFillProperties,
    CT_SolidColorFillProperties,
)
register_custom_element_class('a:blipFill',  CT_BlipFillProperties)
register_custom_element_class('a:gradFill',  CT_GradientFillProperties)
register_custom_element_class('a:grpFill',   CT_GroupFillProperties)
register_custom_element_class('a:noFill',    CT_NoFillProperties)
register_custom_element_class('a:pattFill',  CT_PatternFillProperties)
register_custom_element_class('a:solidFill', CT_SolidColorFillProperties)


from pptx.oxml.presentation import (
    CT_Presentation, CT_SlideId, CT_SlideIdList
)
register_custom_element_class('p:presentation', CT_Presentation)
register_custom_element_class('p:sldId',        CT_SlideId)
register_custom_element_class('p:sldIdLst',     CT_SlideIdList)

from pptx.oxml.text import (
    CT_Hyperlink, CT_RegularTextRun, CT_TextBody, CT_TextBodyProperties,
    CT_TextCharacterProperties, CT_TextParagraph, CT_TextParagraphProperties
)
register_custom_element_class('a:bodyPr',     CT_TextBodyProperties)
register_custom_element_class('a:defRPr',     CT_TextCharacterProperties)
register_custom_element_class('a:endParaRPr', CT_TextCharacterProperties)
register_custom_element_class('a:hlinkClick', CT_Hyperlink)
register_custom_element_class('a:r',          CT_RegularTextRun)
register_custom_element_class('a:p',          CT_TextParagraph)
register_custom_element_class('a:pPr',        CT_TextParagraphProperties)
register_custom_element_class('a:rPr',        CT_TextCharacterProperties)
register_custom_element_class('p:txBody',     CT_TextBody)
