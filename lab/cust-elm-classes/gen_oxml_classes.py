#!/usr/bin/env python
# -*- coding: utf-8 -*-

# gen_oxml_classes.py
#

"""
Generate class definitions for Open XML elements based on set of declarative
properties.
"""

import sys
from string import Template

class ElementDef(object):
    """
    Schema-related definition of an Open XML element
    """
    instances = []

    def __init__(self, tag, classname):
        self.instances.append(self)
        self.tag = tag
        tagparts = tag.split(':')
        self.nsprefix = tagparts[0]
        self.tagname = tagparts[1]
        self.classname = classname
        self.children = []
        self.attributes = []

    def __getitem__(self, key):
        return self.__getattribute__(key)

    def add_child(self, tag, cardinality='?'):
        self.children.append(ChildDef(self, tag, cardinality))

    def add_attribute(self, name, required=False, default=None):
        self.attributes.append(AttributeDef(self, name, required, default))

    def add_attributes(self, *names):
        for name in names:
            self.attributes.append(AttributeDef(self, name, False, None))

    @property
    def indent(self):
        indent_len = 12 - len(self.tagname)
        if indent_len < 0:
            indent_len = 0
        return ' ' * indent_len

    @property
    def max_attr_name_len(self):
        return max([len(attr.name) for attr in self.attributes])

    @property
    def max_child_tagname_len(self):
        return max([len(child.tagname) for child in self.children])

    @property
    def optional_children(self):
        return [child for child in self.children if not child.is_required]

    @property
    def required_attributes(self):
        return [attr for attr in self.attributes if attr.is_required]

    @property
    def required_children(self):
        return [child for child in self.children if child.is_required]


class AttributeDef(object):
    """
    Attribute definition
    """
    def __init__(self, element, name, required, default):
        self.element = element
        self.name = name
        self.required = required
        self.default = default
        self.varname = name.replace(':', '_')

    def __getitem__(self, key):
        return self.__getattribute__(key)

    @property
    def padding(self):
        return ' ' * (self.element.max_attr_name_len - len(self.name))

    @property
    def indent(self):
        return ' ' * self.element.max_attr_name_len

    @property
    def is_required(self):
        return self.required


class ChildDef(object):
    """
    Child element definition
    """
    def __init__(self, element, tag, cardinality):
        self.element = element
        self.tag = tag
        self.cardinality = cardinality
        if not ':' in tag:
            tmpl = "missing namespace prefix in tag: '%s'"
            raise ValueError(tmpl % tag)
        tagparts = tag.split(':')
        self.nsprefix = tagparts[0]
        self.tagname = tagparts[1]

    def __getitem__(self, key):
        return self.__getattribute__(key)

    @property
    def indent(self):
        indent_len = self.element.max_child_tagname_len - len(self.tagname)
        return ' ' * indent_len

    @property
    def is_required(self):
        return self.cardinality in '1+'


# ============================================================================
# Code templates
# ============================================================================

class tmpl(object):
    attr_accessor = Template("    $varname$padding = property(lambda self: se"
        "lf.get('$name'),\n$indent                lambda self, value: self.se"
        "t('$name', value))\n")

    class_def_head = Template('''class $classname(ElementBase):\n    """<$n'''
        '''sprefix:$tagname> custom element class"""\n''')

    class_mapping = Template("        , '$tagname'$indent : $classname\n\n")

    ns_attr_accessor = Template(
        "    $varname$padding = property(lambda self: self.get(_qtag('$name')"
        "),\n$indent                lambda self, value: self.set(_qtag('$name"
        "'), value))\n")

    ns_reqd_attribute_constructor = Template("        _required_attribute(sel"
        "f, _qtag('$name'), default='$default')\n")

    optional_child_accessor = Template("    $tagname$indent = property(lambda"
        " self: _get_child_or_append(self, '$tag'))\n")

    reqd_attr_constructor = Template("        _required_attribute(self, '$nam"
        "e', default='$default')\n")

    reqd_child_accessor = Template("    $tagname$indent = property(lambda sel"
        "f: _child(self, '$tag'))\n")

    reqd_child_constructor = Template("        _required_child(self, '$tag')"
        "\n")


# ============================================================================
# binding specs
# ============================================================================

# sldMaster = ElementDef('p:sldMaster', 'CT_SlideMaster')
# sldMaster.add_child('p:cSld'          , cardinality='1')
# sldMaster.add_child('p:clrMap'        , cardinality='1')
# sldMaster.add_child('p:sldLayoutIdLst', cardinality='?')
# sldMaster.add_child('p:transition'    , cardinality='?')
# sldMaster.add_child('p:timing'        , cardinality='?')
# sldMaster.add_child('p:hf'            , cardinality='?')
# sldMaster.add_child('p:txStyles'      , cardinality='?')
# sldMaster.add_child('p:extLst'        , cardinality='?')
# sldMaster.add_attributes('preserve')



def class_template(element):
    out = ''
    out += tmpl.class_mapping.substitute(element)
    out += tmpl.class_def_head.substitute(element)
    if element.required_children or element.required_attributes:
        out += '    def _init(self):\n'
        for child in element.required_children:
            out += tmpl.reqd_child_constructor.substitute(child)
        for attribute in element.required_attributes:
            if ':' in attribute.name:
                out += tmpl.ns_reqd_attr_constructor.substitute(attribute)
            else:
                out += tmpl.reqd_attr_constructor.substitute(attribute)
        out += '\n'
    if element.children:
        out += '    # child accessors -----------------\n'
        for child in element.required_children:
            out += tmpl.reqd_child_accessor.substitute(child)
        for child in element.optional_children:
            out += tmpl.optional_child_accessor.substitute(child)
        out += '\n'
    if element.attributes:
        out += '    # attribute accessors -------------\n'
        for attribute in element.attributes:
            if ':' in attribute.name:
                out += tmpl.ns_attr_accessor.substitute(attribute)
            else:
                out += tmpl.attr_accessor.substitute(attribute)
        out += '\n'
    out += '\n'
    return out



elements = ElementDef.instances

out = '\n'
for element in elements:
    out += class_template(element)

print out

sys.exit()


# ============================================================================
# Element definitions
# ============================================================================

# blip = ElementDef('a:blip', 'CT_Blip')
# blip.add_child('a:alphaBiLevel' , cardinality='?')
# blip.add_child('a:alphaCeiling' , cardinality='?')
# blip.add_child('a:alphaFloor'   , cardinality='?')
# blip.add_child('a:alphaInv'     , cardinality='?')
# blip.add_child('a:alphaMod'     , cardinality='?')
# blip.add_child('a:alphaModFix'  , cardinality='?')
# blip.add_child('a:alphaRepl'    , cardinality='?')
# blip.add_child('a:biLevel'      , cardinality='?')
# blip.add_child('a:blur'         , cardinality='?')
# blip.add_child('a:clrChange'    , cardinality='?')
# blip.add_child('a:clrRepl'      , cardinality='?')
# blip.add_child('a:duotone'      , cardinality='?')
# blip.add_child('a:fillOverlay'  , cardinality='?')
# blip.add_child('a:grayscl'      , cardinality='?')
# blip.add_child('a:hsl'          , cardinality='?')
# blip.add_child('a:lum'          , cardinality='?')
# blip.add_child('a:tint'         , cardinality='?')
# blip.add_child('a:extLst'       , cardinality='?')
# blip.add_attributes('r_embed', 'r_link', 'cstate')

# blipFill = ElementDef('p:blipFill', 'CT_BlipFillProperties')
# blipFill.add_child('a:blip'    , cardinality='?')
# blipFill.add_child('a:srcRect' , cardinality='?')
# blipFill.add_child('a:tile'    , cardinality='?')
# blipFill.add_child('a:stretch' , cardinality='?')
# blipFill.add_attributes('dpi', 'rotWithShape')

# bodyPr = ElementDef('a:bodyPr', 'CT_TextBodyProperties')
# bodyPr.add_child('a:spAutoFit', cardinality='?')
# bodyPr.add_attribute('wrap')
# bodyPr.add_attribute('rtlCol')
# bodyPr.add_attribute('anchor')
# bodyPr.add_attribute('anchorCtr')

# cNvPr = ElementDef('p:cNvPr', 'CT_NonVisualDrawingProps')
# cNvPr.add_child('a:hlinkClick' , cardinality='?')
# cNvPr.add_child('a:hlinkHover' , cardinality='?')
# cNvPr.add_child('a:extLst'     , cardinality='?')
# cNvPr.add_attribute('id', required=True, default='0')
# cNvPr.add_attribute('name', required=True, default='Unnamed')
# cNvPr.add_attributes('descr', 'hidden', 'title')

# cNvSpPr = ElementDef('p:cNvSpPr', 'CT_NonVisualDrawingShapeProps')
# cNvSpPr.add_child('a:spLocks', cardinality='?')
# cNvSpPr.add_child('a:extLst' , cardinality='?')
# cNvSpPr.add_attributes('txBox')

# cSld = ElementDef('p:cSld', 'CT_CommonSlideData')
# cSld.add_child('p:bg'         , cardinality='?')
# cSld.add_child('p:spTree'     , cardinality='1')
# cSld.add_child('p:custDataLst', cardinality='?')
# cSld.add_child('p:controls'   , cardinality='?')
# cSld.add_child('p:extLst'     , cardinality='?')
# cSld.add_attributes('name')

# defRPr = ElementDef('a:defRPr', 'CT_TextCharacterProperties')
# defRPr.add_child('a:ln'            , cardinality='?')
# defRPr.add_child('a:noFill'        , cardinality='?')
# defRPr.add_child('a:solidFill'     , cardinality='?')
# defRPr.add_child('a:gradFill'      , cardinality='?')
# defRPr.add_child('a:blipFill'      , cardinality='?')
# defRPr.add_child('a:pattFill'      , cardinality='?')
# defRPr.add_child('a:grpFill'       , cardinality='?')
# defRPr.add_child('a:effectLst'     , cardinality='?')
# defRPr.add_child('a:effectDag'     , cardinality='?')
# defRPr.add_child('a:highlight'     , cardinality='?')
# defRPr.add_child('a:uLnTx'         , cardinality='?')
# defRPr.add_child('a:uLn'           , cardinality='?')
# defRPr.add_child('a:uFillTx'       , cardinality='?')
# defRPr.add_child('a:uFill'         , cardinality='?')
# defRPr.add_child('a:latin'         , cardinality='?')
# defRPr.add_child('a:ea'            , cardinality='?')
# defRPr.add_child('a:cs'            , cardinality='?')
# defRPr.add_child('a:sym'           , cardinality='?')
# defRPr.add_child('a:hlinkClick'    , cardinality='?')
# defRPr.add_child('a:hlinkMouseOver', cardinality='?')
# defRPr.add_child('a:rtl'           , cardinality='?')
# defRPr.add_child('a:extLst'        , cardinality='?')
# defRPr.add_attributes('kumimoji', 'lang', 'altLang', 'sz', 'b', 'i', 'u',
#     'strike', 'kern', 'cap', 'spc', 'normalizeH', 'baseline', 'noProof',
#     'dirty', 'err', 'smtClean', 'smtId', 'bmk')

# nvGrpSpPr = ElementDef('p:nvGrpSpPr', 'CT_GroupShapeNonVisual')
# nvGrpSpPr.add_child('p:cNvPr'     , cardinality='1')
# nvGrpSpPr.add_child('p:cNvGrpSpPr', cardinality='1')
# nvGrpSpPr.add_child('p:nvPr'      , cardinality='1')

# nvPicPr = ElementDef('p:nvPicPr', 'CT_PictureNonVisual')
# nvPicPr.add_child('p:cNvPr'    , cardinality='1')
# nvPicPr.add_child('p:cNvPicPr' , cardinality='1')
# nvPicPr.add_child('p:nvPr'     , cardinality='1')

# nvPr = ElementDef('p:nvPr', 'CT_ApplicationNonVisualDrawingProps')
# nvPr.add_child('p:ph'           , cardinality='?')
# nvPr.add_child('p:audioCd'      , cardinality='?')
# nvPr.add_child('p:wavAudioFile' , cardinality='?')
# nvPr.add_child('p:audioFile'    , cardinality='?')
# nvPr.add_child('p:videoFile'    , cardinality='?')
# nvPr.add_child('p:quickTimeFile', cardinality='?')
# nvPr.add_child('p:custDataLst'  , cardinality='?')
# nvPr.add_child('p:extLst'       , cardinality='?')
# nvPr.add_attributes('isPhoto', 'userDrawn')

# nvSpPr = ElementDef('p:nvSpPr', 'CT_ShapeNonVisual')
# nvSpPr.add_child('p:cNvPr'  , cardinality='1')
# nvSpPr.add_child('p:cNvSpPr', cardinality='1')
# nvSpPr.add_child('p:nvPr'   , cardinality='1')

# off = ElementDef('a:off', 'CT_Point2D')
# off.add_attribute('x', required=True, default='0')
# off.add_attribute('y', required=True, default='0')

# p = ElementDef('a:p', 'CT_TextParagraph')
# p.add_child('a:pPr' , cardinality='?')
# p.add_child('a:r'   , cardinality='*')
# p.add_child('a:br'  , cardinality='*')
# p.add_child('a:fld' , cardinality='*')
# p.add_child('a:endParaRPr', cardinality='?')

# ph = ElementDef('p:ph', 'CT_Placeholder')
# ph.add_child('p:extLst', cardinality='?')
# ph.add_attributes('type', 'orient', 'sz', 'idx', 'hasCustomPrompt')

# pic = ElementDef('p:pic', 'CT_Picture')
# pic.add_child('p:nvPicPr'  , cardinality='1')
# pic.add_child('p:blipFill' , cardinality='1')
# pic.add_child('p:spPr'     , cardinality='1')
# pic.add_child('p:style'    , cardinality='?')
# pic.add_child('p:extLst'   , cardinality='?')

# pPr = ElementDef('a:pPr', 'CT_TextParagraphProperties')
# pPr.add_child('a:lnSpc'    , cardinality='?')
# pPr.add_child('a:spcBef'   , cardinality='?')
# pPr.add_child('a:spcAft'   , cardinality='?')
# pPr.add_child('a:buClrTx'  , cardinality='?')
# pPr.add_child('a:buClr'    , cardinality='?')
# pPr.add_child('a:buSzTx'   , cardinality='?')
# pPr.add_child('a:buSzPct'  , cardinality='?')
# pPr.add_child('a:buSzPts'  , cardinality='?')
# pPr.add_child('a:buFontTx' , cardinality='?')
# pPr.add_child('a:buFont'   , cardinality='?')
# pPr.add_child('a:buNone'   , cardinality='?')
# pPr.add_child('a:buAutoNum', cardinality='?')
# pPr.add_child('a:buChar'   , cardinality='?')
# pPr.add_child('a:buBlip'   , cardinality='?')
# pPr.add_child('a:tabLst'   , cardinality='?')
# pPr.add_child('a:defRPr'   , cardinality='?')
# pPr.add_child('a:extLst'   , cardinality='?')
# pPr.add_attributes('marL', 'marR', 'lvl', 'indent', 'algn', 'defTabSz',
#     'rtl', 'eaLnBrk', 'fontAlgn', 'latinLnBrk', 'hangingPunct')

# presentation = ElementDef('p:presentation', 'CT_Presentation')
# presentation.add_child('p:sldMasterIdLst'    , cardinality='?')
# presentation.add_child('p:notesMasterIdLst'  , cardinality='?')
# presentation.add_child('p:handoutMasterIdLst', cardinality='?')
# presentation.add_child('p:sldIdLst'          , cardinality='?')
# presentation.add_child('p:sldSz'             , cardinality='?')
# presentation.add_child('p:notesSz'           , cardinality='1')
# presentation.add_child('p:smartTags'         , cardinality='?')
# presentation.add_child('p:embeddedFontLst'   , cardinality='?')
# presentation.add_child('p:custShowLst'       , cardinality='?')
# presentation.add_child('p:photoAlbum'        , cardinality='?')
# presentation.add_child('p:custDataLst'       , cardinality='?')
# presentation.add_child('p:kinsoku'           , cardinality='?')
# presentation.add_child('p:defaultTextStyle'  , cardinality='?')
# presentation.add_child('p:modifyVerifier'    , cardinality='?')
# presentation.add_child('p:extLst'            , cardinality='?')
# presentation.add_attributes('serverZoom', 'firstSlideNum',
#     'showSpecialPlsOnTitleSld', 'rtl', 'removePersonalInfoOnSave',
#     'compatMode', 'strictFirstAndLastChars', 'embedTrueTypeFonts',
#     'saveSubsetFonts', 'autoCompressPictures', 'bookmarkIdSeed',
#     'conformance')

# rPr = ElementDef('a:rPr', 'CT_TextCharacterProperties')
# rPr.add_child('a:ln', cardinality='?')
# rPr.add_attribute('sz')
# rPr.add_attribute('b')
# rPr.add_attribute('i')
# rPr.add_attribute('u')
# rPr.add_attribute('strike')
# rPr.add_attribute('kern')
# rPr.add_attribute('cap')
# rPr.add_attribute('spc')
# rPr.add_attribute('baseline')

# sld = ElementDef('p:sld', 'CT_Slide')
# sld.add_child('p:cSld'       , cardinality='1')
# sld.add_child('p:clrMapOvr'  , cardinality='?')
# sld.add_child('p:transition' , cardinality='?')
# sld.add_child('p:timing'     , cardinality='?')
# sld.add_child('p:extLst'     , cardinality='?')
# sld.add_attributes('showMasterSp', 'showMasterPhAnim', 'show')

# sldId = ElementDef('p:sldId', 'CT_SlideIdListEntry')
# sldId.add_child('p:extLst', cardinality='?')
# sldId.add_attribute('id', required=True, default='')
# sldId.add_attribute('r:id', required=True, default='')

# sldIdLst = ElementDef('p:sldIdLst', 'CT_SlideIdList')
# sldIdLst.add_child('p:sldId', cardinality='*')

# sldLayout = ElementDef('p:sldLayout', 'CT_SlideLayout')
# sldLayout.add_child('p:cSld'      , cardinality='1')
# sldLayout.add_child('p:clrMapOvr' , cardinality='?')
# sldLayout.add_child('p:transition', cardinality='?')
# sldLayout.add_child('p:timing'    , cardinality='?')
# sldLayout.add_child('p:hf'        , cardinality='?')
# sldLayout.add_child('p:extLst'    , cardinality='?')
# sldLayout.add_attributes('showMasterSp', 'showMasterPhAnim', 'matchingName', 'type', 'preserve', 'userDrawn')

# spLocks = ElementDef('a:spLocks', 'CT_ShapeLocking')
# spLocks.add_child('a:extLst', cardinality='?')
# spLocks.add_attributes('noGrp', 'noSelect', 'noRot', 'noChangeAspect',
#     'noMove', 'noResize', 'noEditPoints', 'noAdjustHandles',
#     'noChangeArrowheads', 'noChangeShapeType', 'noTextEdit')

# spPr = ElementDef('p:spPr', 'CT_ShapeProperties')
# spPr.add_child('a:xfrm'      , cardinality='?')
# spPr.add_child('a:custGeom'  , cardinality='?')
# spPr.add_child('a:prstGeom'  , cardinality='?')
# spPr.add_child('a:noFill'    , cardinality='?')
# spPr.add_child('a:solidFill' , cardinality='?')
# spPr.add_child('a:gradFill'  , cardinality='?')
# spPr.add_child('a:blipFill'  , cardinality='?')
# spPr.add_child('a:pattFill'  , cardinality='?')
# spPr.add_child('a:grpFill'   , cardinality='?')
# spPr.add_child('a:ln'        , cardinality='?')
# spPr.add_child('a:effectLst' , cardinality='?')
# spPr.add_child('a:effectDag' , cardinality='?')
# spPr.add_child('a:scene3d'   , cardinality='?')
# spPr.add_child('a:sp3d'      , cardinality='?')
# spPr.add_child('a:extLst'    , cardinality='?')
# spPr.add_attribute('bwMode')

# spTree = ElementDef('p:spTree', 'CT_GroupShape')
# spTree.add_child('p:nvGrpSpPr'   , cardinality='1')
# spTree.add_child('p:grpSpPr'     , cardinality='1')
# spTree.add_child('p:sp'          , cardinality='?')
# spTree.add_child('p:grpSp'       , cardinality='?')
# spTree.add_child('p:graphicFrame', cardinality='?')
# spTree.add_child('p:cxnSp'       , cardinality='?')
# spTree.add_child('p:pic'         , cardinality='?')
# spTree.add_child('p:contentPart' , cardinality='?')
# spTree.add_child('p:extLst'      , cardinality='?')

# stretch = ElementDef('a:stretch', 'CT_StretchInfoProperties')
# stretch.add_child('a:fillRect' , cardinality='?')

# txBody = ElementDef('p:txBody', 'CT_TextBody')
# txBody.add_child('a:bodyPr'   , cardinality='1')
# txBody.add_child('a:lstStyle' , cardinality='?')
# txBody.add_child('a:p'        , cardinality='+')

