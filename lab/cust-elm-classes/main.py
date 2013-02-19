#!/usr/bin/env python
# -*- coding: utf-8 -*-

# lxml Custom Element Classes
#

"""
Experimental code to explore lxml custom Element classes
"""

import sys

from lxml import etree

nsmap =\
    { 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    , 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    , 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
    }

etree.register_namespace('a', nsmap['a'])
etree.register_namespace('p', nsmap['p'])

def _Element(tag, nsmap=None):
    if nsmap:
        element = parser.makeelement(_qtag(tag), nsmap=nsmap)
    else:
        element = parser.makeelement(_qtag(tag))
    return element

def _child(element, child_tagname):
    """
    Return direct child of *element* having *child_tagname* or :class:`None`
    if no such child element is present.
    """
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None

def _child_list(element, child_tagname):
    """
    Return list containing the direct children of *element* having
    *child_tagname*.
    """
    xpath = './%s' % child_tagname
    return element.xpath(xpath, namespaces=nsmap)

def _get_child_or_append(element, tag):
    child = _child(element, tag)
    if child is None:
        child = _Element(tag)
        element.append(child)
    return child

def _get_child_or_first(element, tag):
    child = _child(element, tag)
    if child is None:
        child = _Element(tag)
        element.insert(0, child)
    return child

def _get_child_or_insert(element, tag, idx):
    child = _child(element, tag)
    if child is None:
        child = _Element(tag)
        element.insert(idx, child)
    return child

def _nsmap(*prefixes):
    """
    Return a dict containing the subset namespace prefix mappings specified by
    *prefixes*. Any number of namespace prefixes can be supplied, e.g.
    namespaces('a', 'r', 'p').
    """
    namespaces = {}
    for prefix in prefixes:
        namespaces[prefix] = nsmap[prefix]
    return namespaces

def _qtag(tag):
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{%s}%s' % (uri, tagroot)

def _required_attribute(element, name, default):
    """
    Add attribute with default value to element if it doesn't already exist.
    """
    if element.get(name) is None:
        element.set(name, default)

def _required_child(parent, tag):
    """
    Add child element with *tag* to *parent* if it doesn't already exist.
    """
    if _child(parent, tag) is None:
        parent.append(_Element(tag))


# ============================================================================
# Custom element classes
# ============================================================================

class ElementBase(etree.ElementBase):
    """Provides the base element interface for custom element classes"""
    pass

class CT_NonVisualDrawingProps(ElementBase):
    """<p:cNvPr>"""
    # attributes ----------------------
    id   = property(lambda self: self.get('id'),
                    lambda self, value: self.set('id', str(value)))
    name = property(lambda self: self.get('name'),
                    lambda self, value: self.set('name', value))


class CT_NonVisualDrawingShapeProps(ElementBase):
    """<p:cNvSpPr>"""
    # attributes ----------------------
    txBox = property(lambda self: self.get('txBox'),
                     lambda self, value: self.set('txBox', value))


class CT_Point2D(ElementBase):
    """<a:off>"""
    def _init(self):
        _required_attribute(self, 'x', default='0')
        _required_attribute(self, 'y', default='0')

    # attribute accessors -------------
    x = property(lambda self: self.get('x'),
                 lambda self, value: self.set('x', value))
    y = property(lambda self: self.get('y'),
                 lambda self, value: self.set('y', value))


class CT_PositiveSize2D(ElementBase):
    """<a:ext>"""
    def _init(self):
        _required_attribute(self, 'cx', default='0')
        _required_attribute(self, 'cy', default='0')

    # attribute accessors -------------
    cx = property(lambda self: self.get('cx'),
                  lambda self, value: self.set('cx', value))
    cy = property(lambda self: self.get('cy'),
                  lambda self, value: self.set('cy', value))


class CT_PresetGeometry2D(ElementBase):
    """<a:prstGeom>"""
    # attributes ----------------------
    prst = property(lambda self: self.get('prst'),
                    lambda self, value: self.set('prst', value))
    # child accessors -----------------
    avLst = property(lambda self: _get_child_or_append(self, 'a:avLst'))


class CT_RegularTextRun(ElementBase):
    """<a:r> custom element class"""
    def _init(self):
        _required_child(self, 'a:t')

    # child accessors -----------------
    t   = property(lambda self: _child(self, 'a:t'))
    rPr = property(lambda self: _get_child_or_first(self, 'a:rPr'))


class CT_Shape(ElementBase):
    """<p:sp> Custom element class"""
    def _init(self):
        _required_child(self, 'p:nvSpPr')
        _required_child(self, 'p:spPr')

    # children ------------------------
    nvSpPr = property(lambda self: _child(self, 'p:nvSpPr'))
    spPr   = property(lambda self: _child(self, 'p:spPr'))
    txBody = property(lambda self: _get_child_or_append(self, 'p:txBody'))

    # convenience accessors
    id   = property(lambda self: self.nvSpPr.cNvPr.id)
    name = property(lambda self: self.nvSpPr.cNvPr.name)


class CT_ShapeNonVisual(ElementBase):
    """<p:nvSpPr>"""
    def _init(self):
        _required_child(self, 'p:cNvPr')
        _required_child(self, 'p:cNvSpPr')
        _required_child(self, 'p:nvPr')

    # child accessors -----------------
    cNvPr   = property(lambda self: _child(self, 'p:cNvPr'))
    cNvSpPr = property(lambda self: _child(self, 'p:cNvSpPr'))


class CT_ShapeProperties(ElementBase):
    """<p:spPr>"""
    # child accessors -----------------
    xfrm     = property(lambda self: _get_child_or_append(self, 'a:xfrm'))
    prstGeom = property(lambda self: _get_child_or_append(self, 'a:prstGeom'))
    noFill   = property(lambda self: _get_child_or_append(self, 'a:noFill'))


class CT_TextBody(ElementBase):
    """<p:txBody> custom element class"""
    def _init(self):
        _required_child(self, 'p:bodyPr')
        _required_child(self, 'a:p')

    # child accessors -----------------
    bodyPr   = property(lambda self: _child(self, 'p:bodyPr'))
    p        = property(lambda self: _child_list(self, 'a:p'))
    lstStyle = property(lambda self: _get_child_or_insert(self, 'a:lstStyle', 1))


class CT_TextBodyProperties(ElementBase):
    """<a:bodyPr> custom element class"""
    # child accessors -----------------
    spAutoFit = property(lambda self: _get_child_or_append(self, 'a:spAutoFit'))

    # attribute accessors -------------
    wrap      = property(lambda self: self.get('wrap'),
                         lambda self, value: self.set('wrap', value))
    rtlCol    = property(lambda self: self.get('rtlCol'),
                         lambda self, value: self.set('rtlCol', value))
    anchor    = property(lambda self: self.get('anchor'),
                         lambda self, value: self.set('anchor', value))
    anchorCtr = property(lambda self: self.get('anchorCtr'),
                         lambda self, value: self.set('anchorCtr', value))


class CT_TextCharacterProperties(ElementBase):
    """<a:rPr> custom element class"""
    # child accessors -----------------
    ln = property(lambda self: _get_child_or_append(self, 'a:ln'))

    # attribute accessors -------------
    sz       = property(lambda self: self.get('sz'),
                        lambda self, value: self.set('sz', value))
    b        = property(lambda self: self.get('b'),
                        lambda self, value: self.set('b', value))
    i        = property(lambda self: self.get('i'),
                        lambda self, value: self.set('i', value))
    u        = property(lambda self: self.get('u'),
                        lambda self, value: self.set('u', value))
    strike   = property(lambda self: self.get('strike'),
                        lambda self, value: self.set('strike', value))
    kern     = property(lambda self: self.get('kern'),
                        lambda self, value: self.set('kern', value))
    cap      = property(lambda self: self.get('cap'),
                        lambda self, value: self.set('cap', value))
    spc      = property(lambda self: self.get('spc'),
                        lambda self, value: self.set('spc', value))
    baseline = property(lambda self: self.get('baseline'),
                        lambda self, value: self.set('baseline', value))


class CT_TextParagraph(ElementBase):
    """<a:p> custom element class"""
    # child accessors -----------------
    pPr        = property(lambda self: _get_child_or_append(self, 'a:pPr'))
    r          = property(lambda self: _get_child_or_append(self, 'a:r'))
    br         = property(lambda self: _get_child_or_append(self, 'a:br'))
    fld        = property(lambda self: _get_child_or_append(self, 'a:fld'))
    endParaRPr = property(lambda self: _get_child_or_append(self, 'a:endParaRPr'))


class CT_Transform2D(ElementBase):
    """<a:xfrm>"""
    # child accessors -----------------
    off = property(lambda self: _get_child_or_append(self, 'a:off'))
    ext = property(lambda self: _get_child_or_append(self, 'a:ext'))


class ElementClassLookup(etree.CustomElementClassLookup):
    cls_map =\
        { 'bodyPr'   : CT_TextBodyProperties
        , 'cNvPr'    : CT_NonVisualDrawingProps
        , 'cNvSpPr'  : CT_NonVisualDrawingShapeProps
        , 'ext'      : CT_PositiveSize2D
        , 'nvSpPr'   : CT_ShapeNonVisual
        , 'off'      : CT_Point2D
        , 'p'        : CT_TextParagraph
        , 'prstGeom' : CT_PresetGeometry2D
        , 'r'        : CT_RegularTextRun
        , 'rPr'      : CT_TextCharacterProperties
        , 'sp'       : CT_Shape
        , 'spPr'     : CT_ShapeProperties
        , 'txBody'   : CT_TextBody
        , 'xfrm'     : CT_Transform2D
        }

    def lookup(self, node_type, document, namespace, name):
        if name in self.cls_map:
            return self.cls_map[name]
        return None


parser = etree.XMLParser()
parser.set_element_class_lookup(ElementClassLookup())


# ============================================================================
# construct from scratch
# ============================================================================

sp = _Element('p:sp', nsmap=_nsmap('p', 'a'))
sp.nvSpPr.cNvPr.id = 7
sp.nvSpPr.cNvPr.name = 'TextBox 6'
sp.nvSpPr.cNvSpPr.txBox = '1'

sp.spPr.xfrm.off.x  = '5580112'
sp.spPr.xfrm.off.y  = '2924944'
sp.spPr.xfrm.ext.cx = '1005403'
sp.spPr.xfrm.ext.cy =  '369332'

sp.spPr.prstGeom.prst = 'rect'
sp.spPr.prstGeom.avLst

sp.spPr.noFill

sp.txBody
sp.txBody.bodyPr.wrap = 'none'
sp.txBody.bodyPr.rtlCol = '0'
sp.txBody.bodyPr.spAutoFit

sp.txBody.lstStyle

sp.txBody.p[0].r.rPr.sz = "2400"
sp.txBody.p[0].r.rPr.b  = "1"
sp.txBody.p[0].r.t.text = "Test text"


# print XML
print "\n%s" % etree.tostring(sp, pretty_print=True)

sys.exit()


# ============================================================================
# parse from XML string
# ============================================================================

xml = """\
<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="7" name="TextBox 6"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
</p:sp>
"""

# <p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
# xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">

#   <p:spPr>
#     <a:xfrm>
#       <a:off x="5580112" y="2924944"/>
#       <a:ext cx="1005403" cy="369332"/>
#     </a:xfrm>
#     <a:prstGeom prst="rect">
#       <a:avLst/>
#     </a:prstGeom>
#     <a:noFill/>
#   </p:spPr>
#   <p:txBody>
#     <a:bodyPr wrap="none" rtlCol="0">
#       <a:spAutoFit/>
#     </a:bodyPr>
#     <a:lstStyle/>
#     <a:p>
#       <a:r>
#         <a:rPr lang="en-US" dirty="0" smtClean="0"/>
#         <a:t>Test text</a:t>
#       </a:r>
#     </a:p>
#   </p:txBody>


sp = etree.XML(xml, parser)

# print "repr(sp) => %r" % sp
# print "type(sp) => %s" % type(sp)
print "sp.id    => %s" % sp.id
print "sp.name  => %s" % sp.name
print "len(sp)  => %d" % len(sp)
print
print "sp.nvSpPr => %s" % type(sp.nvSpPr)
print "sp.nvSpPr.cNvPr.id => %s" % sp.nvSpPr.cNvPr.id
print "sp.nvSpPr.cNvPr.name => %s" % sp.nvSpPr.cNvPr.name
sp.nvSpPr.cNvPr.name = 'New Name 99'
print "sp.nvSpPr.cNvPr.name => %s" % sp.nvSpPr.cNvPr.name
print
print "________\n\n%s" % etree.tostring(sp)

sys.exit()


# ============================================================================
# Direct use of inheriting class
# ============================================================================

sp = CT_Shape()

print sp
print sp.id
print len(sp)
print etree.tostring(sp)


