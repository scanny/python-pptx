#!/usr/bin/env python
# -*- coding: utf-8 -*-

# objectify_lab.py
#

"""
Experimental code to trial lxml.objectify as possible replacement for most
oxml custom element classes.
"""

from lxml import etree
from lxml import objectify
from lxml.objectify import Element, SubElement

# ============================================================================
# utility items
# ============================================================================

nsmap =\
    { 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    , 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    , 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
    }

def nsdecls(*prefixes):
    return ' '.join(['xmlns:%s="%s"' % (pfx, nsmap[pfx]) for pfx in prefixes])

def pfxdtag(tag):
    """
    Return short-form prefixed tag from fully qualified (Clark notation)
    tagname.
    """
    uri, tagroot = tag[1:].split('}')
    prefix = reverse_nsmap[uri]
    return '%s:%s' % (prefix, tagroot)

def qn(tag):
    """
    qn is short for *qualified name*. Return fully qualified (Clark notation)
    tagname corresponding to short-form prefixed tagname *tag*.
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{%s}%s' % (uri, tagroot)

def new(tag, **extra):
    return Element(qn(tag), **extra)

def sub_elm(parent, tag, **extra):
    return SubElement(parent, qn(tag), **extra)

# ============================================================================
# generate XML for Table shape
# ============================================================================

def empty_cell():
    tc = new('a:tc')
    tc.txBody = new('a:txBody')
    tc.txBody.bodyPr = new('a:bodyPr')
    tc.txBody.lstStyle = new('a:lstStyle')
    tc.txBody.p = new('a:p')
    tc.tcPr = new('a:tcPr')
    return tc


graphicFrame_tmpl = """
    <p:graphicFrame %s>
      <p:nvGraphicFramePr>
        <p:cNvPr id="%s" name="%s"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="%s" y="%s"/>
        <a:ext cx="%s" cy="%s"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1" bandRow="1">
              <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
            </a:tblPr>
            <a:tblGrid/>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>""" % (nsdecls('p', 'a'),
                            '%d', '%s', '%d', '%d', '%d', '%d')

sp_id = 2
name  = 'Table 1'
x     = 1524000
y     = 1397000
cx    = 6096000
cy    = 741680
rows  = 2
cols  = 2

rowheight = '370840'
colwidth  = '3048000'


graphicFrame_xml = graphicFrame_tmpl % (sp_id, name, x, y, cx, cy)

graphicFrame = objectify.fromstring(graphicFrame_xml)

tbl = graphicFrame[qn('a:graphic')].graphicData.tbl

for row in range(rows):
    # tr = sub_elm(tbl, 'a:tr', h=rowheight)
    tr = new('a:tr', h=rowheight)
    for col in range(cols):
        sub_elm(tbl.tblGrid, 'a:gridCol', w=colwidth)
        tr.append(empty_cell())
    tbl.append(tr)


objectify.deannotate(graphicFrame, cleanup_namespaces=True)
print etree.tostring(graphicFrame, pretty_print = True)



# ============================================================================
# early experiments with objectivy
# ============================================================================

# def tagname(tag):
#     tag_parts = tag.split('}')
#     return tag_parts[1]
# 
# def print_subtree(elm, indent=0):
#     indent_spaces = indent * ' '
#     print '%s%s' % (indent_spaces, tagname(elm.tag))
#     for child in elm.iterchildren():
#         print_subtree(child, indent+2)
# 
# tree = objectify.parse('slide1.xml')
# sld = tree.getroot()
# # print_subtree(sld)
# # print tagname(sld.cSld.spTree.tag)


