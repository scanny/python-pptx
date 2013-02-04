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
    
    def __init__(self, classname, nsprefix, tagname):
        self.instances.append(self)
        self.classname  = classname
        self.nsprefix   = nsprefix
        self.tagname    = tagname
        self.children   = []
        self.attributes = []
    
    def __getitem__(self, key):
        return self.__getattribute__(key)
    

class_head = Template('''\
class $classname(BaseElement):
    """Wrapper for ``<$nsprefix:$tagname>`` element"""
    def __init__(self):
        super($classname, self).__init__()
        self._element = Element(qname('$nsprefix', '$tagname'), nsmap=nsmap)
''')

required_children = Template('''\
        self.$tagname = $classname()
        self._element.append(self.$tagname.element)
''')

attributes = Template('''\
    
    @property
    def $attr_name(self):
        return self._element.get('$attr_name')
    
    @$attr_name.setter
    def $attr_name(self, value):
        self._element.set('$attr_name', str(value))
''')


nvPr      = ElementDef('CT_ApplicationNonVisualDrawingProps' , 'p', 'nvPr'      )
avLst     = ElementDef('CT_GeomGuideList'                    , 'a', 'avLst'     )
noFill    = ElementDef('CT_NoFillProperties'                 , 'a', 'noFill'    )
cNvPr     = ElementDef('CT_NonVisualDrawingProps'            , 'p', 'cNvPr'     )
cNvSpPr   = ElementDef('CT_NonVisualDrawingShapeProps'       , 'p', 'cNvSpPr'   )
off       = ElementDef('CT_Point2D'                          , 'a', 'off'       )
ext       = ElementDef('CT_PositiveSize2D'                   , 'a', 'ext'       )
prstGeom  = ElementDef('CT_PresetGeometry2D'                 , 'a', 'prstGeom'  )
sp        = ElementDef('CT_Shape'                            , 'p', 'sp'        )
nvSpPr    = ElementDef('CT_ShapeNonVisual'                   , 'p', 'nvSpPr'    )
spPr      = ElementDef('CT_ShapeProperties'                  , 'p', 'spPr'      )
txBody    = ElementDef('CT_TextBody'                         , 'p', 'txBody'    )
bodyPr    = ElementDef('CT_TextBodyProperties'               , 'a', 'bodyPr'    )
lstStyle  = ElementDef('CT_TextListStyle'                    , 'a', 'lstStyle'  )
p         = ElementDef('CT_TextParagraph'                    , 'a', 'p'         )
spAutoFit = ElementDef('CT_TextShapeAutofit'                 , 'a', 'spAutoFit' )
xfrm      = ElementDef('CT_Transform2D'                      , 'a', 'xfrm'      )

sp.children       = [nvSpPr, spPr, txBody]
nvSpPr.children   = [cNvPr, cNvSpPr, nvPr]
spPr.children     = [xfrm, prstGeom, noFill]
txBody.children   = [bodyPr, lstStyle, p]
xfrm.children     = [off, ext]
prstGeom.children = [avLst]
bodyPr.children   = [spAutoFit]

cNvPr.attributes    = ['id', 'name']
cNvSpPr.attributes  = ['txBox']
prstGeom.attributes = ['prst']
off.attributes      = ['x', 'y']
ext.attributes      = ['cx', 'cy']
bodyPr.attributes   = ['wrap']

elements = ElementDef.instances


out = ''
for element in elements:
    out += class_head.substitute(element)
    for child in element.children:
        out += required_children.substitute(child)
    for attr_name in element.attributes:
        out += attributes.substitute(dict(attr_name=attr_name))
    out += '    \n\n'

print out

sys.exit()

# ----

# out = tmpl.substitute(classname='CT_ShapeProperties', nsprefix='p', tagname='spPr'  )
# out = tmpl.substitute(classname='CT_TextBody'       , nsprefix='p', tagname='txBody')

    # def __sp(self, sp_id, shapename, x, y, cx, cy, is_textbox=False):
    #     
    #     """Return minimal ``<p:sp>`` element based on parameters."""
    #     nvSpPr = etree.SubElement(sp, qname('p', 'nvSpPr'))
    #     spPr   = etree.SubElement(sp, qname('p', 'spPr'  ))
    #     txBody = etree.SubElement(sp, qname('p', 'txBody'))
    #     
    #     cNvPr   = etree.SubElement(nvSpPr, qname('p', 'cNvPr'  ))
    #     cNvSpPr = etree.SubElement(nvSpPr, qname('p', 'cNvSpPr'))
    #     nvPr    = etree.SubElement(nvSpPr, qname('p', 'nvPr'   ))
    #     cNvPr.set('id', str(sp_id))
    #     cNvPr.set('name', shapename)
    #     if is_textbox:
    #         cNvSpPr.set('txBox', '1')
    #     
    #     xfrm     = etree.SubElement(spPr, qname('a', 'xfrm'    ))
    #     prstGeom = etree.SubElement(spPr, qname('a', 'prstGeom'))
    #     noFill   = etree.SubElement(spPr, qname('a', 'noFill'  ))
    #     prstGeom.set('prst', 'rect')
    #     
    #     bodyPr   = etree.SubElement(txBody, qname('a', 'bodyPr'  ))
    #     lstStyle = etree.SubElement(txBody, qname('a', 'lstStyle'))
    #     p        = etree.SubElement(txBody, qname('a', 'p'       ))
    #     
    #     off = etree.SubElement(xfrm, qname('a', 'off'))
    #     ext = etree.SubElement(xfrm, qname('a', 'ext'))
    #     off.set('x', str(x))
    #     off.set('y', str(y))
    #     ext.set('cx', str(cx))
    #     ext.set('cy', str(cy))
    #     
    #     avLst = etree.SubElement(prstGeom, qname('a', 'avLst'))
    #     
    #     spAutoFit = etree.SubElement(bodyPr, qname('a', 'spAutoFit'))
    #     
    #     return sp


