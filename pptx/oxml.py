# -*- coding: utf-8 -*-
#
# oxml.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
Classes that directly manipulate Open XML and provide direct object-oriented
access to the XML elements. Classes are implemented as a wrapper around their
bit of the lxml graph that spans the entire Open XML package part, e.g. a
slide.
"""

from lxml.etree import Element

from pptx.spec import namespaces, qtag

# default namespace map
nsmap = namespaces('a', 'r', 'p')

class BaseElement(object):
    """Base class for Open XML element classes"""
    def __init__(self):
        super(BaseElement, self).__init__()
    
    @property
    def element(self):
        """Return lxml element for this Open XML element"""
        return self._element
    



# ============================================================================
# Generated Classes
# ============================================================================

class CT_ApplicationNonVisualDrawingProps(BaseElement):
    """Wrapper for ``<p:nvPr>`` element"""
    def __init__(self):
        super(CT_ApplicationNonVisualDrawingProps, self).__init__()
        self._element = Element(qtag('p:nvPr'), nsmap=nsmap)
    

class CT_GeomGuideList(BaseElement):
    """Wrapper for ``<a:avLst>`` element"""
    def __init__(self):
        super(CT_GeomGuideList, self).__init__()
        self._element = Element(qtag('a:avLst'), nsmap=nsmap)
    

class CT_NoFillProperties(BaseElement):
    """Wrapper for ``<a:noFill>`` element"""
    def __init__(self):
        super(CT_NoFillProperties, self).__init__()
        self._element = Element(qtag('a:noFill'), nsmap=nsmap)
    

class CT_NonVisualDrawingProps(BaseElement):
    """Wrapper for ``<p:cNvPr>`` element"""
    def __init__(self):
        super(CT_NonVisualDrawingProps, self).__init__()
        self._element = Element(qtag('p:cNvPr'), nsmap=nsmap)
    
    @property
    def id(self):
        return self._element.get('id')
    
    @id.setter
    def id(self, value):
        self._element.set('id', str(value))
    
    @property
    def name(self):
        return self._element.get('name')
    
    @name.setter
    def name(self, value):
        self._element.set('name', str(value))
    

class CT_NonVisualDrawingShapeProps(BaseElement):
    """Wrapper for ``<p:cNvSpPr>`` element"""
    def __init__(self):
        super(CT_NonVisualDrawingShapeProps, self).__init__()
        self._element = Element(qtag('p:cNvSpPr'), nsmap=nsmap)
    
    @property
    def txBox(self):
        return self._element.get('txBox')
    
    @txBox.setter
    def txBox(self, value):
        self._element.set('txBox', str(value))
    

class CT_Point2D(BaseElement):
    """Wrapper for ``<a:off>`` element"""
    def __init__(self):
        super(CT_Point2D, self).__init__()
        self._element = Element(qtag('a:off'), nsmap=nsmap)
    
    @property
    def x(self):
        return self._element.get('x')
    
    @x.setter
    def x(self, value):
        self._element.set('x', str(value))
    
    @property
    def y(self):
        return self._element.get('y')
    
    @y.setter
    def y(self, value):
        self._element.set('y', str(value))
    

class CT_PositiveSize2D(BaseElement):
    """Wrapper for ``<a:ext>`` element"""
    def __init__(self):
        super(CT_PositiveSize2D, self).__init__()
        self._element = Element(qtag('a:ext'), nsmap=nsmap)
    
    @property
    def cx(self):
        return self._element.get('cx')
    
    @cx.setter
    def cx(self, value):
        self._element.set('cx', str(value))
    
    @property
    def cy(self):
        return self._element.get('cy')
    
    @cy.setter
    def cy(self, value):
        self._element.set('cy', str(value))
    

class CT_PresetGeometry2D(BaseElement):
    """Wrapper for ``<a:prstGeom>`` element"""
    def __init__(self):
        super(CT_PresetGeometry2D, self).__init__()
        self._element = Element(qtag('a:prstGeom'), nsmap=nsmap)
        self.avLst = CT_GeomGuideList()
        self._element.append(self.avLst.element)
    
    @property
    def prst(self):
        return self._element.get('prst')
    
    @prst.setter
    def prst(self, value):
        self._element.set('prst', str(value))
    

class CT_Shape(BaseElement):
    """Wrapper for ``<p:sp>`` element"""
    def __init__(self):
        super(CT_Shape, self).__init__()
        self._element = Element(qtag('p:sp'), nsmap=nsmap)
        self.nvSpPr = CT_ShapeNonVisual()
        self._element.append(self.nvSpPr.element)
        self.spPr = CT_ShapeProperties()
        self._element.append(self.spPr.element)
        self.txBody = CT_TextBody()
        self._element.append(self.txBody.element)
    

class CT_ShapeNonVisual(BaseElement):
    """Wrapper for ``<p:nvSpPr>`` element"""
    def __init__(self):
        super(CT_ShapeNonVisual, self).__init__()
        self._element = Element(qtag('p:nvSpPr'), nsmap=nsmap)
        self.cNvPr = CT_NonVisualDrawingProps()
        self._element.append(self.cNvPr.element)
        self.cNvSpPr = CT_NonVisualDrawingShapeProps()
        self._element.append(self.cNvSpPr.element)
        self.nvPr = CT_ApplicationNonVisualDrawingProps()
        self._element.append(self.nvPr.element)
    

class CT_ShapeProperties(BaseElement):
    """Wrapper for ``<p:spPr>`` element"""
    def __init__(self):
        super(CT_ShapeProperties, self).__init__()
        self._element = Element(qtag('p:spPr'), nsmap=nsmap)
        self.xfrm = CT_Transform2D()
        self._element.append(self.xfrm.element)
        self.prstGeom = CT_PresetGeometry2D()
        self._element.append(self.prstGeom.element)
        self.noFill = CT_NoFillProperties()
        self._element.append(self.noFill.element)
    

class CT_TextBody(BaseElement):
    """Wrapper for ``<p:txBody>`` element"""
    def __init__(self):
        super(CT_TextBody, self).__init__()
        self._element = Element(qtag('p:txBody'), nsmap=nsmap)
        self.bodyPr = CT_TextBodyProperties()
        self._element.append(self.bodyPr.element)
        self.lstStyle = CT_TextListStyle()
        self._element.append(self.lstStyle.element)
        self.p = CT_TextParagraph()
        self._element.append(self.p.element)
    

class CT_TextBodyProperties(BaseElement):
    """Wrapper for ``<a:bodyPr>`` element"""
    def __init__(self):
        super(CT_TextBodyProperties, self).__init__()
        self._element = Element(qtag('a:bodyPr'), nsmap=nsmap)
        self.spAutoFit = CT_TextShapeAutofit()
        self._element.append(self.spAutoFit.element)
    
    @property
    def wrap(self):
        return self._element.get('wrap')
    
    @wrap.setter
    def wrap(self, value):
        self._element.set('wrap', str(value))
    

class CT_TextListStyle(BaseElement):
    """Wrapper for ``<a:lstStyle>`` element"""
    def __init__(self):
        super(CT_TextListStyle, self).__init__()
        self._element = Element(qtag('a:lstStyle'), nsmap=nsmap)
    

class CT_TextParagraph(BaseElement):
    """Wrapper for ``<a:p>`` element"""
    def __init__(self):
        super(CT_TextParagraph, self).__init__()
        self._element = Element(qtag('a:p'), nsmap=nsmap)
    

class CT_TextShapeAutofit(BaseElement):
    """Wrapper for ``<a:spAutoFit>`` element"""
    def __init__(self):
        super(CT_TextShapeAutofit, self).__init__()
        self._element = Element(qtag('a:spAutoFit'), nsmap=nsmap)
    

class CT_Transform2D(BaseElement):
    """Wrapper for ``<a:xfrm>`` element"""
    def __init__(self):
        super(CT_Transform2D, self).__init__()
        self._element = Element(qtag('a:xfrm'), nsmap=nsmap)
        self.off = CT_Point2D()
        self._element.append(self.off.element)
        self.ext = CT_PositiveSize2D()
        self._element.append(self.ext.element)
    


