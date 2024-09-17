#===============================================================================
#
#  Flatmap viewer and annotation tools
#
#  Copyright (c) 2019  David Brooks
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.
#
#===============================================================================

import os.path

import pptx.oxml as oxml
import pptx.oxml.ns as ns

from pptx.oxml.shapes.autoshape import CT_GeomGuideList
from pptx.oxml.simpletypes import XsdString

from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne
)

#===============================================================================

ns._nsmap['drawml'] = ("http://www.ecma-international.org/flat/publications/standards/Ec"
                       "ma-376/drawingml/")

#===============================================================================

class PresetShapeDefinition(BaseOxmlElement):
    """`drawml:presetShapeDefinition` element class."""

    presetShapeLst = ZeroOrMore("drawml:presetShape")

    @classmethod
    def new(cls, xml):
        """Return shape definitions configured as ..."""
        return oxml.parse_xml(xml)

#===============================================================================

class Geometry2D(BaseOxmlElement):

    avLst = ZeroOrOne("a:avLst")
    gdLst = ZeroOrOne("a:gdLst")
    pathLst = ZeroOrOne("a:pathLst")

#===============================================================================

class PresetShape(Geometry2D):
    """`drawml:PresetShape` element class."""

    name = RequiredAttribute("name", XsdString)

#===============================================================================

oxml.register_element_cls("a:custGeom", Geometry2D)
oxml.register_element_cls("a:gdLst", CT_GeomGuideList)

oxml.register_element_cls("drawml:presetShapeDefinition", PresetShapeDefinition)
oxml.register_element_cls("drawml:presetShape", PresetShape)

#===============================================================================

class Shapes(object):

    definitions_ = {}



    with open(os.path.join(os.path.dirname(__file__), 'presetShapeDefinitions.xml'), 'rb') as defs:
        for defn in PresetShapeDefinition.new(defs.read()):
            definitions_[defn.name] = defn

    @staticmethod
    def lookup(name):
        return Shapes.definitions_[name]

#===============================================================================
