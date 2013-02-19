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

    @property
    def indent(self):
        indent_len = 8 - len(self.tagname)
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

attribute_accessor = Template('''\
    $name$padding = property(lambda self: self.get('$name'),
$indent                lambda self, value: self.set('$name', value))
''')

class_def_head = Template('''\
class $classname(ElementBase):
    """<$nsprefix:$tagname> custom element class"""
''')

class_mapping = Template('''\
        , '$tagname'$indent : $classname

''')

optional_child_accessor = Template('''\
    $tagname$indent = property(lambda self: _get_child_or_append(self, '$tag'))
''')

required_attribute_constructor = Template('''\
        _required_attribute(self, '$name', default='$default')
''')

required_child_accessor = Template('''\
    $tagname$indent = property(lambda self: _child(self, '$tag'))
''')

required_child_constructor = Template('''\
        _required_child(self, '$tag')
''')


rPr = ElementDef('a:rPr', 'CT_TextCharacterProperties')
rPr.add_child('a:ln', cardinality='?')
rPr.add_attribute('sz')
rPr.add_attribute('b')
rPr.add_attribute('i')
rPr.add_attribute('u')
rPr.add_attribute('strike')
rPr.add_attribute('kern')
rPr.add_attribute('cap')
rPr.add_attribute('spc')
rPr.add_attribute('baseline')


def class_template(element):
    out = ''
    out += class_mapping.substitute(element)
    out += class_def_head.substitute(element)
    if element.required_children or element.required_attributes:
        out += '    def _init(self):\n'
        for child in element.required_children:
            out += required_child_constructor.substitute(child)
        for attribute in element.required_attributes:
            out += required_attribute_constructor.substitute(attribute)
        out += '\n'
    if element.children:
        out += '    # child accessors -----------------\n'
        for child in element.required_children:
            out += required_child_accessor.substitute(child)
        for child in element.optional_children:
            out += optional_child_accessor.substitute(child)
        out += '\n'
    if element.attributes:
        out += '    # attribute accessors -------------\n'
        for attribute in element.attributes:
            out += attribute_accessor.substitute(attribute)
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

# bodyPr = ElementDef('a:bodyPr', 'CT_TextBodyProperties')
# bodyPr.add_child('a:spAutoFit', cardinality='?')
# bodyPr.add_attribute('wrap')
# bodyPr.add_attribute('rtlCol')
# bodyPr.add_attribute('anchor')
# bodyPr.add_attribute('anchorCtr')

# off = ElementDef('a:off', 'CT_Point2D')
# off.add_attribute('x', required=True, default='0')
# off.add_attribute('y', required=True, default='0')

# p = ElementDef('a:p', 'CT_TextParagraph')
# p.add_child('a:pPr', cardinality='?')
# p.add_child('a:r', cardinality='*')
# p.add_child('a:br', cardinality='*')
# p.add_child('a:fld', cardinality='*')
# p.add_child('a:endParaRPr', cardinality='?')

# txBody = ElementDef('p:txBody', 'CT_TextBody')
# txBody.add_child('p:bodyPr'   , cardinality='1')
# txBody.add_child('a:lstStyle' , cardinality='?')
# txBody.add_child('a:p'        , cardinality='+')

