# encoding: utf-8

"""
Temporary stand-in for main oxml module that came across with the
PackageReader transplant. Probably much will get replaced with objects from
the pptx.oxml.core and then this module will either get deleted or only hold
the package related custom element classes.
"""

from __future__ import absolute_import

from lxml import etree, objectify

from .constants import NAMESPACE as NS, RELATIONSHIP_TARGET_MODE as RTM


# configure objectified XML parser
fallback_lookup = objectify.ObjectifyElementClassLookup()
element_class_lookup = etree.ElementNamespaceClassLookup(fallback_lookup)
oxml_parser = etree.XMLParser(remove_blank_text=True)
oxml_parser.set_element_class_lookup(element_class_lookup)

nsmap = {
    'ct': NS.OPC_CONTENT_TYPES,
    'pr': NS.OPC_RELATIONSHIPS,
}


# ===========================================================================
# functions
# ===========================================================================

def oxml_fromstring(text):
    """``etree.fromstring()`` replacement that uses oxml parser"""
    return objectify.fromstring(text, oxml_parser)


def oxml_tostring(elm, encoding=None, pretty_print=False, standalone=None):
    # if xsi parameter is not set to False, PowerPoint won't load without a
    # repair step; deannotate removes some original xsi:type tags in core.xml
    # if this parameter is left out (or set to True)
    objectify.deannotate(elm, xsi=False, cleanup_namespaces=True)
    return etree.tostring(elm, encoding=encoding, pretty_print=pretty_print,
                          standalone=standalone)


# ===========================================================================
# Custom element classes
# ===========================================================================

class OxmlBaseElement(objectify.ObjectifiedElement):
    """
    Base class for all custom element classes, to add standardized behavior
    to all classes in one place.
    """
    @property
    def xml(self):
        """
        Return XML string for this element, suitable for testing purposes.
        Pretty printed for readability and without an XML declaration at the
        top.
        """
        return oxml_tostring(self, encoding='unicode', pretty_print=True)


class CT_Default(OxmlBaseElement):
    """
    ``<Default>`` element, specifying the default content type to be applied
    to a part with the specified extension.
    """
    @property
    def content_type(self):
        """
        String held in the ``ContentType`` attribute of this ``<Default>``
        element.
        """
        return self.get('ContentType')

    @property
    def extension(self):
        """
        String held in the ``Extension`` attribute of this ``<Default>``
        element.
        """
        return self.get('Extension')

    @staticmethod
    def new(ext, content_type):
        """
        Return a new ``<Default>`` element with attributes set to parameter
        values.
        """
        xml = '<Default xmlns="%s"/>' % nsmap['ct']
        default = oxml_fromstring(xml)
        default.set('Extension', ext[1:])
        default.set('ContentType', content_type)
        objectify.deannotate(default, cleanup_namespaces=True)
        return default


class CT_Override(OxmlBaseElement):
    """
    ``<Override>`` element, specifying the content type to be applied for a
    part with the specified partname.
    """
    @property
    def content_type(self):
        """
        String held in the ``ContentType`` attribute of this ``<Override>``
        element.
        """
        return self.get('ContentType')

    @staticmethod
    def new(partname, content_type):
        """
        Return a new ``<Override>`` element with attributes set to parameter
        values.
        """
        xml = '<Override xmlns="%s"/>' % nsmap['ct']
        override = oxml_fromstring(xml)
        override.set('PartName', partname)
        override.set('ContentType', content_type)
        objectify.deannotate(override, cleanup_namespaces=True)
        return override

    @property
    def partname(self):
        """
        String held in the ``PartName`` attribute of this ``<Override>``
        element.
        """
        return self.get('PartName')


class CT_Relationship(OxmlBaseElement):
    """
    ``<Relationship>`` element, representing a single relationship from a
    source to a target part.
    """
    @staticmethod
    def new(rId, reltype, target, target_mode=RTM.INTERNAL):
        """
        Return a new ``<Relationship>`` element.
        """
        xml = '<Relationship xmlns="%s"/>' % nsmap['pr']
        relationship = oxml_fromstring(xml)
        relationship.set('Id', rId)
        relationship.set('Type', reltype)
        relationship.set('Target', target)
        if target_mode == RTM.EXTERNAL:
            relationship.set('TargetMode', RTM.EXTERNAL)
        objectify.deannotate(relationship, cleanup_namespaces=True)
        return relationship

    @property
    def rId(self):
        """
        String held in the ``Id`` attribute of this ``<Relationship>``
        element.
        """
        return self.get('Id')

    @property
    def reltype(self):
        """
        String held in the ``Type`` attribute of this ``<Relationship>``
        element.
        """
        return self.get('Type')

    @property
    def target_ref(self):
        """
        String held in the ``Target`` attribute of this ``<Relationship>``
        element.
        """
        return self.get('Target')

    @property
    def target_mode(self):
        """
        String held in the ``TargetMode`` attribute of this
        ``<Relationship>`` element, either ``Internal`` or ``External``.
        Defaults to ``Internal``.
        """
        return self.get('TargetMode', RTM.INTERNAL)


class CT_Relationships(OxmlBaseElement):
    """
    ``<Relationships>`` element, the root element in a .rels file.
    """
    def add_rel(self, rId, reltype, target, is_external=False):
        """
        Add a child ``<Relationship>`` element with attributes set according
        to parameter values.
        """
        target_mode = RTM.EXTERNAL if is_external else RTM.INTERNAL
        relationship = CT_Relationship.new(rId, reltype, target, target_mode)
        self.append(relationship)

    @staticmethod
    def new():
        """
        Return a new ``<Relationships>`` element.
        """
        xml = '<Relationships xmlns="%s"/>' % nsmap['pr']
        relationships = oxml_fromstring(xml)
        objectify.deannotate(relationships, cleanup_namespaces=True)
        return relationships

    @property
    def xml(self):
        """
        Return XML string for this element, suitable for saving in a .rels
        stream, not pretty printed and with an XML declaration at the top.
        """
        return oxml_tostring(self, encoding='UTF-8', standalone=True)


class CT_Types(OxmlBaseElement):
    """
    ``<Types>`` element, the container element for Default and Override
    elements in [Content_Types].xml.
    """
    def add_default(self, ext, content_type):
        """
        Add a child ``<Default>`` element with attributes set to parameter
        values.
        """
        default = CT_Default.new(ext, content_type)
        self.append(default)

    def add_override(self, partname, content_type):
        """
        Add a child ``<Override>`` element with attributes set to parameter
        values.
        """
        override = CT_Override.new(partname, content_type)
        self.append(override)

    @property
    def defaults(self):
        try:
            return self.Default[:]
        except AttributeError:
            return []

    @staticmethod
    def new():
        """
        Return a new ``<Types>`` element.
        """
        xml = '<Types xmlns="%s"/>' % nsmap['ct']
        types = oxml_fromstring(xml)
        objectify.deannotate(types, cleanup_namespaces=True)
        return types

    @property
    def overrides(self):
        try:
            return self.Override[:]
        except AttributeError:
            return []


ct_namespace = element_class_lookup.get_namespace(nsmap['ct'])
ct_namespace['Default'] = CT_Default
ct_namespace['Override'] = CT_Override
ct_namespace['Types'] = CT_Types

pr_namespace = element_class_lookup.get_namespace(nsmap['pr'])
pr_namespace['Relationship'] = CT_Relationship
pr_namespace['Relationships'] = CT_Relationships
