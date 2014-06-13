# encoding: utf-8

"""
Test suite for the pptx.oxml.xmlchemy module, focused on the metaclass and
element and attribute definition classes. A major part of the fixture is
provided by the metaclass-built test classes at the end of the file.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.exc import InvalidXmlError
from pptx.oxml import register_element_cls
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import BaseIntType
from pptx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrOne

from ..unitdata import EtreeBaseBuilder as BaseBuilder


class DescribeCustomElementClass(object):

    def it_has_the_MetaOxmlElement_metaclass(self):
        assert type(CT_Parent).__name__ == 'MetaOxmlElement'


class DescribeRequiredAttribute(object):

    def it_adds_a_getter_property_for_the_attr_value(self, getter_fixture):
        parent, reqAttr_python_value = getter_fixture
        assert parent.reqAttr is reqAttr_python_value

    def it_adds_a_setter_property_for_the_attr(self, setter_fixture):
        parent, value, expected_xml = setter_fixture
        parent.reqAttr = value
        assert parent.xml == expected_xml

    def it_adds_a_docstring_for_the_property(self):
        assert CT_Parent.reqAttr.__doc__.startswith(
            "ST_IntegerType type-converted value of "
        )

    def it_raises_on_get_when_attribute_not_present(self):
        parent = a_parent().with_nsdecls().element
        with pytest.raises(InvalidXmlError):
            parent.reqAttr

    def it_raises_on_assign_invalid_value(self, invalid_assign_fixture):
        parent, value, expected_exception = invalid_assign_fixture
        with pytest.raises(expected_exception):
            parent.reqAttr = value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getter_fixture(self):
        parent = a_parent().with_nsdecls().with_reqAttr('42').element
        return parent, 42

    @pytest.fixture(params=[
        (None, TypeError),
        (-4,   ValueError),
        ('2',  TypeError),
    ])
    def invalid_assign_fixture(self, request):
        invalid_value, expected_exception = request.param
        parent = a_parent().with_nsdecls().with_reqAttr(1).element
        return parent, invalid_value, expected_exception

    @pytest.fixture
    def setter_fixture(self):
        parent = a_parent().with_nsdecls().with_reqAttr('42').element
        value = 24
        expected_xml = a_parent().with_nsdecls().with_reqAttr(value).xml()
        return parent, value, expected_xml


class DescribeZeroOrOne(object):

    def it_adds_a_getter_property_for_the_child_element(self, getter_fixture):
        parent, zooChild = getter_fixture
        assert parent.zooChild is zooChild

    def it_adds_an_add_method_for_the_child_element(self, add_fixture):
        parent, expected_xml = add_fixture
        zooChild = parent._add_zooChild()
        assert parent.xml == expected_xml
        assert isinstance(zooChild, CT_ZooChild)
        assert parent._add_zooChild.__doc__.startswith(
            'Add a new ``<p:zooChild>`` child element '
        )

    def it_adds_an_insert_method_for_the_child_element(self, insert_fixture):
        parent, zooChild, expected_xml = insert_fixture
        parent._insert_zooChild(zooChild)
        assert parent.xml == expected_xml
        assert parent._insert_zooChild.__doc__.startswith(
            'Insert the passed ``<p:zooChild>`` '
        )

    def it_adds_a_get_or_add_method_for_the_child_element(
            self, get_or_add_fixture):
        parent, expected_xml = get_or_add_fixture
        zooChild = parent.get_or_add_zooChild()
        assert isinstance(zooChild, CT_ZooChild)
        assert parent.xml == expected_xml

    def it_adds_a_remover_method_for_the_child_element(self, remove_fixture):
        parent, expected_xml = remove_fixture
        parent._remove_zooChild()
        assert parent.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self):
        parent = self.parent_bldr(False).element
        expected_xml = self.parent_bldr(True).xml()
        return parent, expected_xml

    @pytest.fixture(params=[True, False])
    def getter_fixture(self, request):
        zooChild_is_present = request.param
        parent = self.parent_bldr(zooChild_is_present).element
        zooChild = parent.find(qn('p:zooChild'))  # None if not found
        return parent, zooChild

    @pytest.fixture(params=[True, False])
    def get_or_add_fixture(self, request):
        zooChild_is_present = request.param
        parent = self.parent_bldr(zooChild_is_present).element
        expected_xml = self.parent_bldr(True).xml()
        return parent, expected_xml

    @pytest.fixture
    def insert_fixture(self):
        parent = a_parent().with_nsdecls().element
        zooChild = a_zooChild().with_nsdecls().element
        expected_xml = self.parent_bldr(True).xml()
        return parent, zooChild, expected_xml

    @pytest.fixture(params=[True, False])
    def remove_fixture(self, request):
        zooChild_is_present = request.param
        parent = self.parent_bldr(zooChild_is_present).element
        expected_xml = self.parent_bldr(False).xml()
        return parent, expected_xml

    # fixture components ---------------------------------------------

    def parent_bldr(self, zooChild_is_present):
        parent_bldr = a_parent().with_nsdecls()
        if zooChild_is_present:
            parent_bldr.with_child(a_zooChild())
        return parent_bldr


# --------------------------------------------------------------------
# static shared fixture
# --------------------------------------------------------------------

class ST_IntegerType(BaseIntType):

    @classmethod
    def validate(cls, value):
        cls.validate_int(value)
        if value < 1 or value > 42:
            raise ValueError(
                "value must be in range 1 to 42 inclusive"
            )


class CT_Parent(BaseOxmlElement):
    """
    ``<p:parent>`` element, an invented element for use in testing.
    """
    zooChild = ZeroOrOne('p:zooChild', successors=())
    reqAttr = RequiredAttribute('reqAttr', ST_IntegerType)


class CT_ZooChild(BaseOxmlElement):
    """
    Zoo standing for 'ZeroOrOne', ``<p:zooChild>`` element, an invented
    element for use in testing.
    """


register_element_cls('p:parent', CT_Parent)
register_element_cls('p:zooChild',  CT_ZooChild)


class CT_ParentBuilder(BaseBuilder):
    __tag__ = 'p:parent'
    __nspfxs__ = ('p',)
    __attrs__ = ('reqAttr',)


class CT_ZooChildBuilder(BaseBuilder):
    __tag__ = 'p:zooChild'
    __nspfxs__ = ('p',)
    __attrs__ = ()


def a_parent():
    return CT_ParentBuilder()


def a_zooChild():
    return CT_ZooChildBuilder()
