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
from pptx.oxml.xmlchemy import (
    BaseOxmlElement, Choice, OneAndOnlyOne, OneOrMore, OptionalAttribute,
    RequiredAttribute, ZeroOrMore, ZeroOrOne, ZeroOrOneChoice
)

from ..unitdata import BaseBuilder


class DescribeCustomElementClass(object):

    def it_has_the_MetaOxmlElement_metaclass(self):
        assert type(CT_Parent).__name__ == 'MetaOxmlElement'


class DescribeChoice(object):

    def it_adds_a_getter_property_for_the_choice_element(
            self, getter_fixture):
        parent, expected_choice = getter_fixture
        assert parent.choice is expected_choice

    def it_adds_a_creator_method_for_the_child_element(self, new_fixture):
        parent, expected_xml = new_fixture
        choice = parent._new_choice()
        assert choice.xml == expected_xml

    def it_adds_an_insert_method_for_the_child_element(self, insert_fixture):
        parent, choice, expected_xml = insert_fixture
        parent._insert_choice(choice)
        assert parent.xml == expected_xml
        assert parent._insert_choice.__doc__.startswith(
            'Return the passed ``<p:choice>`` '
        )

    def it_adds_an_add_method_for_the_child_element(self, add_fixture):
        parent, expected_xml = add_fixture
        choice = parent._add_choice()
        assert parent.xml == expected_xml
        assert isinstance(choice, CT_Choice)
        assert parent._add_choice.__doc__.startswith(
            'Add a new ``<p:choice>`` child element '
        )

    def it_adds_a_get_or_change_to_method_for_the_child_element(
            self, get_or_change_to_fixture):
        parent, expected_xml = get_or_change_to_fixture
        choice = parent.get_or_change_to_choice()
        assert isinstance(choice, CT_Choice)
        assert parent.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self):
        parent = self.parent_bldr().element
        expected_xml = self.parent_bldr('choice').xml()
        return parent, expected_xml

    @pytest.fixture(params=[
        ('choice2', 'choice'),
        (None,      'choice'),
        ('choice',  'choice'),
    ])
    def get_or_change_to_fixture(self, request):
        before_member_tag, after_member_tag = request.param
        parent = self.parent_bldr(before_member_tag).element
        expected_xml = self.parent_bldr(after_member_tag).xml()
        return parent, expected_xml

    @pytest.fixture(params=['choice', None])
    def getter_fixture(self, request):
        choice_tag = request.param
        parent = self.parent_bldr(choice_tag).element
        expected_choice = parent.find(qn('p:choice'))  # None if not found
        return parent, expected_choice

    @pytest.fixture
    def insert_fixture(self):
        parent = (
            a_parent().with_nsdecls().with_child(
                an_oomChild()).with_child(
                an_oooChild())
        ).element
        choice = a_choice().with_nsdecls().element
        expected_xml = (
            a_parent().with_nsdecls().with_child(
                a_choice()).with_child(
                an_oomChild()).with_child(
                an_oooChild())
        ).xml()
        return parent, choice, expected_xml

    @pytest.fixture
    def new_fixture(self):
        parent = self.parent_bldr().element
        expected_xml = a_choice().with_nsdecls().xml()
        return parent, expected_xml

    # fixture components ---------------------------------------------

    def parent_bldr(self, choice_tag=None):
        parent_bldr = a_parent().with_nsdecls()
        if choice_tag == 'choice':
            parent_bldr.with_child(a_choice())
        if choice_tag == 'choice2':
            parent_bldr.with_child(a_choice2())
        return parent_bldr


class DescribeOneAndOnlyOne(object):

    def it_adds_a_getter_property_for_the_child_element(self, getter_fixture):
        parent, oooChild = getter_fixture
        assert parent.oooChild is oooChild

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getter_fixture(self):
        parent = a_parent().with_nsdecls().with_child(an_oooChild()).element
        oooChild = parent.find(qn('p:oooChild'))
        return parent, oooChild


class DescribeOneOrMore(object):

    def it_adds_a_getter_property_for_the_child_element_list(
            self, getter_fixture):
        parent, oomChild = getter_fixture
        assert parent.oomChild_lst[0] is oomChild

    def it_adds_a_creator_method_for_the_child_element(self, new_fixture):
        parent, expected_xml = new_fixture
        oomChild = parent._new_oomChild()
        assert oomChild.xml == expected_xml

    def it_adds_an_insert_method_for_the_child_element(self, insert_fixture):
        parent, oomChild, expected_xml = insert_fixture
        parent._insert_oomChild(oomChild)
        assert parent.xml == expected_xml
        assert parent._insert_oomChild.__doc__.startswith(
            'Return the passed ``<p:oomChild>`` '
        )

    def it_adds_a_private_add_method_for_the_child_element(self, add_fixture):
        parent, expected_xml = add_fixture
        oomChild = parent._add_oomChild()
        assert parent.xml == expected_xml
        assert isinstance(oomChild, CT_OomChild)
        assert parent._add_oomChild.__doc__.startswith(
            'Add a new ``<p:oomChild>`` child element '
        )

    def it_adds_a_public_add_method_for_the_child_element(self, add_fixture):
        parent, expected_xml = add_fixture
        oomChild = parent.add_oomChild()
        assert parent.xml == expected_xml
        assert isinstance(oomChild, CT_OomChild)
        assert parent._add_oomChild.__doc__.startswith(
            'Add a new ``<p:oomChild>`` child element '
        )

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self):
        parent = self.parent_bldr(False).element
        expected_xml = self.parent_bldr(True).xml()
        return parent, expected_xml

    @pytest.fixture
    def getter_fixture(self):
        parent = self.parent_bldr(True).element
        oomChild = parent.find(qn('p:oomChild'))
        return parent, oomChild

    @pytest.fixture
    def insert_fixture(self):
        parent = (
            a_parent().with_nsdecls().with_child(
                an_oooChild()).with_child(
                a_zomChild()).with_child(
                a_zooChild())
        ).element
        oomChild = an_oomChild().with_nsdecls().element
        expected_xml = (
            a_parent().with_nsdecls().with_child(
                an_oomChild()).with_child(
                an_oooChild()).with_child(
                a_zomChild()).with_child(
                a_zooChild())
        ).xml()
        return parent, oomChild, expected_xml

    @pytest.fixture
    def new_fixture(self):
        parent = self.parent_bldr(False).element
        expected_xml = an_oomChild().with_nsdecls().xml()
        return parent, expected_xml

    # fixture components ---------------------------------------------

    def parent_bldr(self, oomChild_is_present):
        parent_bldr = a_parent().with_nsdecls()
        if oomChild_is_present:
            parent_bldr.with_child(an_oomChild())
        return parent_bldr


class DescribeOptionalAttribute(object):

    def it_adds_a_getter_property_for_the_attr_value(self, getter_fixture):
        parent, optAttr_python_value = getter_fixture
        assert parent.optAttr == optAttr_python_value

    def it_adds_a_setter_property_for_the_attr(self, setter_fixture):
        parent, value, expected_xml = setter_fixture
        parent.optAttr = value
        assert parent.xml == expected_xml

    def it_adds_a_docstring_for_the_property(self):
        assert CT_Parent.optAttr.__doc__.startswith(
            "ST_IntegerType type-converted value of "
        )

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getter_fixture(self):
        parent = a_parent().with_nsdecls().with_optAttr('24').element
        return parent, 24

    @pytest.fixture(params=[36, None])
    def setter_fixture(self, request):
        value = request.param
        parent = a_parent().with_nsdecls().with_optAttr('42').element
        if value is None:
            expected_xml = a_parent().with_nsdecls().xml()
        else:
            expected_xml = a_parent().with_nsdecls().with_optAttr(value).xml()
        return parent, value, expected_xml


class DescribeRequiredAttribute(object):

    def it_adds_a_getter_property_for_the_attr_value(self, getter_fixture):
        parent, reqAttr_python_value = getter_fixture
        assert parent.reqAttr == reqAttr_python_value

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


class DescribeZeroOrMore(object):

    def it_adds_a_getter_property_for_the_child_element_list(
            self, getter_fixture):
        parent, zomChild = getter_fixture
        assert parent.zomChild_lst[0] is zomChild

    def it_adds_a_creator_method_for_the_child_element(self, new_fixture):
        parent, expected_xml = new_fixture
        zomChild = parent._new_zomChild()
        assert zomChild.xml == expected_xml

    def it_adds_an_insert_method_for_the_child_element(self, insert_fixture):
        parent, zomChild, expected_xml = insert_fixture
        parent._insert_zomChild(zomChild)
        assert parent.xml == expected_xml
        assert parent._insert_zomChild.__doc__.startswith(
            'Return the passed ``<p:zomChild>`` '
        )

    def it_adds_an_add_method_for_the_child_element(self, add_fixture):
        parent, expected_xml = add_fixture
        zomChild = parent._add_zomChild()
        assert parent.xml == expected_xml
        assert isinstance(zomChild, CT_ZomChild)
        assert parent._add_zomChild.__doc__.startswith(
            'Add a new ``<p:zomChild>`` child element '
        )

    def it_removes_the_property_root_name_used_for_declaration(self):
        assert not hasattr(CT_Parent, 'zomChild')

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self):
        parent = self.parent_bldr(False).element
        expected_xml = self.parent_bldr(True).xml()
        return parent, expected_xml

    @pytest.fixture
    def getter_fixture(self):
        parent = self.parent_bldr(True).element
        zomChild = parent.find(qn('p:zomChild'))
        return parent, zomChild

    @pytest.fixture
    def insert_fixture(self):
        parent = (
            a_parent().with_nsdecls().with_child(
                an_oomChild()).with_child(
                an_oooChild()).with_child(
                a_zooChild())
        ).element
        zomChild = a_zomChild().with_nsdecls().element
        expected_xml = (
            a_parent().with_nsdecls().with_child(
                an_oomChild()).with_child(
                an_oooChild()).with_child(
                a_zomChild()).with_child(
                a_zooChild())
        ).xml()
        return parent, zomChild, expected_xml

    @pytest.fixture
    def new_fixture(self):
        parent = self.parent_bldr(False).element
        expected_xml = a_zomChild().with_nsdecls().xml()
        return parent, expected_xml

    def parent_bldr(self, zomChild_is_present):
        parent_bldr = a_parent().with_nsdecls()
        if zomChild_is_present:
            parent_bldr.with_child(a_zomChild())
        return parent_bldr


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
            'Return the passed ``<p:zooChild>`` '
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
        parent = (
            a_parent().with_nsdecls().with_child(
                an_oomChild()).with_child(
                an_oooChild()).with_child(
                a_zomChild())
        ).element
        zooChild = a_zooChild().with_nsdecls().element
        expected_xml = (
            a_parent().with_nsdecls().with_child(
                an_oomChild()).with_child(
                an_oooChild()).with_child(
                a_zomChild()).with_child(
                a_zooChild())
        ).xml()
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


class DescribeZeroOrOneChoice(object):

    def it_adds_a_getter_for_the_current_choice(self, getter_fixture):
        parent, expected_choice = getter_fixture
        assert parent.eg_zooChoice is expected_choice

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[None, 'choice', 'choice2'])
    def getter_fixture(self, request):
        choice_tag = request.param
        parent = self.parent_bldr(choice_tag).element
        tagname = 'p:%s' % choice_tag
        expected_choice = parent.find(qn(tagname))  # None if not found
        return parent, expected_choice

    # fixture components ---------------------------------------------

    def parent_bldr(self, choice_tag=None):
        parent_bldr = a_parent().with_nsdecls()
        if choice_tag == 'choice':
            parent_bldr.with_child(a_choice())
        if choice_tag == 'choice2':
            parent_bldr.with_child(a_choice2())
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
    eg_zooChoice = ZeroOrOneChoice(
        (Choice('p:choice'), Choice('p:choice2')),
        successors=('p:oomChild', 'p:oooChild')
    )
    oomChild = OneOrMore('p:oomChild', successors=(
        'p:oooChild', 'p:zomChild', 'p:zooChild'
    ))
    oooChild = OneAndOnlyOne('p:oooChild')
    zomChild = ZeroOrMore('p:zomChild', successors=('p:zooChild',))
    zooChild = ZeroOrOne('p:zooChild', successors=())
    optAttr = OptionalAttribute('p:optAttr', ST_IntegerType)
    reqAttr = RequiredAttribute('reqAttr', ST_IntegerType)


class CT_Choice(BaseOxmlElement):
    """
    ``<p:choice>`` element
    """


class CT_OomChild(BaseOxmlElement):
    """
    Oom standing for 'OneOrMore', ``<p:oomChild>`` element, representing a
    child element that can appear multiple times in sequence, but must appear
    at least once.
    """


class CT_ZomChild(BaseOxmlElement):
    """
    Zom standing for 'ZeroOrMore', ``<p:zomChild>`` element, representing an
    optional child element that can appear multiple times in sequence.
    """


class CT_ZooChild(BaseOxmlElement):
    """
    Zoo standing for 'ZeroOrOne', ``<p:zooChild>`` element, an invented
    element for use in testing.
    """


register_element_cls('p:parent',   CT_Parent)
register_element_cls('p:choice',   CT_Choice)
register_element_cls('p:oomChild', CT_OomChild)
register_element_cls('p:zomChild', CT_ZomChild)
register_element_cls('p:zooChild', CT_ZooChild)


class CT_ChoiceBuilder(BaseBuilder):
    __tag__ = 'p:choice'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_Choice2Builder(BaseBuilder):
    __tag__ = 'p:choice2'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_ParentBuilder(BaseBuilder):
    __tag__ = 'p:parent'
    __nspfxs__ = ('p',)
    __attrs__ = ('p:optAttr', 'reqAttr')


class CT_OomChildBuilder(BaseBuilder):
    __tag__ = 'p:oomChild'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_OooChildBuilder(BaseBuilder):
    __tag__ = 'p:oooChild'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_ZomChildBuilder(BaseBuilder):
    __tag__ = 'p:zomChild'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_ZooChildBuilder(BaseBuilder):
    __tag__ = 'p:zooChild'
    __nspfxs__ = ('p',)
    __attrs__ = ()


def a_choice():
    return CT_ChoiceBuilder()


def a_choice2():
    return CT_Choice2Builder()


def a_parent():
    return CT_ParentBuilder()


def a_zomChild():
    return CT_ZomChildBuilder()


def a_zooChild():
    return CT_ZooChildBuilder()


def an_oomChild():
    return CT_OomChildBuilder()


def an_oooChild():
    return CT_OooChildBuilder()
