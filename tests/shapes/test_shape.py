# encoding: utf-8

"""Test suite for pptx.shape module."""

from __future__ import absolute_import

import pytest

from hamcrest import assert_that, is_

from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.text import CT_TextBody
from pptx.oxml.ns import namespaces
from pptx.parts.slide import _SlideShapeTree
from pptx.shapes import Subshape
from pptx.shapes.shape import BaseShape

from ..unitutil import (
    absjoin, instance_mock, loose_mock, parse_xml_file, TestCase, test_file_dir
)


slide1_path = absjoin(test_file_dir, 'slide1.xml')

nsmap = namespaces('a', 'r', 'p')


class DescribeBaseShape(object):

    def it_knows_its_shape_id(self, id_fixture):
        shape, shape_id = id_fixture
        assert shape.id == shape_id

    def it_knows_the_part_it_belongs_to(self, part_fixture):
        shape, parent_ = part_fixture
        part = shape.part
        assert part is parent_.part

    def it_knows_whether_it_can_contain_text(self, has_textframe_fixture):
        shape, has_textframe = has_textframe_fixture
        assert shape.has_textframe is has_textframe

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def id_fixture(self, shape_elm_, shape_id):
        shape = BaseShape(shape_elm_, None)
        return shape, shape_id

    @pytest.fixture(params=[True, False])
    def has_textframe_fixture(self, request, shape_elm_, txBody_):
        has_textframe = request.param
        shape_elm_.txBody = txBody_ if has_textframe else None
        shape = BaseShape(shape_elm_, None)
        return shape, has_textframe

    @pytest.fixture
    def part_fixture(self, shapes_):
        parent_ = shapes_
        shape = BaseShape(None, parent_)
        return shape, parent_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def shape_elm_(self, request, shape_id):
        return instance_mock(request, BaseShapeElement, shape_id=shape_id)

    @pytest.fixture
    def shape_id(self):
        return 42

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, _SlideShapeTree)

    @pytest.fixture
    def txBody_(self, request):
        return instance_mock(request, CT_TextBody)


class DescribeSubshape(object):

    def it_knows_the_part_it_belongs_to(self, subshape_with_parent_):
        subshape, parent_ = subshape_with_parent_
        part = subshape.part
        assert part is parent_.part

    # fixtures ---------------------------------------------

    @pytest.fixture
    def subshape_with_parent_(self, request):
        parent_ = loose_mock(request, name='parent_')
        subshape = Subshape(parent_)
        return subshape, parent_


class TestBaseShape(TestCase):
    """Test BaseShape"""
    def setUp(self):
        path = slide1_path
        self.sld = parse_xml_file(path).getroot()
        xpath = './p:cSld/p:spTree/p:pic'
        pic = self.sld.xpath(xpath, namespaces=nsmap)[0]
        self.base_shape = BaseShape(pic, None)

    def test_is_placeholder_true_for_placeholder(self):
        """BaseShape.is_placeholder True for placeholder shape"""
        # setup ------------------------
        xpath = './p:cSld/p:spTree/p:sp'
        sp = self.sld.xpath(xpath, namespaces=nsmap)[0]
        base_shape = BaseShape(sp, None)
        # verify -----------------------
        actual = base_shape.is_placeholder
        msg = "expected True, got %s" % (actual)
        self.assertTrue(actual, msg)

    def test_is_placeholder_false_for_non_placeholder(self):
        """BaseShape.is_placeholder False for non-placeholder shape"""
        # verify -----------------------
        actual = self.base_shape.is_placeholder
        msg = "expected False, got %s" % (actual)
        self.assertFalse(actual, msg)

    def test__is_title_true_for_title_placeholder(self):
        """BaseShape._is_title True for title placeholder shape"""
        # setup ------------------------
        xpath = './p:cSld/p:spTree/p:sp'
        title_placeholder_sp = self.sld.xpath(xpath, namespaces=nsmap)[0]
        base_shape = BaseShape(title_placeholder_sp, None)
        # verify -----------------------
        actual = base_shape._is_title
        msg = "expected True, got %s" % (actual)
        self.assertTrue(actual, msg)

    # # needs a shape XML builder, this early version made too many assumptions
    # def test__is_title_false_for_no_ph_element(self):
    #     """BaseShape._is_title False on shape has no <p:ph> element"""
    #     # setup ------------------------
    #     self.base_shape._element = Mock(name='_element')
    #     self.base_shape._element.xpath.return_value = []
    #     # verify -----------------------
    #     assert_that(self.base_shape._is_title, is_(False))

    def test_name_value(self):
        """BaseShape.name value is correct"""
        # exercise ---------------------
        name = self.base_shape.name
        # verify -----------------------
        expected = 'Picture 5'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_shape_name_returns_none_for_unimplemented_shape_types(self):
        """BaseShape.shape_name returns None for unimplemented shape types"""
        assert_that(self.base_shape.shape_type, is_(None))

    def test_textframe_raises_on_no_textframe(self):
        """BaseShape.textframe raises on shape with no text frame"""
        with self.assertRaises(ValueError):
            self.base_shape.textframe

    def test_text_setter_structure_and_value(self):
        """assign to BaseShape.text yields single run para set to value"""
        # setup ------------------------
        test_text = 'python-pptx was here!!'
        xpath = './p:cSld/p:spTree/p:sp'
        textbox_sp = self.sld.xpath(xpath, namespaces=nsmap)[2]
        base_shape = BaseShape(textbox_sp, None)
        # exercise ---------------------
        base_shape.text = test_text
        # verify paragraph count ------
        expected = 1
        actual = len(base_shape.textframe.paragraphs)
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value ----------------
        expected = test_text
        actual = base_shape.textframe.paragraphs[0].runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_text_setter_raises_on_no_textframe(self):
        """assignment to BaseShape.text raises for shape with no text frame"""
        with self.assertRaises(TypeError):
            self.base_shape.text = 'test text'
