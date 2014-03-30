# encoding: utf-8

"""Test suite for pptx.shape module."""

from __future__ import absolute_import

import pytest

from hamcrest import assert_that, is_

from pptx.oxml.ns import namespaces
from pptx.parts.slide import Slide
from pptx.shapes import Subshape
from pptx.shapes.shape import BaseShape

from ..unitutil import (
    absjoin, loose_mock, parse_xml_file, TestCase, test_file_dir
)


slide1_path = absjoin(test_file_dir, 'slide1.xml')

nsmap = namespaces('a', 'r', 'p')


class DescribeBaseShape(object):

    def it_knows_the_part_it_belongs_to(self, shape_with_parent_):
        shape, parent_ = shape_with_parent_
        part = shape.part
        assert part is parent_.part

    # fixtures ---------------------------------------------

    @pytest.fixture
    def shape_with_parent_(self, request):
        parent_ = loose_mock(request, name='parent_')
        shape = BaseShape(None, parent_)
        return shape, parent_


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

    def test_has_textframe_value(self):
        """
        BaseShape.has_textframe value correct
        """
        # setup ------------------------
        with open(slide1_path) as f:
            xml_bytes = f.read()
        slide = Slide.load(None, None, xml_bytes, None)
        shapes = slide.shapes
        indexes = []
        # exercise ---------------------
        for idx, shape in enumerate(shapes):
            if shape.has_textframe:
                indexes.append(idx)
        # verify -----------------------
        expected = [0, 1, 3, 5, 6]
        actual = indexes
        msg = ("expected txBody element in shapes %s, got %s" %
               (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test_id_value(self):
        """BaseShape.id value is correct"""
        # exercise ---------------------
        id = self.base_shape.id
        # verify -----------------------
        expected = 6
        actual = id
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

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
