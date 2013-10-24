# encoding: utf-8

"""Test suite for pptx.image module."""

from __future__ import absolute_import

import os

from StringIO import StringIO

from hamcrest import assert_that, equal_to, is_

from pptx.parts.image import _Image
from pptx.presentation import _Package
from pptx.util import Px

from ..unitutil import TestCase


def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))


thisdir = os.path.split(__file__)[0]
test_file_dir = absjoin(thisdir, '../test_files')

images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

test_bmp_path = absjoin(test_file_dir, 'python.bmp')
test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
new_image_path = absjoin(test_file_dir, 'monty-truth.png')


class Test_Image(TestCase):
    """Test _Image"""
    def test_construction_from_file(self):
        """_Image(path) constructor produces correct attribute values"""
        # exercise ---------------------
        image = _Image(test_image_path)
        # verify -----------------------
        assert_that(image.ext, is_(equal_to('.jpeg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))
        assert_that(image._desc, is_(equal_to('python-icon.jpeg')))

    def test_construction_from_stream(self):
        """_Image(stream) construction produces correct attribute values"""
        # exercise ---------------------
        with open(test_image_path, 'rb') as f:
            stream = StringIO(f.read())
        image = _Image(stream)
        # verify -----------------------
        assert_that(image.ext, is_(equal_to('.jpg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))
        assert_that(image._desc, is_(equal_to('image.jpg')))

    def test_construction_from_file_raises_on_bad_path(self):
        """_Image(path) constructor raises on bad path"""
        # verify -----------------------
        with self.assertRaises(IOError):
            _Image('foobar27.png')

    def test__scale_calculates_correct_dimensions(self):
        """_Image._scale() calculates correct dimensions"""
        # setup ------------------------
        test_cases = (
            ((None, None), (Px(204), Px(204))),
            ((1000, None), (1000, 1000)),
            ((None, 3000), (3000, 3000)),
            ((3337, 9999), (3337, 9999)))
        image = _Image(test_image_path)
        # verify -----------------------
        for params, expected in test_cases:
            width, height = params
            assert_that(image._scale(width, height), is_(equal_to(expected)))

    def test__size_returns_image_native_pixel_dimensions(self):
        """_Image._size is width, height tuple of image pixel dimensions"""
        image = _Image(test_image_path)
        assert_that(image._size, is_(equal_to((204, 204))))

    def test___ext_from_image_stream_raises_on_incompatible_format(self):
        """_Image.__ext_from_image_stream() raises on incompatible format"""
        # verify -----------------------
        with self.assertRaises(ValueError):
            with open(test_bmp_path) as stream:
                _Image._Image__ext_from_image_stream(stream)

    def test___image_ext_content_type_known_type(self):
        """_Image.__image_ext_content_type() correct for known content type"""
        # exercise ---------------------
        content_type = _Image._Image__image_ext_content_type('.jpeg')
        # verify -----------------------
        expected = 'image/jpeg'
        actual = content_type
        msg = ("expected content type '%s', got '%s'" % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test___image_ext_content_type_raises_on_bad_ext(self):
        """_Image.__image_ext_content_type() raises on bad extension"""
        # verify -----------------------
        with self.assertRaises(TypeError):
            _Image._Image__image_ext_content_type('.xj7')

    def test___image_ext_content_type_raises_on_non_img_ext(self):
        """_Image.__image_ext_content_type() raises on non-image extension"""
        # verify -----------------------
        with self.assertRaises(TypeError):
            _Image._Image__image_ext_content_type('.xml')


class Test_ImageCollection(TestCase):
    """Test _ImageCollection"""
    def test_add_image_returns_matching_image(self):
        """_ImageCollection.add_image() returns existing image on match"""
        # setup ------------------------
        pkg = _Package(images_pptx_path)
        matching_idx = 4
        matching_image = pkg._images[matching_idx]
        # exercise ---------------------
        image = pkg._images.add_image(test_image_path)
        # verify -----------------------
        expected = matching_image
        actual = image
        msg = ("expected images[%d], got images[%d]"
               % (matching_idx, pkg._images.index(image)))
        self.assertEqual(expected, actual, msg)

    def test_add_image_adds_new_image(self):
        """_ImageCollection.add_image() adds new image on no match"""
        # setup ------------------------
        pkg = _Package(images_pptx_path)
        expected_partname = '/ppt/media/image8.png'
        expected_len = len(pkg._images) + 1
        expected_sha1 = '79769f1e202add2e963158b532e36c2c0f76a70c'
        # exercise ---------------------
        image = pkg._images.add_image(new_image_path)
        # verify -----------------------
        expected = (expected_partname, expected_len, expected_sha1)
        actual = (image.partname, len(pkg._images), image._sha1)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
