# encoding: utf-8

"""Test suite for pptx.image module."""

from __future__ import absolute_import

from StringIO import StringIO

from hamcrest import assert_that, equal_to, is_

from pptx.opc.packuri import PackURI
from pptx.parts.image import Image
from pptx.presentation import Package
from pptx.util import Px

from ..unitutil import absjoin, TestCase, test_file_dir


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

test_bmp_path = absjoin(test_file_dir, 'python.bmp')
test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
new_image_path = absjoin(test_file_dir, 'monty-truth.png')


class TestImage(TestCase):
    """Test Image"""
    def test_construction_from_file(self):
        """Image(path) constructor produces correct attribute values"""
        # exercise ---------------------
        partname = PackURI('/ppt/media/image1.jpeg')
        image = Image.new(partname, test_image_path)
        # verify -----------------------
        assert_that(image.ext, is_(equal_to('.jpeg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))
        assert_that(image._desc, is_(equal_to('python-icon.jpeg')))

    def test_construction_from_stream(self):
        """Image(stream) construction produces correct attribute values"""
        # exercise ---------------------
        partname = PackURI('/ppt/media/image1.jpeg')
        with open(test_image_path, 'rb') as f:
            stream = StringIO(f.read())
        image = Image.new(partname, stream)
        # verify -----------------------
        assert_that(image.ext, is_(equal_to('.jpg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))
        assert_that(image._desc, is_(equal_to('image.jpg')))

    def test_construction_from_file_raises_on_bad_path(self):
        """Image(path) constructor raises on bad path"""
        partname = PackURI('/ppt/media/image1.jpeg')
        with self.assertRaises(IOError):
            Image.new(partname, 'foobar27.png')

    def test__scale_calculates_correct_dimensions(self):
        """Image._scale() calculates correct dimensions"""
        # setup ------------------------
        test_cases = (
            ((None, None), (Px(204), Px(204))),
            ((1000, None), (1000, 1000)),
            ((None, 3000), (3000, 3000)),
            ((3337, 9999), (3337, 9999)))
        partname = PackURI('/ppt/media/image1.png')
        image = Image.new(partname, test_image_path)
        # verify -----------------------
        for params, expected in test_cases:
            width, height = params
            assert_that(image._scale(width, height), is_(equal_to(expected)))

    def test__size_returns_image_native_pixel_dimensions(self):
        """Image._size is width, height tuple of image pixel dimensions"""
        partname = PackURI('/ppt/media/image1.png')
        image = Image.new(partname, test_image_path)
        assert_that(image._size, is_(equal_to((204, 204))))

    def test__ext_from_image_stream_raises_on_incompatible_format(self):
        with self.assertRaises(ValueError):
            with open(test_bmp_path) as stream:
                Image._ext_from_image_stream(stream)

    def test__image_ext_content_type_known_type(self):
        """Image._image_ext_content_type() correct for known content type"""
        # exercise ---------------------
        content_type = Image._image_ext_content_type('.jpeg')
        # verify -----------------------
        expected = 'image/jpeg'
        actual = content_type
        msg = ("expected content type '%s', got '%s'" % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test__image_ext_content_type_raises_on_bad_ext(self):
        with self.assertRaises(TypeError):
            Image._image_ext_content_type('.xj7')

    def test__image_ext_content_type_raises_on_non_img_ext(self):
        with self.assertRaises(TypeError):
            Image._image_ext_content_type('.xml')


class TestImageCollection(TestCase):
    """Test ImageCollection"""
    def test_add_image_returns_matching_image(self):
        """ImageCollection.add_image() returns existing image on match"""
        # setup ------------------------
        pkg = Package(images_pptx_path)
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
        """ImageCollection.add_image() adds new image on no match"""
        # setup ------------------------
        pkg = Package(images_pptx_path)
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
