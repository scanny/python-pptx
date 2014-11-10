# encoding: utf-8

"""
Test suite for pptx.parts.image module.
"""

from __future__ import absolute_import, print_function

from StringIO import StringIO

from pptx.opc.packuri import PackURI
from pptx.parts.image import ImagePart
from pptx.package import Package
from pptx.util import Px

from ..unitutil.file import absjoin, test_file_dir
from ..unitutil.legacy import TestCase


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
test_eps_path = absjoin(test_file_dir, 'cdw-logo.eps')
new_image_path = absjoin(test_file_dir, 'monty-truth.png')


class TestImagePart(TestCase):
    """Test ImagePart"""
    def test_construction_from_file(self):
        """ImagePart(path) constructor produces correct attribute values"""
        # exercise ---------------------
        partname = PackURI('/ppt/media/image1.jpeg')
        image = ImagePart.new(partname, test_image_path)
        # verify -----------------------
        assert image.ext == 'jpeg'
        assert image.content_type == 'image/jpeg'
        assert len(image._blob) == 3277
        assert image._desc == 'python-icon.jpeg'

    def test_construction_from_stream(self):
        """ImagePart(stream) construction produces correct attribute values"""
        # exercise ---------------------
        partname = PackURI('/ppt/media/image1.jpeg')
        with open(test_image_path, 'rb') as f:
            stream = StringIO(f.read())
        image = ImagePart.new(partname, stream)
        # verify -----------------------
        assert image.ext == 'jpg'
        assert image.content_type == 'image/jpeg'
        assert len(image._blob) == 3277
        assert image._desc == 'image.jpg'

    def test_construction_from_file_raises_on_bad_path(self):
        """ImagePart(path) constructor raises on bad path"""
        partname = PackURI('/ppt/media/image1.jpeg')
        with self.assertRaises(IOError):
            ImagePart.new(partname, 'foobar27.png')

    def test__scale_calculates_correct_dimensions(self):
        """ImagePart._scale() calculates correct dimensions"""
        # setup ------------------------
        test_cases = (
            ((None, None), (Px(204), Px(204))),
            ((1000, None), (1000, 1000)),
            ((None, 3000), (3000, 3000)),
            ((3337, 9999), (3337, 9999)))
        partname = PackURI('/ppt/media/image1.png')
        image = ImagePart.new(partname, test_image_path)
        # verify -----------------------
        for params, expected in test_cases:
            width, height = params
            assert image._scale(width, height) == expected

    def test__size_returns_image_native_pixel_dimensions(self):
        """ImagePart._size is width, height tuple of image pixel dimensions"""
        partname = PackURI('/ppt/media/image1.png')
        image = ImagePart.new(partname, test_image_path)
        assert image._size == (204, 204)

    def test__ext_from_image_stream_raises_on_incompatible_format(self):
        with self.assertRaises(ValueError):
            with open(test_eps_path) as stream:
                ImagePart._ext_from_image_stream(stream)

    def test__image_ext_content_type_known_type(self):
        """
        ImagePart._image_ext_content_type() correct for known content type
        """
        # exercise ---------------------
        content_type = ImagePart._image_ext_content_type('jPeG')
        # verify -----------------------
        expected = 'image/jpeg'
        actual = content_type
        msg = ("expected content type '%s', got '%s'" % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test__image_ext_content_type_raises_on_bad_ext(self):
        with self.assertRaises(ValueError):
            ImagePart._image_ext_content_type('xj7')


class TestImageCollection(TestCase):
    """Test ImageCollection"""
    def test_add_image_returns_matching_image(self):
        pkg = Package.open(images_pptx_path)
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
        pkg = Package.open(images_pptx_path)
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
