# encoding: utf-8

"""
Test suite for pptx.parts.image module.
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from StringIO import StringIO

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.parts.image import ImagePart
from pptx.package import Package
from pptx.util import Px

from ..unitutil.file import absjoin, test_file_dir
from ..unitutil.legacy import TestCase


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
test_eps_path = absjoin(test_file_dir, 'cdw-logo.eps')
new_image_path = absjoin(test_file_dir, 'monty-truth.png')


class DescribeImagePart(object):

    def it_can_construct_from_an_image_file(self, new_fixture):
        """
        Integration test over complete contruction from file.
        """
        partname, image_file, content_type, ext, sha1, filename = new_fixture

        image_part = ImagePart.new(partname, image_file)

        assert image_part.partname == partname
        assert image_part.content_type == content_type
        assert image_part.ext == ext
        assert image_part._sha1 == sha1
        assert image_part._desc == filename

    def it_can_scale_its_dimensions(self, scale_fixture):
        image, width, height, expected_values = scale_fixture
        assert image._scale(width, height) == expected_values

    def it_knows_its_pixel_dimensions(self, size_fixture):
        image, expected_size = size_fixture
        assert image._size == expected_size

    def it_knows_its_image_content_type(self):
        content_type = ImagePart._image_ext_content_type('jPeG')
        assert content_type == CT.JPEG

    def it_raises_on_unsupported_image_stream_type(self):
        with pytest.raises(ValueError):
            with open(test_eps_path) as stream:
                ImagePart._ext_from_image_stream(stream)

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (False, 'jpeg', 'python-icon.jpeg'),
        (True,  'jpg',  'image.jpg'),
    ])
    def new_fixture(self, request):

        def image_stream():
            with open(test_image_path, 'rb') as f:
                return StringIO(f.read())

        use_stream, ext, filename = request.param
        image_file = image_stream() if use_stream else test_image_path
        partname = '/ppt/media/image1.png'
        content_type = CT.JPEG
        sha1 = '1be010ea47803b00e140b852765cdf84f491da47'
        return partname, image_file, content_type, ext, sha1, filename

    @pytest.fixture(params=[
        (None, None, Px(204), Px(204)),
        (1000, None, 1000,    1000),
        (None, 3000, 3000,    3000),
        (3337, 9999, 3337,    9999),
    ])
    def scale_fixture(self, request):
        width, height, expected_width, expected_height = request.param
        image = ImagePart.new(None, test_image_path)
        return image, width, height, (expected_width, expected_height)

    @pytest.fixture
    def size_fixture(self):
        image = ImagePart.new(None, test_image_path)
        return image, (204, 204)


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
