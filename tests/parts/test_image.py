# encoding: utf-8

"""
Test suite for pptx.parts.image module.
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from StringIO import StringIO

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.parts.image import Image, ImagePart
from pptx.package import Package
from pptx.util import Px

from ..unitutil.file import absjoin, test_file_dir
from ..unitutil.mock import instance_mock, method_mock


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
        assert image_part.sha1 == sha1
        assert image_part._desc == filename

    def it_can_scale_its_dimensions(self, scale_fixture):
        image, width, height, expected_values = scale_fixture
        assert image.scale(width, height) == expected_values

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


class DescribeImage(object):

    def it_can_construct_from_a_path(self, from_path_fixture):
        image_file, blob, filename, image_ = from_path_fixture
        image = Image.from_file(image_file)
        Image.from_blob.assert_called_once_with(blob, filename)
        assert image is image_

    def it_can_construct_from_a_stream(self, from_stream_fixture):
        image_file, blob, image_ = from_stream_fixture
        image = Image.from_file(image_file)
        Image.from_blob.assert_called_once_with(blob, None)
        assert image is image_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def from_path_fixture(self, from_blob_, image_):
        image_file = test_image_path
        with open(test_image_path, 'rb') as f:
            blob = f.read()
        filename = 'python-icon.jpeg'
        from_blob_.return_value = image_
        return image_file, blob, filename, image_

    @pytest.fixture
    def from_stream_fixture(self, from_blob_, image_):
        with open(test_image_path, 'rb') as f:
            blob = f.read()
            image_file = StringIO(blob)
        from_blob_.return_value = image_
        return image_file, blob, image_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def from_blob_(self, request):
        return method_mock(request, Image, 'from_blob')

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)


class DescribeImageCollection(object):

    def it_finds_a_matching_image_part_if_there_is_one(self):
        pkg = Package.open(images_pptx_path)
        matching_idx = 4
        matching_image = pkg._images[matching_idx]
        # exercise ---------------------
        image = pkg._images.add_image(test_image_path)
        # verify -----------------------
        msg = ("expected images[%d], got images[%d]"
               % (matching_idx, pkg._images.index(image)))
        assert image is matching_image, msg

    def it_adds_an_image_part_if_no_matching_found(self):
        pkg = Package.open(images_pptx_path)
        expected_partname = '/ppt/media/image8.png'
        expected_len = len(pkg._images) + 1
        expected_sha1 = '79769f1e202add2e963158b532e36c2c0f76a70c'
        # exercise ---------------------
        image = pkg._images.add_image(new_image_path)
        # verify -----------------------
        expected = (expected_partname, expected_len, expected_sha1)
        actual = (image.partname, len(pkg._images), image.sha1)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        assert actual == expected, msg
