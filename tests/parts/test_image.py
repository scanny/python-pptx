# encoding: utf-8

"""Unit test suite for pptx.parts.image module."""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.compat import BytesIO
from pptx.package import Package
from pptx.parts.image import Image, ImagePart
from pptx.util import Emu

from ..unitutil.file import absjoin, test_file_dir
from ..unitutil.mock import (
    class_mock,
    initializer_mock,
    instance_mock,
    method_mock,
    property_mock,
)


images_pptx_path = absjoin(test_file_dir, "with_images.pptx")

test_image_path = absjoin(test_file_dir, "python-icon.jpeg")
test_eps_path = absjoin(test_file_dir, "cdw-logo.eps")
new_image_path = absjoin(test_file_dir, "monty-truth.png")


class DescribeImagePart(object):
    def it_can_construct_from_an_image_object(self, new_fixture):
        package_, image_, _init_, partname_ = new_fixture

        image_part = ImagePart.new(package_, image_)

        package_.next_image_partname.assert_called_once_with(image_.ext)
        _init_.assert_called_once_with(
            partname_, image_.content_type, image_.blob, package_, image_.filename
        )
        assert isinstance(image_part, ImagePart)

    def it_provides_access_to_its_image(self, image_fixture):
        image_part, Image_, blob, desc, image_ = image_fixture
        image = image_part.image
        Image_.assert_called_once_with(blob, desc)
        assert image is image_

    def it_can_scale_its_dimensions(self, scale_fixture):
        image_part, width, height, expected_values = scale_fixture
        assert image_part.scale(width, height) == expected_values

    def it_knows_its_pixel_dimensions(self, size_fixture):
        image, expected_size = size_fixture
        assert image._px_size == expected_size

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def image_fixture(self, Image_, image_):
        blob, filename = "blob", "foobar.png"
        image_part = ImagePart(None, None, blob, None, filename)
        return image_part, Image_, blob, filename, image_

    @pytest.fixture
    def new_fixture(self, request, package_, image_, _init_):
        partname_ = package_.next_image_partname.return_value
        return package_, image_, _init_, partname_

    @pytest.fixture(
        params=[
            (None, None, Emu(2590800), Emu(2590800)),
            (1000, None, 1000, 1000),
            (None, 3000, 3000, 3000),
            (3337, 9999, 3337, 9999),
        ]
    )
    def scale_fixture(self, request):
        width, height, expected_width, expected_height = request.param
        with open(test_image_path, "rb") as f:
            blob = f.read()
        image = ImagePart(None, None, blob, None)
        return image, width, height, (expected_width, expected_height)

    @pytest.fixture
    def size_fixture(self):
        with open(test_image_path, "rb") as f:
            blob = f.read()
        image = ImagePart(None, None, blob, None)
        return image, (204, 204)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Image_(self, request, image_):
        return class_mock(request, "pptx.parts.image.Image", return_value=image_)

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, ImagePart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)


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

    def it_can_construct_from_a_blob(self, from_blob_fixture):
        blob, filename = from_blob_fixture
        image = Image.from_blob(blob, filename)
        Image.__init__.assert_called_once_with(blob, filename)
        assert isinstance(image, Image)

    def it_knows_its_blob(self, blob_fixture):
        image, expected_value = blob_fixture
        assert image.blob == expected_value

    def it_knows_its_content_type(self, content_type_fixture):
        image, expected_value = content_type_fixture
        assert image.content_type == expected_value

    def it_knows_its_canonical_filename_extension(self, ext_fixture):
        image, expected_value = ext_fixture
        assert image.ext == expected_value

    def it_knows_its_dpi(self, dpi_fixture):
        image, expected_value = dpi_fixture
        assert image.dpi == expected_value

    def it_knows_its_filename(self, filename_fixture):
        image, expected_value = filename_fixture
        assert image.filename == expected_value

    def it_knows_its_sha1_hash(self):
        image = Image(b"foobar", None)
        assert image.sha1 == "8843d7f92416211de9ebb963ff4ce28125932878"

    def it_knows_its_PIL_properties_to_help(self, pil_fixture):
        image, size, format, dpi = pil_fixture
        assert image.size == size
        assert image._format == format
        assert image.dpi == dpi
        assert image._pil_props == (format, size, None)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def blob_fixture(self):
        blob = b"foobar"
        image = Image(blob, None)
        return image, blob

    @pytest.fixture(
        params=[
            ("BMP", "image/bmp"),
            ("GIF", "image/gif"),
            ("JPEG", "image/jpeg"),
            ("PNG", "image/png"),
            ("TIFF", "image/tiff"),
            ("WMF", "image/x-wmf"),
        ]
    )
    def content_type_fixture(self, request, _format_):
        format, expected_value = request.param
        image = Image(None, None)
        _format_.return_value = format
        return image, expected_value

    @pytest.fixture(
        params=[
            ((42, 24), (42, 24)),
            ((42.1, 23.6), (42, 24)),
            (None, (72, 72)),
            ((3047, 2388), (72, 72)),
            ("foobar", (72, 72)),
        ]
    )
    def dpi_fixture(self, request, _pil_props_):
        raw_dpi, expected_dpi = request.param
        image = Image(None, None)
        _pil_props_.return_value = (None, None, raw_dpi)
        return image, expected_dpi

    @pytest.fixture(
        params=[
            ("BMP", "bmp"),
            ("GIF", "gif"),
            ("JPEG", "jpg"),
            ("PNG", "png"),
            ("TIFF", "tiff"),
            ("WMF", "wmf"),
        ]
    )
    def ext_fixture(self, request, _format_):
        format, expected_value = request.param
        image = Image(None, None)
        _format_.return_value = format
        return image, expected_value

    @pytest.fixture(params=["foo.bar", None])
    def filename_fixture(self, request):
        filename = request.param
        image = Image(None, filename)
        return image, filename

    @pytest.fixture
    def from_blob_fixture(self, _init_):
        blob, filename = b"foobar", "foo.png"
        return blob, filename

    @pytest.fixture
    def from_path_fixture(self, from_blob_, image_):
        image_file = test_image_path
        with open(test_image_path, "rb") as f:
            blob = f.read()
        filename = "python-icon.jpeg"
        from_blob_.return_value = image_
        return image_file, blob, filename, image_

    @pytest.fixture
    def from_stream_fixture(self, from_blob_, image_):
        with open(test_image_path, "rb") as f:
            blob = f.read()
            image_file = BytesIO(blob)
        from_blob_.return_value = image_
        return image_file, blob, image_

    @pytest.fixture
    def pil_fixture(self):
        with open(test_image_path, "rb") as f:
            blob = f.read()
        image = Image(blob, None)
        size = (204, 204)
        format = "JPEG"
        dpi = (72, 72)
        return image, size, format, dpi

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _format_(self, request):
        return property_mock(request, Image, "_format")

    @pytest.fixture
    def from_blob_(self, request):
        return method_mock(request, Image, "from_blob")

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, Image)

    @pytest.fixture
    def _pil_props_(self, request):
        return property_mock(request, Image, "_pil_props")
