# encoding: utf-8

"""Unit-test suite for `pptx.parts.image` module."""

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
    """Unit-test suite for `pptx.parts.image.ImagePart` objects."""

    def it_can_construct_from_an_image_object(self, request, image_):
        package_ = instance_mock(request, Package)
        _init_ = initializer_mock(request, ImagePart)
        partname_ = package_.next_image_partname.return_value

        image_part = ImagePart.new(package_, image_)

        package_.next_image_partname.assert_called_once_with(image_.ext)
        _init_.assert_called_once_with(
            image_part,
            partname_,
            image_.content_type,
            package_,
            image_.blob,
            image_.filename,
        )
        assert isinstance(image_part, ImagePart)

    def it_provides_access_to_its_image(self, request, image_):
        Image_ = class_mock(request, "pptx.parts.image.Image")
        Image_.return_value = image_
        property_mock(request, ImagePart, "desc", return_value="foobar.png")
        image_part = ImagePart(None, None, None, b"blob", None)

        image = image_part.image

        Image_.assert_called_once_with(b"blob", "foobar.png")
        assert image is image_

    @pytest.mark.parametrize(
        "width, height, expected_width, expected_height",
        (
            (None, None, Emu(2590800), Emu(2590800)),
            (1000, None, 1000, 1000),
            (None, 3000, 3000, 3000),
            (3337, 9999, 3337, 9999),
        ),
    )
    def it_can_scale_its_dimensions(
        self, width, height, expected_width, expected_height
    ):
        with open(test_image_path, "rb") as f:
            blob = f.read()
        image_part = ImagePart(None, None, None, blob)

        assert image_part.scale(width, height) == (expected_width, expected_height)

    def it_knows_its_pixel_dimensions_to_help(self):
        with open(test_image_path, "rb") as f:
            blob = f.read()
        image_part = ImagePart(None, None, None, blob)

        assert image_part._px_size == (204, 204)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)


class DescribeImage(object):
    """Unit-test suite for `pptx.parts.image.Image` objects."""

    def it_can_construct_from_a_path(self, from_blob_, image_):
        with open(test_image_path, "rb") as f:
            blob = f.read()
        from_blob_.return_value = image_

        image = Image.from_file(test_image_path)

        Image.from_blob.assert_called_once_with(blob, "python-icon.jpeg")
        assert image is image_

    def it_can_construct_from_a_stream(self, from_stream_fixture):
        image_file, blob, image_ = from_stream_fixture
        image = Image.from_file(image_file)
        Image.from_blob.assert_called_once_with(blob, None)
        assert image is image_

    def it_can_construct_from_a_blob(self, _init_):
        image = Image.from_blob(b"blob", "foo.png")

        _init_.assert_called_once_with(image, b"blob", "foo.png")
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
        return method_mock(request, Image, "from_blob", autospec=False)

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, Image)

    @pytest.fixture
    def _pil_props_(self, request):
        return property_mock(request, Image, "_pil_props")
