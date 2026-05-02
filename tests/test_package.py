# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.package` module."""

from __future__ import annotations

import os

import pytest

import io

import pptx
from pptx.exc import UnsupportedImageTypeError
from pptx.media import Video
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, _Relationship
from pptx.opc.packuri import PackURI
from pptx.package import (
    Package,
    _ImageParts,
    _looks_like_svg,
    _MediaParts,
    _raise_if_svg,
)
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.extprops import ExtendedPropertiesPart
from pptx.parts.image import Image, ImagePart
from pptx.parts.media import MediaPart

from .unitutil.mock import call, class_mock, instance_mock, method_mock, property_mock


class DescribePackage(object):
    """Unit-test suite for `pptx.package.Package` objects."""

    def it_provides_access_to_its_core_properties_part(self):
        default_pptx = os.path.abspath(
            os.path.join(os.path.split(pptx.__file__)[0], "templates", "default.pptx")
        )
        pkg = Package.open(default_pptx)
        assert isinstance(pkg.core_properties, CorePropertiesPart)

    def it_provides_access_to_its_extended_properties_part(self):
        """The default template already has an extended-properties part."""
        default_pptx = os.path.abspath(
            os.path.join(os.path.split(pptx.__file__)[0], "templates", "default.pptx")
        )
        pkg = Package.open(default_pptx)
        assert isinstance(pkg.extended_properties, ExtendedPropertiesPart)

    def it_creates_a_default_extended_properties_part_when_missing(self, request):
        """If the package has no extended-properties rel, one is created on demand."""
        package = Package(None)
        # -- short-circuit the lookup so it raises KeyError (no rel present) --
        method_mock(request, Package, "part_related_by", side_effect=KeyError)
        relate_to_ = method_mock(request, Package, "relate_to")

        ext_props = package.extended_properties

        assert isinstance(ext_props, ExtendedPropertiesPart)
        relate_to_.assert_called_once_with(package, ext_props, RT.EXTENDED_PROPERTIES)

    def it_triggers_lazy_ext_props_creation_before_save(self, request):
        """`Package.save()` triggers `extended_properties` so the part is included."""
        from pptx.opc.package import OpcPackage

        package = Package(None)
        ext_props_prop_ = property_mock(request, Package, "extended_properties")
        base_save_ = method_mock(request, OpcPackage, "save")
        pkg_file = "foobar.pptx"

        package.save(pkg_file)

        ext_props_prop_.assert_called_once_with()
        base_save_.assert_called_once_with(package, pkg_file, None)

    def it_can_get_or_add_an_image_part(self, image_part_fixture):
        package, image_file, image_parts_, image_part_ = image_part_fixture
        image_part = package.get_or_add_image_part(image_file)
        image_parts_.get_or_add_image_part.assert_called_once_with(image_file)
        assert image_part is image_part_

    def it_can_get_or_add_a_media_part(self, media_part_fixture):
        package, media, media_part_ = media_part_fixture
        media_part = package.get_or_add_media_part(media)
        package._media_parts.get_or_add_media_part.assert_called_once_with(media)
        assert media_part is media_part_

    def it_knows_the_next_available_image_partname(self, next_fixture):
        package, ext, expected_value = next_fixture
        partname = package.next_image_partname(ext)
        assert partname == expected_value

    def it_knows_the_next_available_media_partname(self, nmp_fixture):
        package, ext, expected_value = nmp_fixture
        partname = package.next_media_partname(ext)
        assert partname == expected_value

    def it_provides_access_to_its_MediaParts_object(self, m_parts_fixture):
        package, _MediaParts_, media_parts_ = m_parts_fixture
        media_parts = package._media_parts
        _MediaParts_.assert_called_once_with(package)
        assert media_parts is media_parts_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def image_part_fixture(self, image_parts_, image_part_, _image_parts_prop_):
        package = Package(None)
        image_file = "foobar.png"
        _image_parts_prop_.return_value = image_parts_
        image_parts_.get_or_add_image_part.return_value = image_part_
        return package, image_file, image_parts_, image_part_

    @pytest.fixture
    def media_part_fixture(self, media_, media_part_, _media_parts_prop_, media_parts_):
        package = Package(None)
        _media_parts_prop_.return_value = media_parts_
        media_parts_.get_or_add_media_part.return_value = media_part_
        return package, media_, media_part_

    @pytest.fixture
    def m_parts_fixture(self, _MediaParts_, media_parts_):
        package = Package(None)
        _MediaParts_.return_value = media_parts_
        return package, _MediaParts_, media_parts_

    @pytest.fixture(params=[((3, 4, 2), 1), ((4, 2, 1), 3), ((2, 3, 1), 4)])
    def next_fixture(self, request, iter_parts_):
        idxs, idx = request.param
        package = Package(None)
        package.iter_parts.return_value = self.i_image_parts(request, idxs)
        ext = "foo"
        expected_value = "/ppt/media/image%d.%s" % (idx, ext)
        return package, ext, expected_value

    @pytest.fixture(params=[((3, 4, 2), 1), ((4, 2, 1), 3), ((2, 3, 1), 4)])
    def nmp_fixture(self, request, iter_parts_):
        idxs, idx = request.param
        package = Package(None)
        package.iter_parts.return_value = self.i_media_parts(request, idxs)
        ext = "foo"
        expected_value = "/ppt/media/media%d.%s" % (idx, ext)
        return package, ext, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def image_parts_(self, request):
        return instance_mock(request, _ImageParts)

    @pytest.fixture
    def _image_parts_prop_(self, request):
        return property_mock(request, Package, "_image_parts")

    def i_image_parts(self, request, idxs):
        def part(idx):
            partname = PackURI("/ppt/media/image%d.png" % idx)
            return instance_mock(request, Part, partname=partname)

        return iter([part(idx) for idx in idxs])

    def i_media_parts(self, request, idxs):
        def part(idx):
            partname = PackURI("/ppt/media/media%d.mp4" % idx)
            return instance_mock(request, Part, partname=partname)

        return iter([part(idx) for idx in idxs])

    @pytest.fixture
    def iter_parts_(self, request):
        return property_mock(request, Package, "iter_parts")

    @pytest.fixture
    def media_(self, request):
        return instance_mock(request, Video)

    @pytest.fixture
    def media_part_(self, request):
        return instance_mock(request, MediaPart)

    @pytest.fixture
    def _MediaParts_(self, request):
        return class_mock(request, "pptx.package._MediaParts")

    @pytest.fixture
    def media_parts_(self, request):
        return instance_mock(request, _MediaParts)

    @pytest.fixture
    def _media_parts_prop_(self, request):
        return property_mock(request, Package, "_media_parts")


class Describe_ImageParts(object):
    """Unit-test suite for `pptx.package._ImageParts` objects."""

    def it_can_iterate_over_the_package_image_parts(self, iter_fixture):
        image_parts, expected_parts = iter_fixture
        assert list(image_parts) == expected_parts

    def it_can_get_a_matching_image_part(self, Image_, image_, image_part_, _find_by_sha1_):
        Image_.from_file.return_value = image_
        _find_by_sha1_.return_value = image_part_
        image_parts = _ImageParts(None)

        image_part = image_parts.get_or_add_image_part("image.png")

        Image_.from_file.assert_called_once_with("image.png")
        _find_by_sha1_.assert_called_once_with(image_parts, image_.sha1)
        assert image_part is image_part_

    def it_can_add_an_image_part(
        self, package_, Image_, image_, _find_by_sha1_, ImagePart_, image_part_
    ):
        Image_.from_file.return_value = image_
        _find_by_sha1_.return_value = None
        ImagePart_.new.return_value = image_part_
        image_parts = _ImageParts(package_)

        image_part = image_parts.get_or_add_image_part("image.png")

        Image_.from_file.assert_called_once_with("image.png")
        _find_by_sha1_.assert_called_once_with(image_parts, image_.sha1)
        ImagePart_.new.assert_called_once_with(package_, image_)
        assert image_part is image_part_

    def it_can_find_an_image_part_by_sha1_hash(self, find_fixture):
        image_parts, sha1, expected_value = find_fixture
        image_part = image_parts._find_by_sha1(sha1)
        assert image_part is expected_value

    def but_it_skips_unsupported_image_types(self, request, _iter_):
        sha1 = "f00beed"
        svg_part_ = instance_mock(request, Part, name="svg_part_")
        png_part_ = instance_mock(request, ImagePart, name="png_part_", sha1=sha1)
        # ---order iteration to encounter svg part before target part---
        _iter_.return_value = iter((svg_part_, png_part_))
        image_parts = _ImageParts(None)

        result = image_parts._find_by_sha1(sha1)

        assert result == png_part_

    def it_raises_UnsupportedImageTypeError_on_svg_path(self, tmp_path):
        svg_path = tmp_path / "logo.svg"
        svg_path.write_bytes(
            b'<?xml version="1.0"?>\n'
            b'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>'
        )
        image_parts = _ImageParts(None)

        with pytest.raises(UnsupportedImageTypeError) as exc_info:
            image_parts.get_or_add_image_part(str(svg_path))

        assert "SVG images are not supported" in str(exc_info.value)
        assert "logo.svg" in str(exc_info.value)
        assert "pre-rasterize" in str(exc_info.value).lower()

    def it_raises_UnsupportedImageTypeError_on_svg_stream(self):
        svg_stream = io.BytesIO(
            b'<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>'
        )
        image_parts = _ImageParts(None)

        with pytest.raises(UnsupportedImageTypeError):
            image_parts.get_or_add_image_part(svg_stream)

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[True, False])
    def find_fixture(self, request, _iter_, image_part_):
        image_part_is_present = request.param
        image_parts = _ImageParts(None)
        _iter_.return_value = iter((image_part_,))
        sha1 = "foobar"
        if image_part_is_present:
            image_part_.sha1 = "foobar"
            expected_value = image_part_
        else:
            image_part_.sha1 = "barfoo"
            expected_value = None
        return image_parts, sha1, expected_value

    @pytest.fixture
    def iter_fixture(self, request, package_):
        def rel(is_external, reltype):
            part = instance_mock(request, Part)
            return instance_mock(
                request,
                _Relationship,
                is_external=is_external,
                reltype=reltype,
                target_part=part,
            )

        rels = (rel(True, RT.IMAGE), rel(False, RT.SLIDE), rel(False, RT.IMAGE))

        package_.iter_rels.return_value = iter((rels[0], rels[1], rels[2], rels[2]))
        image_parts = _ImageParts(package_)
        expected_parts = [rels[2].target_part]
        return image_parts, expected_parts

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _find_by_sha1_(self, request):
        return method_mock(request, _ImageParts, "_find_by_sha1")

    @pytest.fixture
    def Image_(self, request):
        return class_mock(request, "pptx.package.Image")

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def ImagePart_(self, request):
        return class_mock(request, "pptx.package.ImagePart")

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def _iter_(self, request):
        return method_mock(request, _ImageParts, "__iter__")

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)


class Describe_raise_if_svg(object):
    """Unit-test suite for `pptx.package._raise_if_svg`."""

    def it_raises_on_a_minimal_svg_blob(self):
        stream = io.BytesIO(b'<svg xmlns="http://www.w3.org/2000/svg"/>')
        with pytest.raises(UnsupportedImageTypeError):
            _raise_if_svg(stream)

    def it_raises_on_svg_with_xml_prologue_and_doctype(self):
        blob = (
            b'<?xml version="1.0" encoding="UTF-8"?>\n'
            b"<!DOCTYPE svg PUBLIC '-//W3C//DTD SVG 1.1//EN' "
            b"'http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd'>\n"
            b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"/>'
        )
        stream = io.BytesIO(blob)
        with pytest.raises(UnsupportedImageTypeError):
            _raise_if_svg(stream)

    def it_raises_on_uppercase_svg_tag(self):
        stream = io.BytesIO(b'<?xml version="1.0"?><SVG xmlns="x"/>')
        with pytest.raises(UnsupportedImageTypeError):
            _raise_if_svg(stream)

    def it_does_not_raise_on_png_blob(self):
        # ---8-byte PNG signature is enough to prove the regex is not fooled---
        stream = io.BytesIO(b"\x89PNG\r\n\x1a\n" + b"\x00" * 32)
        _raise_if_svg(stream)
        # ---and the stream position is restored so downstream reads still work---
        assert stream.tell() == 0

    def it_preserves_stream_position_on_non_svg(self):
        stream = io.BytesIO(b"\x89PNG\r\n\x1a\n" + b"\x00" * 4000)
        stream.seek(5)
        _raise_if_svg(stream)
        assert stream.tell() == 5

    def it_does_not_raise_on_unreadable_path(self):
        # ---missing path returns empty head, no raise---
        _raise_if_svg("/nonexistent/path/to/image.png")

    def it_includes_filename_in_error_for_path(self, tmp_path):
        svg_path = tmp_path / "icon.svg"
        svg_path.write_bytes(b'<svg xmlns="x"/>')
        with pytest.raises(UnsupportedImageTypeError) as exc_info:
            _raise_if_svg(str(svg_path))
        assert "icon.svg" in str(exc_info.value)

    def it_uses_stream_name_attribute_when_available(self):
        stream = io.BytesIO(b'<svg xmlns="x"/>')
        stream.name = "/tmp/named.svg"
        with pytest.raises(UnsupportedImageTypeError) as exc_info:
            _raise_if_svg(stream)
        assert "named.svg" in str(exc_info.value)


class Describe_looks_like_svg(object):
    """Unit-test suite for `pptx.package._looks_like_svg`."""

    @pytest.mark.parametrize(
        "head",
        [
            b'<svg xmlns="x"/>',
            b'<?xml version="1.0"?><svg/>',
            b"   <SVG xmlns='x'/>",
            b"<svg\n",
        ],
    )
    def it_detects_svg_root_in_head(self, head):
        assert _looks_like_svg(head, None) is True

    @pytest.mark.parametrize(
        "head",
        [
            b"\x89PNG\r\n\x1a\n",
            b"\xff\xd8\xff\xe0",  # JPEG
            b"GIF89a",
            b"<html><body>Not an SVG</body></html>",
            b"",
        ],
    )
    def it_rejects_non_svg_head(self, head):
        assert _looks_like_svg(head, None) is False

    def it_trusts_svg_extension_for_xml_like_content(self):
        # ---no <svg tag in head but starts with XML prologue and name ends .svg---
        head = b'<?xml version="1.0"?>\n' + b"<!-- comment -->\n" * 100
        assert _looks_like_svg(head, "logo.svg") is True

    def it_ignores_svg_extension_for_binary_content(self):
        head = b"\x89PNG\r\n\x1a\n"
        assert _looks_like_svg(head, "oops.svg") is False


class Describe_MediaParts(object):
    """Unit-test suite for `pptx.package._MediaParts` objects."""

    def it_can_iterate_the_media_parts_in_the_package(self, iter_fixture):
        media_parts, expected_parts = iter_fixture
        assert list(media_parts) == expected_parts

    def it_can_get_or_add_a_media_part(self, get_or_add_fixture):
        media_parts, media_, sha1, MediaPart_, calls = get_or_add_fixture[:5]
        media_part_ = get_or_add_fixture[5]

        media_part = media_parts.get_or_add_media_part(media_)

        media_parts._find_by_sha1.assert_called_once_with(media_parts, sha1)
        assert MediaPart_.new.call_args_list == calls
        assert media_part is media_part_

    def it_can_find_a_media_part_by_sha1(self, find_fixture):
        media_parts, sha1, expected_value = find_fixture
        media_part = media_parts._find_by_sha1(sha1)
        assert media_part is expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[True, False])
    def find_fixture(self, request, _iter_, media_part_):
        media_part_is_present = request.param
        media_parts = _MediaParts(None)
        _iter_.return_value = iter((media_part_,))
        sha1 = "foobar"
        if media_part_is_present:
            media_part_.sha1 = "foobar"
            expected_value = media_part_
        else:
            media_part_.sha1 = "barfoo"
            expected_value = None
        return media_parts, sha1, expected_value

    @pytest.fixture(params=[True, False])
    def get_or_add_fixture(
        self, request, package_, media_, MediaPart_, media_part_, _find_by_sha1_
    ):
        media_present = request.param
        media_parts = _MediaParts(package_)
        media_.sha1 = sha1 = "2468"
        calls = [] if media_present else [call(package_, media_)]
        _find_by_sha1_.return_value = media_part_ if media_present else None
        MediaPart_.new.return_value = None if media_present else media_part_
        return media_parts, media_, sha1, MediaPart_, calls, media_part_

    @pytest.fixture
    def iter_fixture(self, request, package_):
        def rel(is_external, reltype, part):
            return instance_mock(
                request,
                _Relationship,
                is_external=is_external,
                reltype=reltype,
                target_part=part,
            )

        part_mocks = (
            instance_mock(request, Part, name="linked-media"),
            instance_mock(request, Part, name="slide"),
            instance_mock(request, Part, name="embeded-media"),
        )

        rels = (
            rel(True, RT.MEDIA, part_mocks[0]),
            rel(True, RT.VIDEO, part_mocks[0]),
            rel(False, RT.SLIDE, part_mocks[1]),
            rel(False, RT.MEDIA, part_mocks[2]),
            rel(False, RT.VIDEO, part_mocks[2]),
        )

        package_.iter_rels.return_value = iter(rels)

        media_parts = _MediaParts(package_)
        expected_parts = [part_mocks[2]]
        return media_parts, expected_parts

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _find_by_sha1_(self, request):
        return method_mock(request, _MediaParts, "_find_by_sha1", autospec=True)

    @pytest.fixture
    def _iter_(self, request):
        return method_mock(request, _MediaParts, "__iter__")

    @pytest.fixture
    def media_(self, request):
        return instance_mock(request, Video)

    @pytest.fixture
    def MediaPart_(self, request):
        return class_mock(request, "pptx.package.MediaPart")

    @pytest.fixture
    def media_part_(self, request):
        return instance_mock(request, MediaPart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)
