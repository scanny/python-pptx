# encoding: utf-8

"""
Test suite for pptx.package module
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.media import Video
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, _Relationship
from pptx.opc.packuri import PackURI
from pptx.package import _ImageParts, _MediaParts, Package
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.image import Image, ImagePart
from pptx.parts.media import MediaPart


from .unitutil.mock import (
    call, class_mock, instance_mock, method_mock, property_mock
)


class DescribePackage(object):

    def it_provides_access_to_its_core_properties_part(self):
        pkg = Package.open('pptx/templates/default.pptx')
        assert isinstance(pkg.core_properties, CorePropertiesPart)

    def it_can_get_or_add_an_image_part(self, image_part_fixture):
        package, image_file, image_parts_, image_part_ = image_part_fixture
        image_part = package.get_or_add_image_part(image_file)
        image_parts_.get_or_add_image_part.assert_called_once_with(
            image_file
        )
        assert image_part is image_part_

    def it_can_get_or_add_a_media_part(self, media_part_fixture):
        package, media, media_part_ = media_part_fixture
        media_part = package.get_or_add_media_part(media)
        package._media_parts.get_or_add_media_part.assert_called_once_with(
            media
        )
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
    def image_part_fixture(self, image_parts_, image_part_,
                           _image_parts_prop_):
        package = Package()
        image_file = 'foobar.png'
        _image_parts_prop_.return_value = image_parts_
        image_parts_.get_or_add_image_part.return_value = image_part_
        return package, image_file, image_parts_, image_part_

    @pytest.fixture
    def media_part_fixture(self, media_, media_part_, _media_parts_prop_,
                           media_parts_):
        package = Package()
        _media_parts_prop_.return_value = media_parts_
        media_parts_.get_or_add_media_part.return_value = media_part_
        return package, media_, media_part_

    @pytest.fixture
    def m_parts_fixture(self, _MediaParts_, media_parts_):
        package = Package()
        _MediaParts_.return_value = media_parts_
        return package, _MediaParts_, media_parts_

    @pytest.fixture(params=[
        ((3, 4, 2), 1),
        ((4, 2, 1), 3),
        ((2, 3, 1), 4),
    ])
    def next_fixture(self, request, iter_parts_):
        idxs, idx = request.param
        package = Package()
        package.iter_parts.return_value = self.i_image_parts(request, idxs)
        ext = 'foo'
        expected_value = '/ppt/media/image%d.%s' % (idx, ext)
        return package, ext, expected_value

    @pytest.fixture(params=[
        ((3, 4, 2), 1),
        ((4, 2, 1), 3),
        ((2, 3, 1), 4),
    ])
    def nmp_fixture(self, request, iter_parts_):
        idxs, idx = request.param
        package = Package()
        package.iter_parts.return_value = self.i_media_parts(request, idxs)
        ext = 'foo'
        expected_value = '/ppt/media/media%d.%s' % (idx, ext)
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
        return property_mock(request, Package, '_image_parts')

    def i_image_parts(self, request, idxs):
        def part(idx):
            partname = PackURI('/ppt/media/image%d.png' % idx)
            return instance_mock(request, Part, partname=partname)
        return iter([part(idx) for idx in idxs])

    def i_media_parts(self, request, idxs):
        def part(idx):
            partname = PackURI('/ppt/media/media%d.mp4' % idx)
            return instance_mock(request, Part, partname=partname)
        return iter([part(idx) for idx in idxs])

    @pytest.fixture
    def iter_parts_(self, request):
        return property_mock(request, Package, 'iter_parts')

    @pytest.fixture
    def media_(self, request):
        return instance_mock(request, Video)

    @pytest.fixture
    def media_part_(self, request):
        return instance_mock(request, MediaPart)

    @pytest.fixture
    def _MediaParts_(self, request):
        return class_mock(request, 'pptx.package._MediaParts')

    @pytest.fixture
    def media_parts_(self, request):
        return instance_mock(request, _MediaParts)

    @pytest.fixture
    def _media_parts_prop_(self, request):
        return property_mock(request, Package, '_media_parts')


class Describe_ImageParts(object):

    def it_can_iterate_over_the_package_image_parts(self, iter_fixture):
        image_parts, expected_parts = iter_fixture
        assert list(image_parts) == expected_parts

    def it_can_get_a_matching_image_part(self, get_fixture):
        image_parts, image_file, Image_, image_, image_part_ = get_fixture

        image_part = image_parts.get_or_add_image_part(image_file)

        Image_.from_file.assert_called_once_with(image_file)
        image_parts._find_by_sha1.assert_called_once_with(image_.sha1)
        assert image_part is image_part_

    def it_can_add_an_image_part(self, add_fixture):
        image_parts, image_file, Image_, image_ = add_fixture[:4]
        ImagePart_, package_, image_part_ = add_fixture[4:]

        image_part = image_parts.get_or_add_image_part(image_file)

        Image_.from_file.assert_called_once_with(image_file)
        image_parts._find_by_sha1.assert_called_once_with(image_.sha1)
        ImagePart_.new.assert_called_once_with(package_, image_)
        assert image_part is image_part_

    def it_can_find_an_image_part_by_sha1_hash(self, find_fixture):
        image_parts, sha1, expected_value = find_fixture
        image_part = image_parts._find_by_sha1(sha1)
        assert image_part is expected_value

    def but_it_skips_unsupported_image_types(self, request, _iter_):
        sha1 = 'f00beed'
        svg_part_ = instance_mock(
            request, Part, name='svg_part_'
        )
        png_part_ = instance_mock(
            request, ImagePart, name='png_part_', sha1=sha1
        )
        # ---order iteration to encounter svg part before target part---
        _iter_.return_value = iter((svg_part_, png_part_))
        image_parts = _ImageParts(None)

        result = image_parts._find_by_sha1(sha1)

        assert result == png_part_

    # fixtures ---------------------------------------------

    @pytest.fixture
    def add_fixture(self, package_, Image_, image_, _find_by_sha1_,
                    ImagePart_, image_part_):
        image_parts = _ImageParts(package_)
        image_file = 'foobar.png'
        Image_.from_file.return_value = image_
        _find_by_sha1_.return_value = None
        ImagePart_.new.return_value = image_part_
        return (
            image_parts, image_file, Image_, image_, ImagePart_, package_,
            image_part_
        )

    @pytest.fixture(params=[True, False])
    def find_fixture(self, request, _iter_, image_part_):
        image_part_is_present = request.param
        image_parts = _ImageParts(None)
        _iter_.return_value = iter((image_part_,))
        sha1 = 'foobar'
        if image_part_is_present:
            image_part_.sha1 = 'foobar'
            expected_value = image_part_
        else:
            image_part_.sha1 = 'barfoo'
            expected_value = None
        return image_parts, sha1, expected_value

    @pytest.fixture
    def get_fixture(self, Image_, image_, image_part_, _find_by_sha1_):
        image_parts = _ImageParts(None)
        image_file = 'foobar.png'
        Image_.from_file.return_value = image_
        _find_by_sha1_.return_value = image_part_
        return image_parts, image_file, Image_, image_, image_part_

    @pytest.fixture
    def iter_fixture(self, request, package_):

        def rel(is_external, reltype):
            part = instance_mock(request, Part)
            return instance_mock(
                request, _Relationship, is_external=is_external,
                reltype=reltype, target_part=part
            )

        rels = (
            rel(True,  RT.IMAGE),
            rel(False, RT.SLIDE),
            rel(False, RT.IMAGE),
        )

        package_.iter_rels.return_value = iter(
            (rels[0], rels[1], rels[2], rels[2])
        )
        image_parts = _ImageParts(package_)
        expected_parts = [rels[2].target_part]
        return image_parts, expected_parts

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _find_by_sha1_(self, request):
        return method_mock(request, _ImageParts, '_find_by_sha1')

    @pytest.fixture
    def Image_(self, request):
        return class_mock(request, 'pptx.package.Image')

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def ImagePart_(self, request):
        return class_mock(request, 'pptx.package.ImagePart')

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def _iter_(self, request):
        return method_mock(request, _ImageParts, '__iter__')

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)


class Describe_MediaParts(object):

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

    @pytest.fixture(params=[
        True,
        False
    ])
    def find_fixture(self, request, _iter_, media_part_):
        media_part_is_present = request.param
        media_parts = _MediaParts(None)
        _iter_.return_value = iter((media_part_,))
        sha1 = 'foobar'
        if media_part_is_present:
            media_part_.sha1 = 'foobar'
            expected_value = media_part_
        else:
            media_part_.sha1 = 'barfoo'
            expected_value = None
        return media_parts, sha1, expected_value

    @pytest.fixture(params=[
        True,
        False
    ])
    def get_or_add_fixture(self, request, package_, media_, MediaPart_,
                           media_part_, _find_by_sha1_):
        media_present = request.param
        media_parts = _MediaParts(package_)
        media_.sha1 = sha1 = '2468'
        calls = [] if media_present else [call(package_, media_)]
        _find_by_sha1_.return_value = media_part_ if media_present else None
        MediaPart_.new.return_value = None if media_present else media_part_
        return media_parts, media_, sha1, MediaPart_, calls, media_part_

    @pytest.fixture
    def iter_fixture(self, request, package_):

        def rel(is_external, reltype, part):
            return instance_mock(
                request, _Relationship, is_external=is_external,
                reltype=reltype, target_part=part
            )

        part_mocks = (
            instance_mock(request, Part, name='linked-media'),
            instance_mock(request, Part, name='slide'),
            instance_mock(request, Part, name='embeded-media'),
        )

        rels = (
            rel(True,  RT.MEDIA, part_mocks[0]),
            rel(True,  RT.VIDEO, part_mocks[0]),
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
        return method_mock(
            request, _MediaParts, '_find_by_sha1', autospec=True
        )

    @pytest.fixture
    def _iter_(self, request):
        return method_mock(request, _MediaParts, '__iter__')

    @pytest.fixture
    def media_(self, request):
        return instance_mock(request, Video)

    @pytest.fixture
    def MediaPart_(self, request):
        return class_mock(request, 'pptx.package.MediaPart')

    @pytest.fixture
    def media_part_(self, request):
        return instance_mock(request, MediaPart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)
