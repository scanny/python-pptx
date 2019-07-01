# encoding: utf-8

"""Unit test suite for pptx.parts.image module."""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.media import Video
from pptx.package import Package
from pptx.parts.media import MediaPart

from ..unitutil.mock import initializer_mock, instance_mock


class DescribeMediaPart(object):
    def it_can_construct_from_a_media_object(self, new_fixture):
        package_, media_, _init_, partname_ = new_fixture

        media_part = MediaPart.new(package_, media_)

        package_.next_media_partname.assert_called_once_with(media_.ext)
        _init_.assert_called_once_with(
            media_part, partname_, media_.content_type, media_.blob, package_
        )
        assert isinstance(media_part, MediaPart)

    def it_knows_the_sha1_hash_of_the_media(self, sha1_fixture):
        media_part, expected_value = sha1_fixture
        sha1 = media_part.sha1
        assert sha1 == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self, request, package_, media_, _init_):
        partname_ = package_.next_media_partname.return_value = "media42.mp4"
        media_.blob, media_.content_type = b"blob-bytes", "video/mp4"
        return package_, media_, _init_, partname_

    @pytest.fixture
    def sha1_fixture(self):
        blob = b"blobish-bytes"
        media_part = MediaPart(None, None, blob, None)
        expected_value = "61efc464c21e54cfc1382fb5b6ef7512e141ceae"
        return media_part, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def media_(self, request):
        return instance_mock(request, Video)

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, MediaPart, autospec=True)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)
