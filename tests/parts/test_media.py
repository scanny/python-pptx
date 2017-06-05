# encoding: utf-8

"""Unit test suite for pptx.parts.image module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.parts.media import MediaPart


class DescribeMediaPart(object):

    def it_knows_the_sha1_hash_of_the_media(self, sha1_fixture):
        media_part, expected_value = sha1_fixture
        sha1 = media_part.sha1
        assert sha1 == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def sha1_fixture(self):
        blob = b'blobish-bytes'
        media_part = MediaPart(None, None, blob, None)
        expected_value = '61efc464c21e54cfc1382fb5b6ef7512e141ceae'
        return media_part, expected_value
