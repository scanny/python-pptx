# encoding: utf-8

"""Unit test suite for `pptx.parts.media` module."""

from pptx.media import Video
from pptx.package import Package
from pptx.parts.media import MediaPart

from ..unitutil.mock import initializer_mock, instance_mock


class DescribeMediaPart(object):
    """Unit-test suite for `pptx.parts.media.MediaPart` objects."""

    def it_can_construct_from_a_media_object(self, request):
        media_ = instance_mock(request, Video)
        _init_ = initializer_mock(request, MediaPart)
        package_ = instance_mock(request, Package)
        package_.next_media_partname.return_value = "media42.mp4"
        media_.blob, media_.content_type = b"blob-bytes", "video/mp4"

        media_part = MediaPart.new(package_, media_)

        package_.next_media_partname.assert_called_once_with(media_.ext)
        _init_.assert_called_once_with(
            media_part, "media42.mp4", media_.content_type, package_, media_.blob
        )
        assert isinstance(media_part, MediaPart)

    def it_knows_the_sha1_hash_of_the_media(self):
        assert MediaPart(None, None, None, b"blobish-bytes").sha1 == (
            "61efc464c21e54cfc1382fb5b6ef7512e141ceae"
        )
