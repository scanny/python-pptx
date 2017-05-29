# encoding: utf-8

"""Unit test suite for pptx.media module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.compat import BytesIO
from pptx.media import Video

from .unitutil.file import absjoin, test_file_dir
from .unitutil.mock import instance_mock, method_mock


TEST_VIDEO_PATH = absjoin(test_file_dir, 'dummy.mp4')


class DescribeVideo(object):

    def it_can_construct_from_a_path(self, from_path_fixture):
        movie_path, mime_type, blob, filename, video_ = from_path_fixture
        video = Video.from_path_or_file_like(movie_path, mime_type)
        Video.from_blob.assert_called_once_with(blob, mime_type, filename)
        assert video is video_

    def it_can_construct_from_a_stream(self, from_stream_fixture):
        movie_stream, mime_type, blob, video_ = from_stream_fixture
        video = Video.from_path_or_file_like(movie_stream, mime_type)
        Video.from_blob.assert_called_once_with(blob, mime_type, None)
        assert video is video_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def from_path_fixture(self, video_, from_blob_):
        movie_path, mime_type = TEST_VIDEO_PATH, 'video/mp4'
        with open(movie_path, 'rb') as f:
            blob = f.read()
        filename = 'dummy.mp4'
        from_blob_.return_value = video_
        return movie_path, mime_type, blob, filename, video_

    @pytest.fixture
    def from_stream_fixture(self, video_, from_blob_):
        with open(TEST_VIDEO_PATH, 'rb') as f:
            blob = f.read()
            movie_stream = BytesIO(blob)
        mime_type = 'video/mp4'
        from_blob_.return_value = video_
        return movie_stream, mime_type, blob, video_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def from_blob_(self, request):
        return method_mock(request, Video, 'from_blob')

    @pytest.fixture
    def video_(self, request):
        return instance_mock(request, Video)
