# encoding: utf-8

"""Objects related to images, audio, and video."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import os

from .compat import is_string
from .opc.constants import CONTENT_TYPE as CT
from .util import lazyproperty


class Video(object):
    """Immutable value object representing a video such as MP4."""

    def __init__(self, blob, mime_type, filename):
        super(Video, self).__init__()
        self._blob = blob
        self._mime_type = mime_type
        self._filename = filename

    @classmethod
    def from_blob(cls, blob, mime_type, filename=None):
        """Return a new |Video| object loaded from image binary in *blob*."""
        return cls(blob, mime_type, filename)

    @classmethod
    def from_path_or_file_like(cls, movie_file, mime_type):
        """Return a new |Video| object containing video in *movie_file*.

        *movie_file* can be either a path (string) or a file-like
        (e.g. StringIO) object.
        """
        if is_string(movie_file):
            # treat movie_file as a path
            with open(movie_file, 'rb') as f:
                blob = f.read()
            filename = os.path.basename(movie_file)
        else:
            # assume movie_file is a file-like object
            blob = movie_file.read()
            filename = None

        return cls.from_blob(blob, mime_type, filename)

    @property
    def ext(self):
        """Return the file extension for this video, e.g. 'mp4'.

        The extension is that from the actual filename if known. Otherwise
        it is the lowercase canonical extension for the video's MIME type.
        'vid' is used if the MIME type is 'video/unknown'.
        """
        if self._filename:
            return os.path.splitext(self._filename)[1].lstrip('.')
        return {
            CT.ASF:            'asf',
            CT.AVI:            'avi',
            'video/avi':       'avi',
            'video/msvideo':   'avi',
            'video/x-msvideo': 'avi',
            CT.MOV:            'mov',
            CT.MP4:            'mp4',
            CT.MPG:            'mpg',
            CT.SWF:            'swf',
            CT.WMV:            'wmv',
        }.get(self._mime_type, 'vid')

    @property
    def filename(self):
        """Return a filename.ext string appropriate to this video.

        The base filename from the original path is used if this image was
        loaded from the filesystem. If no filename is available, such as when
        the video object is created from an in-memory stream, the string
        'movie.{ext}' is used where 'ext' is suitable to the video format,
        such as 'mp4'.
        """
        if self._filename is not None:
            return self._filename
        return 'movie.%s' % self.ext

    @lazyproperty
    def sha1(self):
        """The SHA1 hash digest for the media binary of this media part.

        Example: `'1be010ea47803b00e140b852765cdf84f491da47'`
        """
        raise NotImplementedError
