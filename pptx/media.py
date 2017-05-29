# encoding: utf-8

"""Objects related to images, audio, and video."""


class Video(object):
    """Immutable value object representing a video such as MP4."""

    @property
    def filename(self):
        """Return a filename.ext string appropriate to this video.

        The base filename from the original path is used if this image was
        loaded from the filesystem. If no filename is available, such as when
        the video object is created from an in-memory stream, the string
        'movie.{ext}' is used where 'ext' is suitable to the video format,
        such as 'mp4'.
        """
        raise NotImplementedError
