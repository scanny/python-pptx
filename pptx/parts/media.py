# encoding: utf-8

"""MediaPart and related objects."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..opc.package import Part


class MediaPart(Part):
    """A media part, containing an audio or video resource.

    A media part generally has a partname matching the regex
    ``ppt/media/media[1-9][0-9]*.*``.
    """
