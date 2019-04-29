# encoding: utf-8

"""MediaPart and related objects."""

from __future__ import absolute_import, division, print_function, unicode_literals

import hashlib

from ..opc.package import Part
from ..util import lazyproperty


class MediaPart(Part):
    """A media part, containing an audio or video resource.

    A media part generally has a partname matching the regex
    ``ppt/media/media[1-9][0-9]*.*``.
    """

    @classmethod
    def new(cls, package, media):
        """Return new |MediaPart| instance containing *media*.

        *media* must be a |Media| object.
        """
        partname = package.next_media_partname(media.ext)
        return cls(partname, media.content_type, media.blob, package)

    @lazyproperty
    def sha1(self):
        """The SHA1 hash digest for the media binary of this media part.

        Example: `'1be010ea47803b00e140b852765cdf84f491da47'`
        """
        return hashlib.sha1(self._blob).hexdigest()
