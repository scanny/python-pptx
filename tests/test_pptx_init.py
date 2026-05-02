"""Unit-test suite for `pptx/__init__.py` content-type registrations."""

from __future__ import annotations

import pytest

from pptx import content_type_to_part_class_map
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory
from pptx.parts.media import MediaPart


class DescribeContentTypeRegistrations(object):
    """Ensure media-related content types map to their expected Part subclasses."""

    @pytest.mark.parametrize(
        "content_type",
        [
            CT.AUDIO,
            CT.AUDIO_AIFF,
            CT.AUDIO_MIDI,
            CT.AUDIO_MP3,
            CT.AUDIO_MP4,
            CT.AUDIO_MPEG,
            CT.AUDIO_OGG,
            CT.AUDIO_WAV,
            CT.AUDIO_X_MS_WMA,
            CT.AUDIO_X_WAV,
        ],
    )
    def it_maps_audio_content_types_to_MediaPart(self, content_type):
        """Audio MIME-types are registered so pre-existing audio parts load as |MediaPart|.

        This is the fix for issue #502 / #926 — opening a deck that already contains audio
        no longer produces a generic `Part` (lacking `.sha1`) and instead yields a
        `MediaPart` with the `.sha1` attribute the de-duplication logic relies on.
        """
        assert content_type_to_part_class_map[content_type] is MediaPart
        assert PartFactory.part_type_for[content_type] is MediaPart

    def it_still_maps_video_content_types_to_MediaPart(self):
        """Regression guard — existing video registrations are preserved."""
        for ct in (
            CT.ASF,
            CT.AVI,
            CT.MOV,
            CT.MP4,
            CT.MPG,
            CT.MS_VIDEO,
            CT.SWF,
            CT.VIDEO,
            CT.WMV,
            CT.X_MS_VIDEO,
        ):
            assert content_type_to_part_class_map[ct] is MediaPart
