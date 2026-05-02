"""Unit-test suite for `pptx/__init__.py` content-type registrations."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from pptx import Presentation, content_type_to_part_class_map
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part, PartFactory
from pptx.parts.image import ImagePart
from pptx.parts.media import MediaPart
from pptx.util import Inches


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

    @pytest.mark.parametrize(
        "content_type",
        [
            CT.BMP,
            CT.GIF,
            CT.JPEG,
            CT.MS_PHOTO,
            CT.PNG,
            CT.TIFF,
            CT.X_EMF,
            CT.X_WMF,
            # -- non-standard aliases (issue #929, #1084) --
            "image/jpg",
            "image/tif",
        ],
    )
    def it_maps_image_content_types_to_ImagePart(self, content_type):
        """Image MIME-types (including non-standard aliases) resolve to |ImagePart|.

        The ``image/tif`` alias was added for issue #1084, which reports
        ``AttributeError: 'Part' object has no attribute 'image'`` when
        reading a deck whose ``[Content_Types].xml`` declares TIFF images
        with the non-IANA ``image/tif`` media type.
        """
        assert content_type_to_part_class_map[content_type] is ImagePart
        assert PartFactory.part_type_for[content_type] is ImagePart


class DescribePartFactoryCaseInsensitiveLookup(object):
    """``PartFactory._part_cls_for`` is case-insensitive — issue #1084.

    Real-world ``.pptx`` files occasionally carry content-type strings whose
    casing differs from python-pptx's canonical lowercase registration
    (e.g. ``Image/Tiff`` rather than ``image/tiff``). Per RFC 2046 §4.1 MIME
    type tokens are case-insensitive, so the lookup should normalize before
    falling through to the generic |Part|.
    """

    @pytest.mark.parametrize(
        "content_type",
        [
            "image/tiff",  # -- canonical, control --
            "Image/Tiff",  # -- mixed case --
            "IMAGE/TIFF",  # -- all upper --
            "image/TIFF",  # -- type/SUBTYPE --
        ],
    )
    def it_resolves_tiff_variants_to_ImagePart(self, content_type):
        assert PartFactory._part_cls_for(content_type) is ImagePart

    @pytest.mark.parametrize(
        "content_type",
        [
            "Audio/MPEG",
            "AUDIO/MPEG",
        ],
    )
    def it_resolves_audio_content_type_case_variants_to_MediaPart(self, content_type):
        assert PartFactory._part_cls_for(content_type) is MediaPart

    def it_still_falls_back_to_Part_for_unknown_content_types(self):
        """Unknown content-types still resolve to the generic |Part|."""
        assert PartFactory._part_cls_for("application/made-up") is Part


class DescribeIssue323Regression(object):
    """End-to-end regression guard for issue #323.

    "Cannot add video slide, if another slide layout in the master template
    has an mp3 media" — resolved by #502 (register audio MIME-types as
    |MediaPart|) and #734 (defensive ``isinstance`` check in
    ``_MediaParts._find_by_sha1``).
    """

    def it_allows_add_movie_when_a_layout_owns_an_mp3(self, tmp_path):
        pkg_bytes = self._deck_with_mp3_in_layout()
        prs = Presentation(io.BytesIO(pkg_bytes))

        slide = prs.slides.add_slide(prs.slide_layouts[0])
        # -- this call raised ``AttributeError: 'Part' object has no
        # -- attribute 'sha1'`` before #502/#734 --
        slide.shapes.add_movie(
            str(self._movie_path()),
            Inches(1),
            Inches(1),
            Inches(2),
            Inches(2),
            mime_type="video/mp4",
        )

        out_path = tmp_path / "out.pptx"
        prs.save(str(out_path))
        # -- round-trip back to ensure the saved deck is well-formed --
        Presentation(str(out_path))

    # -- helpers --

    @staticmethod
    def _movie_path() -> Path:
        return Path(__file__).parent / "test_files" / "dummy.mp4"

    @staticmethod
    def _deck_with_mp3_in_layout() -> bytes:
        """Return PPTX bytes where ``slideLayout1`` owns an ``audio/mpeg`` part."""
        src = Path(__file__).parent / "test_files" / "minimal.pptx"
        with zipfile.ZipFile(io.BytesIO(src.read_bytes()), "r") as zin:
            items = {name: zin.read(name) for name in zin.namelist()}

        items["ppt/media/media1.mp3"] = b"ID3\x03\x00\x00\x00\x00\x00\x00mp3-bytes"

        ct_xml = items["[Content_Types].xml"].decode("utf-8")
        ct_xml = ct_xml.replace(
            "</Types>",
            '<Override PartName="/ppt/media/media1.mp3"'
            ' ContentType="audio/mpeg"/></Types>',
        )
        items["[Content_Types].xml"] = ct_xml.encode("utf-8")

        rels_key = "ppt/slideLayouts/_rels/slideLayout1.xml.rels"
        rels_xml = items.get(rels_key, b"").decode("utf-8") or (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/'
            'package/2006/relationships"/>'
        )
        rels_xml = rels_xml.replace(
            "</Relationships>",
            '<Relationship Id="rIdAudio1"'
            ' Type="http://schemas.microsoft.com/office/2007/relationships/media"'
            ' Target="../media/media1.mp3"/></Relationships>',
        )
        items[rels_key] = rels_xml.encode("utf-8")

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in items.items():
                zout.writestr(name, data)
        return out.getvalue()
