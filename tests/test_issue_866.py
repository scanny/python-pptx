"""End-to-end regression tests for issue #866.

The user reported that passing a file-like object without a filename — for
example a reader backed by a browser-generated blob URL, an HTTP-upload
stream, or raw bytes wrapped in a stream adapter — should still let
:meth:`SlideShapes.add_picture` detect the image content type and choose the
right `.ext` / MIME type.

These tests exercise the full ``Presentation -> add_slide -> add_picture``
path with several stream shapes (seekable, non-seekable, no-tell, and a
seek-that-raises stream) across all PIL-supported raster formats.
"""

from __future__ import annotations

import io

import pytest
from PIL import Image as PILImage

from pptx import Presentation


class _NoSeek:
    """File-like that exposes only ``read`` — no ``seek``, no ``tell``."""

    def __init__(self, data: bytes):
        self._buf = io.BytesIO(data)

    def read(self, *a, **k):
        return self._buf.read(*a, **k)


class _NoTell:
    """File-like that has ``seek`` but no ``tell`` — SVG sniff must still work."""

    def __init__(self, data: bytes):
        self._buf = io.BytesIO(data)

    def read(self, *a, **k):
        return self._buf.read(*a, **k)

    def seek(self, *a, **k):
        return self._buf.seek(*a, **k)


class _AngrySeek:
    """File-like that advertises ``seek``/``tell`` but raises on ``seek``."""

    def __init__(self, data: bytes):
        self._buf = io.BytesIO(data)

    def read(self, *a, **k):
        return self._buf.read(*a, **k)

    def tell(self):
        return 0

    def seek(self, *a, **k):
        raise OSError("unseekable")


def _encode(fmt: str) -> bytes:
    buf = io.BytesIO()
    PILImage.new("RGB", (10, 10)).save(buf, fmt)
    return buf.getvalue()


FORMAT_PARAMS = [
    ("PNG", "png", "image/png"),
    ("JPEG", "jpg", "image/jpeg"),
    ("GIF", "gif", "image/gif"),
    ("BMP", "bmp", "image/bmp"),
    ("TIFF", "tiff", "image/tiff"),
]


class DescribeIssue866:
    """Smoke tests covering the reported blob-image scenario."""

    @pytest.mark.parametrize(("fmt", "ext", "ct"), FORMAT_PARAMS)
    def it_detects_format_from_a_BytesIO(self, fmt, ext, ct):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        stream = io.BytesIO(_encode(fmt))

        picture = slide.shapes.add_picture(stream, 0, 0)

        assert picture.image.ext == ext
        assert picture.image.content_type == ct

    @pytest.mark.parametrize(("fmt", "ext", "ct"), FORMAT_PARAMS)
    def it_detects_format_from_a_non_seekable_stream(self, fmt, ext, ct):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        stream = _NoSeek(_encode(fmt))

        picture = slide.shapes.add_picture(stream, 0, 0)

        assert picture.image.ext == ext
        assert picture.image.content_type == ct

    @pytest.mark.parametrize(("fmt", "ext", "ct"), FORMAT_PARAMS)
    def it_detects_format_from_a_no_tell_stream(self, fmt, ext, ct):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        stream = _NoTell(_encode(fmt))

        picture = slide.shapes.add_picture(stream, 0, 0)

        assert picture.image.ext == ext
        assert picture.image.content_type == ct

    def it_detects_format_from_a_seek_that_raises(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        stream = _AngrySeek(_encode("PNG"))

        picture = slide.shapes.add_picture(stream, 0, 0)

        assert picture.image.ext == "png"

    def it_still_rejects_SVG_content_on_a_non_seekable_stream(self):
        from pptx.exc import UnsupportedImageTypeError

        svg_bytes = (
            b'<?xml version="1.0"?>'
            b'<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>'
        )
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        stream = _NoSeek(svg_bytes)

        with pytest.raises(UnsupportedImageTypeError):
            slide.shapes.add_picture(stream, 0, 0)

    def it_infers_a_filename_from_a_non_seekable_streams_name_attr(self):
        """A non-seekable stream with a ``.name`` still contributes to image.desc."""

        class NoSeekNamed(_NoSeek):
            def __init__(self, data: bytes, name: str):
                super().__init__(data)
                self.name = name

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        stream = NoSeekNamed(_encode("PNG"), "screenshot.png")

        picture = slide.shapes.add_picture(stream, 0, 0)

        # -- filename survives the materialization round-trip --
        assert picture.image.filename == "screenshot.png"
