"""Exceptions used with python-pptx.

The base exception class is PythonPptxError.
"""

from __future__ import annotations


class PythonPptxError(Exception):
    """Generic error class."""


class PackageNotFoundError(PythonPptxError):
    """
    Raised when a package cannot be found at the specified path.
    """


class InvalidXmlError(PythonPptxError):
    """
    Raised when a value is encountered in the XML that is not valid according
    to the schema.
    """


class UnsupportedImageTypeError(PythonPptxError):
    """Raised when an image file has a format python-pptx cannot embed.

    Most commonly raised when an SVG file is passed to ``shapes.add_picture()`` or
    ``PicturePlaceholder.insert_picture()``. SVG embedding in PowerPoint (Office 2016+)
    requires a companion PNG raster fallback stored alongside the SVG; python-pptx
    does not include an SVG rasterizer, so callers must pre-rasterize SVG content to
    PNG (or another supported raster format) before inserting it. See the
    "Inserting SVG images" section of the user guide for a recommended workaround.
    """
