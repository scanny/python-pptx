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


class TextLayoutError(PythonPptxError):
    """Raised when text cannot be laid out to fit within a given shape.

    This occurs for example when a single word in the text is too wide to fit
    the shape's width, even at the smallest font size considered.
    """
