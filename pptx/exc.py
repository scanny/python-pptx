# encoding: utf-8

"""
Exceptions used with python-pptx.

The base exception class is PythonPptxError.
"""


class PythonPptxError(Exception):
    """Generic error class."""


class NotXMLError(PythonPptxError):
    """
    Raised when an XML operation (such as parsing) is attempted on a binary
    package item.
    """


class PackageNotFoundError(PythonPptxError):
    """
    Raised when a package cannot be found at the specified path.
    """
