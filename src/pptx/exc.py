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


class EncryptedPackageError(PythonPptxError):
    """Raised when opening an encrypted .pptx without a valid password.

    Also raised when the optional ``msoffcrypto-tool`` dependency is required to decrypt
    or encrypt a package but is not installed.
    """
