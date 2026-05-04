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


class PackageTooLargeError(PythonPptxError):
    """Raised when a package's declared uncompressed size exceeds the configured limit.

    This guards against maliciously-crafted "zip-bomb" packages whose central-directory
    entries declare combined uncompressed sizes large enough to exhaust memory when the
    package reader loads all parts into memory.

    The default threshold can be adjusted via the ``PPTX_MAX_UNCOMPRESSED_SIZE``
    environment variable or by assigning to
    ``pptx.opc.serialized.MAX_UNCOMPRESSED_PACKAGE_SIZE`` before opening a package.
    See ``docs/dev/security.rst`` for the trust model and recommended handling for
    untrusted inputs.
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


class UnsupportedImageTypeError(PythonPptxError):
    """Raised when an image file has a format python-pptx cannot embed.

    Most commonly raised when an SVG file is passed to ``shapes.add_picture()`` or
    ``PicturePlaceholder.insert_picture()``. SVG embedding in PowerPoint (Office 2016+)
    requires a companion PNG raster fallback stored alongside the SVG; python-pptx
    does not include an SVG rasterizer, so callers must pre-rasterize SVG content to
    PNG (or another supported raster format) before inserting it. See the
    "Inserting SVG images" section of the user guide for a recommended workaround.
    """


class EncryptedPackageError(PythonPptxError):
    """Raised when opening an encrypted .pptx without a valid password.

    Also raised when the optional ``msoffcrypto-tool`` dependency is required to decrypt
    or encrypt a package but is not installed.
    """


class RmsProtectedPackageError(EncryptedPackageError):
    """Raised when opening a .pptx wrapped in Azure RMS / AIP / IRM protection.

    "Rights Management Services" (also marketed as Azure Information Protection /
    Microsoft Purview Information Protection / "Information Rights Management") wraps
    the regular OOXML zip inside a CFBF (OLE2 compound file) container that stores the
    encrypted payload under a ``DRMContent`` stream and a ``DRMEncryptedTransform``
    descriptor. Unlike an ECMA-376 Agile-Encryption package, an RMS package cannot be
    decrypted with a password alone — the user's Azure AD / Microsoft 365 identity must
    be presented to the RMS service to retrieve the content key.

    python-pptx does not bundle an RMS client (the Microsoft Information Protection SDK
    is C#/.NET-only and requires an interactive Azure AD login flow). See
    :ref:`rms-protected` in the user guide for recommended workarounds — delegating
    decryption to Microsoft Office automation, the MIP SDK, or a pre-processing step.
    """
