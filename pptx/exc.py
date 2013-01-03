# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Exceptions used with python-pptx.

The base exception class is PythonPptxError.
"""


class PythonPptxError(Exception):
    """Generic error class."""


class CorruptedPackageError(PythonPptxError):
    """
    Raised when a package item in an Open XML package cannot be found or is
    invalid.
    """

class DuplicateKeyError(PythonPptxError):
    """
    Raised by a unique collection when an attempt is made to add an item with
    a key already in the collection.
    """

class InvalidPackageError(PythonPptxError):
    """
    Raised when a package does not contain a valid presentation part (possibly
    because it's a Word or Excel package).
    """

class NotXMLError(PythonPptxError):
    """
    Raised when an XML operation (such as parsing) is attempted on a binary
    package item.
    """

class PackageNotFoundError(PythonPptxError):
    """
    Raised when a package cannot be found at the specified path.
    """


