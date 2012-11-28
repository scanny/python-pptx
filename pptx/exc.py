# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Exceptions used with python-pptx.

The base exception class is PythonPptxError.
"""


class PythonPptxError(Exception):
    """Generic error class."""


class CorruptedTemplateError(PythonPptxError):
    """Raised when a package-part in a PowerPoint Open XML presentation template cannot be found or is invalid."""


