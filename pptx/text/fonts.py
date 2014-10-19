# encoding: utf-8

"""
Objects related to system font file lookup.
"""

from __future__ import absolute_import, print_function

import sys


class FontFiles(object):
    """
    A class-based singleton serving as a lazy cache for system font details.
    """

    _font_files = None

    @classmethod
    def find(cls, family_name, is_bold, is_italic):
        """
        Return the absolute path to the installed OpenType font having
        *family_name* and the styles *is_bold* and *is_italic*.
        """
        if cls._font_files is None:
            cls._font_files = cls._installed_fonts()
        return cls._font_files[(family_name, is_bold, is_italic)]

    @classmethod
    def _installed_fonts(cls):
        """
        Return a dict mapping a font descriptor to its font file path,
        containing all the font files resident on the current machine. The
        font descriptor is a (family_name, is_bold, is_italic) 3-tuple.
        """
        fonts = {}
        for d in cls._font_directories():
            for key, path in cls._iter_font_files_in(d):
                fonts[key] = path
        return fonts

    @classmethod
    def _font_directories(cls):
        """
        Return a sequence of directory paths likely to contain fonts on the
        current platform.
        """
        if sys.platform.startswith('darwin'):
            return cls._os_x_font_directories()
        if sys.platform.startswith('win32'):
            return cls._windows_font_directories()
        raise OSError('unsupported operating system')

    @classmethod
    def _iter_font_files_in(cls, directory):
        """
        Generate the OpenType font files found in and under *directory*. Each
        item is a key/value pair. The key is a (family_name, is_bold,
        is_italic) 3-tuple, like ('Arial', True, False), and the value is the
        absolute path to the font file.
        """
        raise NotImplementedError

    @classmethod
    def _os_x_font_directories(cls):
        """
        Return a sequence of directory paths on a Mac in which fonts are
        likely to be located.
        """
        raise NotImplementedError

    @classmethod
    def _windows_font_directories(cls):
        """
        Return a sequence of directory paths on Windows in which fonts are
        likely to be located.
        """
        raise NotImplementedError
