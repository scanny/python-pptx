# encoding: utf-8

"""
Objects related to layout of rendered text, such as TextFitter.
"""

from __future__ import absolute_import, print_function


class TextFitter(tuple):
    """
    Value object that knows how to fit text into given rectangular extents.
    """
    def __new__(cls, line_source, extents, font_file):
        width, height = extents
        return tuple.__new__(cls, (line_source, width, height, font_file))

    @classmethod
    def best_fit_font_size(cls, text, extents, max_size, font_file):
        """
        Return the largest whole-number point size less than or equal to
        *max_size* that allows *text* to fit completely within *extents* when
        rendered using font defined in *font_file*.
        """
        raise NotImplementedError
