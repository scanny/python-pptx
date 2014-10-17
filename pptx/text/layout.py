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
        line_source = _LineSource(text)
        text_fitter = cls(line_source, extents, font_file)
        return text_fitter._best_fit_font_size(max_size)

    def _best_fit_font_size(self, max_size):
        """
        Return the largest whole-number point size less than or equal to
        *max_size* that this fitter can fit.
        """
        predicate = self._fits_inside_predicate
        sizes = _BinarySearchTree.from_ordered_sequence(
            range(1, int(max_size)+1)
        )
        return sizes.find_max(predicate)

    @property
    def _fits_inside_predicate(self):
        """
        Return a function taking an integer point size argument that returns
        |True| if the text in this fitter can be wrapped to fit entirely
        within its extents when rendered at that point size.
        """
        raise NotImplementedError


class _BinarySearchTree(object):
    """
    A node in a binary search tree. Uniform for root, subtree root, and leaf
    nodes.
    """
    def __init__(self, value):
        self._value = value
        self._lesser = None
        self._greater = None

    @classmethod
    def from_ordered_sequence(cls, iseq):
        """
        Return the root of a balanced binary search tree populated with the
        values in iterable *iseq*.
        """
        raise NotImplementedError


class _LineSource(object):
    """
    Generates all the possible even-word line breaks in a string of text,
    each in the form of a (line, remainder) 2-tuple where *line* contains the
    text before the break and *remainder* the text after as a |_LineSource|
    object. Its boolean value is |True| when it contains text, |False| when
    its text is the empty string or whitespace only.
    """
    def __init__(self, text):
        self._text = text
