# encoding: utf-8

"""
Objects related to system font file lookup.
"""

from __future__ import absolute_import, print_function


class FontFiles(object):
    """
    A class-based singleton serving as a lazy cache for system font details.
    """
    @classmethod
    def find(cls, family_name, is_bold, is_italic):
        """
        Return the absolute path to the installed OpenType font having
        *family_name* and the styles *is_bold* and *is_italic*.
        """
        raise NotImplementedError
