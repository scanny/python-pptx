# encoding: utf-8

"""
Series-related oxml objects.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..xmlchemy import BaseOxmlElement, ZeroOrOne


class CT_Marker(BaseOxmlElement):
    """
    `c:marker` custom element class, containing visual properties for a data
    point marker on line-type charts.
    """
    _tag_seq = (
        'c:symbol', 'c:size', 'c:spPr', 'c:extLst'
    )
    spPr = ZeroOrOne('c:spPr', successors=_tag_seq[3:])
    del _tag_seq
