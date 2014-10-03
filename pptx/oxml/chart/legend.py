# encoding: utf-8

"""
lxml custom element classes for legend-related XML elements.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ...enum.chart import XL_LEGEND_POSITION
from ..xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class CT_Legend(BaseOxmlElement):
    """
    ``<c:legend>`` custom element class
    """
    _tag_seq = (
        'c:legendPos', 'c:legendEntry', 'c:layout', 'c:overlay', 'c:spPr',
        'c:txPr', 'c:extLst'
    )
    legendPos = ZeroOrOne('c:legendPos', successors=_tag_seq[1:])
    layout = ZeroOrOne('c:layout', successors=_tag_seq[3:])
    overlay = ZeroOrOne('c:overlay', successors=_tag_seq[4:])
    del _tag_seq

    @property
    def horz_offset(self):
        """
        The float value in ./c:layout/c:manualLayout/c:x when
        ./c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that
        XPath expression has no match.
        """
        layout = self.layout
        if layout is None:
            return 0.0
        return layout.horz_offset

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:layout/c:manualLayout/c:x@val to *offset* and
        ./c:layout/c:manualLayout/c:xMode@val to "factor". Remove
        ./c:layout/c:manualLayout if *offset* == 0.
        """
        layout = self.get_or_add_layout()
        layout.horz_offset = offset


class CT_LegendPos(BaseOxmlElement):
    """
    ``<c:legendPos>`` element specifying position of legend with respect to
    chart as a member of ST_LegendPos.
    """
    val = OptionalAttribute(
        'val', XL_LEGEND_POSITION, default=XL_LEGEND_POSITION.RIGHT
    )
