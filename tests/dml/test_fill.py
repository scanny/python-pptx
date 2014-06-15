# encoding: utf-8

"""
Test suite for pptx.dml.fill module
"""

from __future__ import absolute_import

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL

from ..oxml.unitdata.dml import (
    a_blipFill, a_gradFill, a_grpFill, a_noFill, a_pattFill, a_solidFill,
    an_spPr
)
from ..oxml.unitdata.table import a_tcPr
from ..oxml.unitdata.text import an_rPr


class DescribeFillFormat(object):

    def it_can_set_the_fill_type_to_no_fill(self, set_noFill_fixture_):
        fill_format, xPr_with_noFill_xml = set_noFill_fixture_
        fill_format.background()
        assert fill_format._xPr.xml == xPr_with_noFill_xml

    def it_can_set_the_fill_type_to_solid(self, set_solid_fixture_):
        fill_format, xPr_with_solidFill_xml = set_solid_fixture_
        fill_format.solid()
        assert fill_format._xPr.xml == xPr_with_solidFill_xml

    def it_knows_the_type_of_fill_it_is(self, fill_type_fixture_):
        fill_format, fill_type = fill_type_fixture_
        assert fill_format.type == fill_type

    def it_provides_access_to_the_foreground_color_object(
            self, fore_color_fixture_):
        fill_format, fore_color_type = fore_color_fixture_
        assert isinstance(fill_format.fore_color, fore_color_type)

    def it_raises_on_fore_color_get_for_fill_types_that_dont_have_one(
            self, fore_color_raise_fixture_):
        fill_format, exception_type = fore_color_raise_fixture_
        with pytest.raises(exception_type):
            fill_format.fore_color

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        'none', 'blip', 'grad', 'grp', 'no', 'patt', 'solid'
    ])
    def fill_type_fixture_(self, request, xPr_bldr):
        mapping = {
            'none':  None,
            'blip':  MSO_FILL.PICTURE,
            'grad':  MSO_FILL.GRADIENT,
            'grp':   MSO_FILL.GROUP,
            'no':    MSO_FILL.BACKGROUND,
            'patt':  MSO_FILL.PATTERNED,
            'solid': MSO_FILL.SOLID,
        }
        fill_type_str = request.param
        fill_type = mapping[fill_type_str]
        xFill_bldr = self._xFill_bldr(fill_type_str)
        if xFill_bldr is not None:
            xPr_bldr.with_child(xFill_bldr)
        xPr = xPr_bldr.element
        fill_format = FillFormat.from_fill_parent(xPr)
        return fill_format, fill_type

    @pytest.fixture(params=['solid'])
    def fore_color_fixture_(self, request, xPr_bldr):
        mapping = {
            'solid': ColorFormat,
        }
        fill_type_str = request.param
        fore_color_type = mapping[fill_type_str]
        xFill_bldr = self._xFill_bldr(fill_type_str)
        if xFill_bldr is not None:
            xPr_bldr.with_child(xFill_bldr)
        xPr = xPr_bldr.element
        fill_format = FillFormat.from_fill_parent(xPr)
        return fill_format, fore_color_type

    @pytest.fixture(params=['none', 'blip', 'grad', 'grp', 'no', 'patt'])
    def fore_color_raise_fixture_(self, request, xPr_bldr):
        mapping = {
            'none':  TypeError,
            'blip':  TypeError,
            'grad':  NotImplementedError,
            'grp':   TypeError,
            'no':    TypeError,
            'patt':  NotImplementedError,
        }
        fill_type_str = request.param
        exception_type = mapping[fill_type_str]
        xFill_bldr = self._xFill_bldr(fill_type_str)
        if xFill_bldr is not None:
            xPr_bldr.with_child(xFill_bldr)
        xPr = xPr_bldr.element
        fill_format = FillFormat.from_fill_parent(xPr)
        return fill_format, exception_type

    @pytest.fixture(
        params=['none', 'blip', 'grad', 'grp', 'no', 'patt', 'solid']
    )
    def set_noFill_fixture_(self, request, xPr_bldr):
        fill_type_str = request.param
        xFill_bldr = self._xFill_bldr(fill_type_str)
        if xFill_bldr is not None:
            xPr_bldr.with_child(xFill_bldr)
        xPr = xPr_bldr.element
        fill_format = FillFormat.from_fill_parent(xPr)
        xPr_with_noFill_xml = (
            xPr_bldr.clear()
                    .with_nsdecls()
                    .with_child(a_noFill())
                    .xml()
        )
        return fill_format, xPr_with_noFill_xml

    @pytest.fixture(
        params=['none', 'blip', 'grad', 'grp', 'no', 'patt', 'solid']
    )
    def set_solid_fixture_(self, request, xPr_bldr):
        fill_type_str = request.param
        xFill_bldr = self._xFill_bldr(fill_type_str)
        if xFill_bldr is not None:
            xPr_bldr.with_child(xFill_bldr)
        xPr = xPr_bldr.element
        fill_format = FillFormat.from_fill_parent(xPr)
        xPr_with_solidFill_xml = (
            xPr_bldr.clear()
                    .with_nsdecls()
                    .with_child(a_solidFill())
                    .xml()
        )
        return fill_format, xPr_with_solidFill_xml

    # fixture components ---------------------------------------------

    @pytest.fixture(params=['rPr', 'spPr', 'tcPr'])
    def xPr_bldr(self, request):
        mapping = {
            'rPr':  an_rPr,
            'spPr': an_spPr,
            'tcPr': a_tcPr,
        }
        xPr_bldr_fn = mapping[request.param]
        return xPr_bldr_fn().with_nsdecls()

    def _xFill_bldr(self, fill_type_str):
        mapping = {
            'none':  None,
            'blip':  a_blipFill,
            'grad':  a_gradFill,
            'grp':   a_grpFill,
            'no':    a_noFill,
            'solid': a_solidFill,
            'patt':  a_pattFill,
        }
        xFill_bldr_fn = mapping[fill_type_str]
        if xFill_bldr_fn is None:
            return None
        return xFill_bldr_fn()
