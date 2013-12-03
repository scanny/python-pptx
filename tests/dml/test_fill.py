# encoding: utf-8

"""
Test suite for pptx.text module.
"""

from __future__ import absolute_import

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat
from pptx.enum import MSO_FILL_TYPE as MSO_FILL

from ..oxml.unitdata.dml import (
    a_blipFill, a_gradFill, a_grpFill, a_noFill, a_pattFill, a_solidFill,
    an_spPr
)
from ..unitutil import actual_xml


class DescribeFillFormat(object):

    def it_can_set_the_fill_type_to_solid(self, set_solid_fixture_):
        fill, spPr_with_solidFill_xml = set_solid_fixture_
        fill.solid()
        assert actual_xml(fill._xPr) == spPr_with_solidFill_xml

    def it_knows_the_type_of_fill_it_is(self, fill_type_fixture_):
        fill_format, fill_type = fill_type_fixture_
        print(actual_xml(fill_format._xPr))
        assert fill_format.fill_type == fill_type

    def it_provides_access_to_the_foreground_color_object(
            self, fore_color_fixture_):
        fill_format, fore_color_type = fore_color_fixture_
        print(actual_xml(fill_format._xPr))
        assert isinstance(fill_format.fore_color, fore_color_type)

    def it_raises_on_fore_color_get_for_fill_types_that_dont_have_one(
            self, fore_color_raise_fixture_):
        fill_format, exception_type = fore_color_raise_fixture_
        with pytest.raises(exception_type):
            fill_format.fore_color

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        'none', 'no', 'solid', 'grad', 'blip', 'patt', 'grp'
    ])
    def fill_type_fixture_(self, request):
        mapping = {
            'none':  ('_spPr_bldr',                None),
            'grad':  ('_spPr_with_gradFill_bldr',  MSO_FILL.GRADIENT),
            'solid': ('_spPr_with_solidFill_bldr', MSO_FILL.SOLID),
            'no':    ('_spPr_with_noFill_bldr',    MSO_FILL.BACKGROUND),
            'blip':  ('_spPr_with_blipFill_bldr',  MSO_FILL.PICTURE),
            'patt':  ('_spPr_with_pattFill_bldr',  MSO_FILL.PATTERNED),
            'grp':   ('_spPr_with_grpFill_bldr',   MSO_FILL.GROUP),
        }
        spPr_bldr_name, fill_type = mapping[request.param]
        spPr_bldr = request.getfuncargvalue(spPr_bldr_name)
        spPr = spPr_bldr.element
        fill_format = FillFormat.from_fill_parent(spPr)
        return fill_format, fill_type

    @pytest.fixture(params=['solid'])
    def fore_color_fixture_(self, request):
        mapping = {
            'solid': ('_spPr_with_solidFill_bldr', ColorFormat),
        }
        spPr_bldr_name, fore_color_type = mapping[request.param]
        spPr_bldr = request.getfuncargvalue(spPr_bldr_name)
        spPr = spPr_bldr.element
        fill_format = FillFormat.from_fill_parent(spPr)
        return fill_format, fore_color_type

    @pytest.fixture(params=['none', 'blip', 'grad', 'grp', 'patt'])
    def fore_color_raise_fixture_(self, request):
        mapping = {
            'none':  ('_spPr_bldr',                TypeError),
            'blip':  ('_spPr_with_blipFill_bldr',  TypeError),
            'grad':  ('_spPr_with_gradFill_bldr',  NotImplementedError),
            'grp':   ('_spPr_with_grpFill_bldr',   TypeError),
            'patt':  ('_spPr_with_pattFill_bldr',  NotImplementedError),
        }
        spPr_bldr_name, exception_type = mapping[request.param]
        spPr_bldr = request.getfuncargvalue(spPr_bldr_name)
        spPr = spPr_bldr.element
        fill_format = FillFormat.from_fill_parent(spPr)
        return fill_format, exception_type

    def _solid_fill_cases():
        # no fill type yet
        spPr = an_spPr().with_nsdecls().element
        # non-solid fill type present
        spPr_with_gradFill = (
            an_spPr().with_nsdecls()
                     .with_child(a_gradFill())
                     .element
        )
        # solidFill already present
        spPr_with_solidFill = (
            an_spPr().with_nsdecls()
                     .with_child(a_solidFill())
                     .element
        )
        return [spPr, spPr_with_gradFill, spPr_with_solidFill]

    @pytest.fixture(params=_solid_fill_cases())
    def set_solid_fixture_(self, request, spPr_with_solidFill_xml):
        spPr = request.param
        fill = FillFormat.from_fill_parent(spPr)
        return fill, spPr_with_solidFill_xml

    @pytest.fixture
    def spPr_with_solidFill_xml(self):
        return (
            an_spPr().with_nsdecls()
                     .with_child(a_solidFill())
                     .xml()
        )

    @pytest.fixture
    def _spPr_bldr(self, request):
        return an_spPr().with_nsdecls()

    @pytest.fixture
    def _spPr_with_noFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_noFill())

    @pytest.fixture
    def _spPr_with_solidFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_solidFill())

    @pytest.fixture
    def _spPr_with_gradFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_gradFill())

    @pytest.fixture
    def _spPr_with_blipFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_blipFill())

    @pytest.fixture
    def _spPr_with_pattFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_pattFill())

    @pytest.fixture
    def _spPr_with_grpFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_grpFill())
