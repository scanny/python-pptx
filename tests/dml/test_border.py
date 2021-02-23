# encoding: utf-8

"""
Test suite for pptx.dml.line module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.dml.border import BorderFormat
from pptx.oxml.shapes.shared import CT_LineProperties
from pptx.shapes.autoshape import Shape

from ..oxml.unitdata.dml import an_ln
from ..unitutil.cxml import element, xml
from ..unitutil.mock import call, class_mock, instance_mock, property_mock


class DescribeBorderFormat(object):
    def it_knows_its_side(self, side_get_fixture):
        line, expected_value = side_get_fixture
        assert line.side == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('left', 'left'),
        ('right', 'right'),
        ('top', 'top'),
        ('bottom', 'bottom'),
        ('bl_tr', 'bl_tr'),
        ('tl_br', 'tl_br')
    ])
    def side_get_fixture(self, request, shape_):
        side, expected_value = request.param
        border = BorderFormat(shape_, side)
        return border, expected_value

    # fixture components ---------------------------------------------


    @pytest.fixture
    def ln_(self, request):
        return instance_mock(request, CT_LineProperties)

    def ln_bldr(self, w):
        ln_bldr = an_ln().with_nsdecls()
        if w is not None:
            ln_bldr.with_w(w)
        return ln_bldr

    @pytest.fixture
    def shape_(self, request, ln_):
        shape_ = instance_mock(request, Shape)
        shape_.get_or_add_ln.return_value = ln_
        return shape_
