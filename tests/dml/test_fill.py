# encoding: utf-8

"""Test suite for pptx.dml.fill module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat, _NoFill, _PattFill, _SolidFill
from pptx.enum.dml import MSO_FILL

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribeFillFormat(object):

    def it_can_set_the_fill_type_to_no_fill(self, background_fixture):
        fill, _NoFill_, expected_xml, no_fill_ = background_fixture

        fill.background()

        assert fill._xPr.xml == expected_xml
        _NoFill_.assert_called_once_with(fill._xPr.eg_fillProperties)
        assert fill._fill is no_fill_

    def it_can_set_the_fill_type_to_solid(self, solid_fixture):
        fill, _SolidFill_, expected_xml, solid_fill_ = solid_fixture

        fill.solid()

        assert fill._xPr.xml == expected_xml
        _SolidFill_.assert_called_once_with(fill._xPr.eg_fillProperties)
        assert fill._fill is solid_fill_

    def it_can_set_the_fill_type_to_patterned(self, patterned_fixture):
        fill, _PattFill_, expected_xml, patt_fill_ = patterned_fixture

        fill.patterned()

        assert fill._xPr.xml == expected_xml
        _PattFill_.assert_called_once_with(fill._xPr.eg_fillProperties)
        assert fill._fill is patt_fill_

    def it_knows_its_fill_type(self, fill_type_fixture):
        fill, expected_value = fill_type_fixture
        fill_type = fill.type
        assert fill_type == expected_value

    def it_provides_access_to_its_foreground_color(self, fore_color_fixture):
        fill, ColorFormat_, xFill, color_ = fore_color_fixture

        color = fill.fore_color

        ColorFormat_.from_colorchoice_parent.assert_called_once_with(xFill)
        assert color is color_

    def it_raises_on_fore_color_get_for_fill_types_that_dont_have_one(
            self, fore_color_raise_fixture):
        fill, exception_type = fore_color_raise_fixture
        with pytest.raises(exception_type):
            fill.fore_color

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('p:spPr{a:b=c}',           'p:spPr{a:b=c}/a:noFill'),
        ('a:tcPr/a:blipFill',       'a:tcPr/a:noFill'),
        ('a:rPr/a:gradFill',        'a:rPr/a:noFill'),
        ('a:tcPr/a:grpFill',        'a:tcPr/a:noFill'),
        ('a:defRPr/a:noFill',       'a:defRPr/a:noFill'),
        ('a:endParaRPr/a:pattFill', 'a:endParaRPr/a:noFill'),
        ('a:rPr/a:solidFill',       'a:rPr/a:noFill'),
    ])
    def background_fixture(self, request, no_fill_):
        cxml, expected_cxml = request.param

        fill = FillFormat.from_fill_parent(element(cxml))
        # --mock must be after FillFormat call to avoid poss. contructor call
        _NoFill_ = class_mock(
            request, 'pptx.dml.fill._NoFill', return_value=no_fill_,
            autospec=True
        )
        expected_xml = xml(expected_cxml)
        return fill, _NoFill_, expected_xml, no_fill_

    @pytest.fixture(params=[
        ('p:spPr',                None),
        ('a:rPr/a:blipFill',      MSO_FILL.PICTURE),
        ('a:tcPr/a:gradFill',     MSO_FILL.GRADIENT),
        ('a:defRPr/a:grpFill',    MSO_FILL.GROUP),
        ('a:endParaRPr/a:noFill', MSO_FILL.BACKGROUND),
        ('p:spPr/a:pattFill',     MSO_FILL.PATTERNED),
        ('a:rPr/a:solidFill',     MSO_FILL.SOLID),
    ])
    def fill_type_fixture(self, request):
        xPr_cxml, expected_value = request.param
        xPr = element(xPr_cxml)

        fill = FillFormat.from_fill_parent(xPr)
        return fill, expected_value

    @pytest.fixture(params=[
        ('p:spPr/a:solidFill', _SolidFill),
    ])
    def fore_color_fixture(self, request, ColorFormat_, color_):
        xPr_cxml, FillCls = request.param
        xPr = element(xPr_cxml)
        xFill = xPr[0]
        x_fill = FillCls(xFill)
        ColorFormat_.from_colorchoice_parent.return_value = color_

        fill = FillFormat(xPr, x_fill)
        return fill, ColorFormat_, xFill, color_

    @pytest.fixture(params=[
        ('p:spPr',                TypeError),
        ('a:rPr/a:blipFill',      TypeError),
        ('a:tcPr/a:gradFill',     NotImplementedError),
        ('a:defRPr/a:grpFill',    TypeError),
        ('a:endParaRPr/a:noFill', TypeError),
        ('p:spPr/a:pattFill',     NotImplementedError),
    ])
    def fore_color_raise_fixture(self, request):
        xPr_cxml, exception_type = request.param
        xPr = element(xPr_cxml)

        fill = FillFormat.from_fill_parent(xPr)
        return fill, exception_type

    @pytest.fixture(params=[
        ('p:spPr{a:b=c}',           'p:spPr{a:b=c}/a:pattFill'),
        ('a:tcPr/a:blipFill',       'a:tcPr/a:pattFill'),
        ('a:rPr/a:gradFill',        'a:rPr/a:pattFill'),
        ('a:tcPr/a:grpFill',        'a:tcPr/a:pattFill'),
        ('a:defRPr/a:noFill',       'a:defRPr/a:pattFill'),
        ('a:endParaRPr/a:pattFill', 'a:endParaRPr/a:pattFill'),
        ('a:rPr/a:solidFill',       'a:rPr/a:pattFill'),
    ])
    def patterned_fixture(self, request, patt_fill_):
        cxml, expected_cxml = request.param

        fill = FillFormat.from_fill_parent(element(cxml))
        # --mock must be after FillFormat call to avoid poss. contructor call
        _PattFill_ = class_mock(
            request, 'pptx.dml.fill._PattFill', return_value=patt_fill_,
            autospec=True
        )
        expected_xml = xml(expected_cxml)
        return fill, _PattFill_, expected_xml, patt_fill_

    @pytest.fixture(params=[
        ('p:spPr{a:b=c}',            'p:spPr{a:b=c}/a:solidFill'),
        ('a:tcPr/a:blipFill',        'a:tcPr/a:solidFill'),
        ('a:rPr/a:gradFill',         'a:rPr/a:solidFill'),
        ('a:tcPr/a:grpFill',         'a:tcPr/a:solidFill'),
        ('a:defRPr/a:noFill',        'a:defRPr/a:solidFill'),
        ('a:rPr/a:pattFill',         'a:rPr/a:solidFill'),
        ('a:endParaRPr/a:solidFill', 'a:endParaRPr/a:solidFill'),
    ])
    def solid_fixture(self, request, solid_fill_):
        cxml, expected_cxml = request.param

        fill = FillFormat.from_fill_parent(element(cxml))
        # --mock must be after FillFormat call to avoid poss. contructor call
        _SolidFill_ = class_mock(
            request, 'pptx.dml.fill._SolidFill', return_value=solid_fill_,
            autospec=True
        )
        expected_xml = xml(expected_cxml)
        return fill, _SolidFill_, expected_xml, solid_fill_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ColorFormat_(self, request):
        return class_mock(request, 'pptx.dml.fill.ColorFormat')

    @pytest.fixture
    def color_(self, request):
        return instance_mock(request, ColorFormat)

    @pytest.fixture
    def no_fill_(self, request):
        return instance_mock(request, _NoFill)

    @pytest.fixture
    def patt_fill_(self, request):
        return instance_mock(request, _PattFill)

    @pytest.fixture
    def solid_fill_(self, request):
        return instance_mock(request, _SolidFill)
