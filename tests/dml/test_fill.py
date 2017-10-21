# encoding: utf-8

"""Test suite for pptx.dml.fill module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.fill import (
    _BlipFill, _Fill, FillFormat, _GradFill, _GrpFill, _NoFill, _NoneFill,
    _PattFill, _SolidFill
)
from pptx.enum.dml import MSO_FILL

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock, method_mock


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

    def it_provides_access_to_its_foreground_color(self, fore_color_fixture):
        fill, color_ = fore_color_fixture
        color = fill.fore_color
        assert color is color_

    def it_knows_its_fill_type(self, fill_type_fixture):
        fill, expected_value = fill_type_fixture
        fill_type = fill.type
        assert fill_type == expected_value

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

    @pytest.fixture
    def fill_type_fixture(self, fill_):
        expected_value = fill_.type = 42
        fill = FillFormat(None, fill_)
        return fill, expected_value

    @pytest.fixture
    def fore_color_fixture(self, fill_, color_):
        fill_.fore_color = color_
        fill_format = FillFormat(None, fill_)
        return fill_format, color_

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
    def color_(self, request):
        return instance_mock(request, ColorFormat)

    @pytest.fixture
    def fill_(self, request):
        return instance_mock(request, _Fill)

    @pytest.fixture
    def no_fill_(self, request):
        return instance_mock(request, _NoFill)

    @pytest.fixture
    def patt_fill_(self, request):
        return instance_mock(request, _PattFill)

    @pytest.fixture
    def solid_fill_(self, request):
        return instance_mock(request, _SolidFill)


class Describe_Fill(object):

    def it_raises_on_fore_color_access(self, raise_fixture):
        fill, exception_type = raise_fixture
        with pytest.raises(exception_type):
            fill.fore_color

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def raise_fixture(self):
        fill = _Fill('foobar')
        exception_type = TypeError
        return fill, exception_type


class Describe_BlipFill(object):

    def it_knows_its_fill_type(self, fill_type_fixture):
        blip_fill, expected_value = fill_type_fixture
        fill_type = blip_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        xFill = element('a:blipFill')
        blip_fill = _BlipFill(xFill)
        expected_value = MSO_FILL.PICTURE
        return blip_fill, expected_value


class Describe_GradFill(object):

    def it_knows_its_fill_type(self, fill_type_fixture):
        grad_fill, expected_value = fill_type_fixture
        fill_type = grad_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        xFill = element('a:gradFill')
        grad_fill = _GradFill(xFill)
        expected_value = MSO_FILL.GRADIENT
        return grad_fill, expected_value


class Describe_GrpFill(object):

    def it_knows_its_fill_type(self, fill_type_fixture):
        grp_fill, expected_value = fill_type_fixture
        fill_type = grp_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        xFill = element('a:grpFill')
        grp_fill = _GrpFill(xFill)
        expected_value = MSO_FILL.GROUP
        return grp_fill, expected_value


class Describe_NoFill(object):

    def it_knows_its_fill_type(self, fill_type_fixture):
        no_fill, expected_value = fill_type_fixture
        fill_type = no_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        xFill = element('a:noFill')
        no_fill = _NoFill(xFill)
        expected_value = MSO_FILL.BACKGROUND
        return no_fill, expected_value


class Describe_NoneFill(object):

    def it_knows_its_fill_type(self, fill_type_fixture):
        none_fill, expected_value = fill_type_fixture
        fill_type = none_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        none_fill = _NoneFill(None)
        expected_value = None
        return none_fill, expected_value


class Describe_PattFill(object):

    def it_knows_its_fill_type(self, fill_type_fixture):
        patt_fill, expected_value = fill_type_fixture
        fill_type = patt_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        patt_fill = _PattFill(element('a:pattFill'))
        expected_value = MSO_FILL.PATTERNED
        return patt_fill, expected_value


class Describe_SolidFill(object):

    def it_knows_its_fill_type(self, fill_type_fixture):
        solid_fill, expected_value = fill_type_fixture
        fill_type = solid_fill.type
        assert fill_type == expected_value

    def it_provides_access_to_its_foreground_color(self, fore_color_fixture):
        solid_fill, solidFill, color_ = fore_color_fixture[:3]
        ColorFormat_from_colorchoice_parent_ = fore_color_fixture[3]

        color = solid_fill.fore_color

        ColorFormat_from_colorchoice_parent_.assert_called_once_with(
            solidFill
        )
        assert color is color_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        solid_fill = _SolidFill(element('a:solidFill'))
        expected_value = MSO_FILL.SOLID
        return solid_fill, expected_value

    @pytest.fixture
    def fore_color_fixture(self, ColorFormat_from_colorchoice_parent_,
                           color_):
        ColorFormat_from_colorchoice_parent_.return_value = color_
        solidFill = element('a:solidFill')

        solid_fill = _SolidFill(solidFill)
        return (
            solid_fill, solidFill, color_,
            ColorFormat_from_colorchoice_parent_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ColorFormat_from_colorchoice_parent_(self, request):
        return method_mock(request, ColorFormat, 'from_colorchoice_parent')

    @pytest.fixture
    def color_(self, request):
        return instance_mock(request, ColorFormat)
