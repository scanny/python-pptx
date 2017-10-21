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
from pptx.enum.dml import MSO_FILL, MSO_PATTERN

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

    def it_provides_access_to_its_background_color(self, back_color_fixture):
        fill, color_ = back_color_fixture
        color = fill.back_color
        assert color is color_

    def it_knows_its_fill_type(self, fill_type_fixture):
        fill, expected_value = fill_type_fixture
        fill_type = fill.type
        assert fill_type == expected_value

    def it_knows_its_pattern(self, pattern_get_fixture):
        fill, expected_value = pattern_get_fixture
        pattern = fill.pattern
        assert pattern == expected_value

    def it_can_change_its_pattern(self, pattern_set_fixture):
        fill, pattern = pattern_set_fixture
        fill.pattern = pattern
        assert fill.pattern is pattern

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def back_color_fixture(self, fill_, color_):
        fill_.back_color = color_
        fill_format = FillFormat(None, fill_)
        return fill_format, color_

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

    @pytest.fixture
    def pattern_get_fixture(self, fill_):
        expected_value = fill_.pattern = MSO_PATTERN.WAVE
        fill = FillFormat(None, fill_)
        return fill, expected_value

    @pytest.fixture
    def pattern_set_fixture(self, fill_):
        pattern = MSO_PATTERN.DIVOT
        fill = FillFormat(None, fill_)
        return fill, pattern

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

    def it_raises_on_fore_color_access(self, fore_raise_fixture):
        fill, exception_type = fore_raise_fixture
        with pytest.raises(exception_type):
            fill.fore_color

    def it_raises_on_back_color_access(self, back_raise_fixture):
        fill, exception_type = back_raise_fixture
        with pytest.raises(exception_type):
            fill.back_color

    def it_raises_on_pattern_access(self, pattern_raise_fixture):
        fill, exception_type = pattern_raise_fixture
        with pytest.raises(exception_type):
            fill.pattern

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def back_raise_fixture(self):
        fill = _Fill('foobar')
        exception_type = TypeError
        return fill, exception_type

    @pytest.fixture
    def fore_raise_fixture(self):
        fill = _Fill('foobar')
        exception_type = TypeError
        return fill, exception_type

    @pytest.fixture
    def pattern_raise_fixture(self):
        fill = _Fill('barfoo')
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

    def it_knows_its_pattern(self, pattern_get_fixture):
        patt_fill, expected_value = pattern_get_fixture
        pattern = patt_fill.pattern
        assert pattern == expected_value

    def it_can_change_its_pattern(self, pattern_set_fixture):
        patt_fill, pattern, pattFill, expected_xml = pattern_set_fixture
        patt_fill.pattern = pattern
        assert pattFill.xml == expected_xml

    def it_provides_access_to_its_foreground_color(self, fore_color_fixture):
        patt_fill, pattFill, expected_xml, color_ = fore_color_fixture[:4]
        ColorFormat_from_colorchoice_parent_ = fore_color_fixture[4]

        color = patt_fill.fore_color

        assert pattFill.xml == expected_xml
        ColorFormat_from_colorchoice_parent_.assert_called_once_with(
            pattFill.fgClr
        )
        assert color is color_

    def it_provides_access_to_its_background_color(self, back_color_fixture):
        patt_fill, pattFill, expected_xml, color_ = back_color_fixture[:4]
        ColorFormat_from_colorchoice_parent_ = back_color_fixture[4]

        color = patt_fill.back_color

        assert pattFill.xml == expected_xml
        ColorFormat_from_colorchoice_parent_.assert_called_once_with(
            pattFill.bgClr
        )
        assert color is color_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('a:pattFill',
         'a:pattFill/a:bgClr/a:srgbClr{val=FFFFFF}'),
        ('a:pattFill/a:bgClr',
         'a:pattFill/a:bgClr'),
        ('a:pattFill/a:bgClr/a:schemeClr{val=accent1}',
         'a:pattFill/a:bgClr/a:schemeClr{val=accent1}'),
    ])
    def back_color_fixture(self, request, color_,
                           ColorFormat_from_colorchoice_parent_):
        pattFill_cxml, expected_cxml = request.param
        pattFill = element(pattFill_cxml)
        expected_xml = xml(expected_cxml)
        ColorFormat_from_colorchoice_parent_.return_value = color_

        patt_fill = _PattFill(pattFill)
        return (
            patt_fill, pattFill, expected_xml, color_,
            ColorFormat_from_colorchoice_parent_
        )

    @pytest.fixture(params=[
        ('a:pattFill',
         'a:pattFill/a:fgClr/a:srgbClr{val=000000}'),
        ('a:pattFill/a:fgClr',
         'a:pattFill/a:fgClr'),
        ('a:pattFill/a:fgClr/a:schemeClr{val=accent2}',
         'a:pattFill/a:fgClr/a:schemeClr{val=accent2}'),
    ])
    def fore_color_fixture(self, request, color_,
                           ColorFormat_from_colorchoice_parent_):
        pattFill_cxml, expected_cxml = request.param
        pattFill = element(pattFill_cxml)
        expected_xml = xml(expected_cxml)
        ColorFormat_from_colorchoice_parent_.return_value = color_

        patt_fill = _PattFill(pattFill)
        return (
            patt_fill, pattFill, expected_xml, color_,
            ColorFormat_from_colorchoice_parent_
        )

    @pytest.fixture
    def fill_type_fixture(self):
        patt_fill = _PattFill(element('a:pattFill'))
        expected_value = MSO_FILL.PATTERNED
        return patt_fill, expected_value

    @pytest.fixture(params=[
        ('a:pattFill',                 None),
        ('a:pattFill{prst=diagCross}', MSO_PATTERN.DIAGONAL_CROSS),
        ('a:pattFill{prst=wave}',      MSO_PATTERN.WAVE),
    ])
    def pattern_get_fixture(self, request):
        pattFill_cxml, expected_value = request.param
        pattFill = element(pattFill_cxml)

        patt_fill = _PattFill(pattFill)
        return patt_fill, expected_value

    @pytest.fixture(params=[
        ('a:pattFill',             MSO_PATTERN.WAVE,
         'a:pattFill{prst=wave}'),
        ('a:pattFill{prst=wave}',  MSO_PATTERN.DIVOT,
         'a:pattFill{prst=divot}'),
        ('a:pattFill{prst=divot}', None,
         'a:pattFill'),
    ])
    def pattern_set_fixture(self, request):
        pattFill_cxml, pattern, expected_cxml = request.param
        pattFill = element(pattFill_cxml)
        expected_xml = xml(expected_cxml)

        patt_fill = _PattFill(pattFill)
        return patt_fill, pattern, pattFill, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ColorFormat_from_colorchoice_parent_(self, request):
        return method_mock(request, ColorFormat, 'from_colorchoice_parent')

    @pytest.fixture
    def color_(self, request):
        return instance_mock(request, ColorFormat)


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
