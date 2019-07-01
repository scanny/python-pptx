# encoding: utf-8

"""Test suite for pptx.dml.fill module."""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.fill import (
    _BlipFill,
    _Fill,
    FillFormat,
    _GradFill,
    _GradientStop,
    _GradientStops,
    _GrpFill,
    _NoFill,
    _NoneFill,
    _PattFill,
    _SolidFill,
)
from pptx.enum.dml import MSO_FILL, MSO_PATTERN
from pptx.oxml.dml.fill import CT_GradientStopList

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock, method_mock, property_mock


class DescribeFillFormat(object):
    def it_can_set_the_fill_type_to_no_fill(self, background_fixture):
        fill, _NoFill_, expected_xml, no_fill_ = background_fixture

        fill.background()

        assert fill._xPr.xml == expected_xml
        _NoFill_.assert_called_once_with(fill._xPr.eg_fillProperties)
        assert fill._fill is no_fill_

    def it_can_set_the_fill_type_to_gradient(self, gradient_fixture):
        fill, _GradFill_, expected_xml, grad_fill_ = gradient_fixture

        fill.gradient()

        assert fill._xPr.xml == expected_xml
        _GradFill_.assert_called_once_with(fill._xPr.eg_fillProperties)
        assert fill._fill is grad_fill_

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

    def it_knows_the_angle_of_a_linear_gradient(self, grad_fill_, type_prop_):
        grad_fill_.gradient_angle = 42.0
        type_prop_.return_value = MSO_FILL.GRADIENT
        fill = FillFormat(None, grad_fill_)

        angle = fill.gradient_angle

        assert angle == 42.0

    def it_can_change_the_angle_of_a_linear_gradient(self, grad_fill_, type_prop_):
        type_prop_.return_value = MSO_FILL.GRADIENT
        fill = FillFormat(None, grad_fill_)

        fill.gradient_angle = 42.24

        assert grad_fill_.gradient_angle == 42.24

    def it_provides_access_to_the_gradient_stops(
        self, type_prop_, grad_fill_, gradient_stops_
    ):
        type_prop_.return_value = MSO_FILL.GRADIENT
        grad_fill_.gradient_stops = gradient_stops_
        fill = FillFormat(None, grad_fill_)

        gradient_stops = fill.gradient_stops

        assert gradient_stops is gradient_stops_

    def it_raises_on_non_gradient_fill(self, grad_fill_, type_prop_):
        type_prop_.return_value = None
        fill = FillFormat(None, grad_fill_)
        with pytest.raises(TypeError):
            fill.gradient_angle
        with pytest.raises(TypeError):
            fill.gradient_angle = 123.4
        with pytest.raises(TypeError):
            fill.gradient_stops

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

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", "p:spPr{a:b=c}/a:noFill"),
            ("a:tcPr/a:blipFill", "a:tcPr/a:noFill"),
            ("a:rPr/a:gradFill", "a:rPr/a:noFill"),
            ("a:tcPr/a:grpFill", "a:tcPr/a:noFill"),
            ("a:defRPr/a:noFill", "a:defRPr/a:noFill"),
            ("a:endParaRPr/a:pattFill", "a:endParaRPr/a:noFill"),
            ("a:rPr/a:solidFill", "a:rPr/a:noFill"),
        ]
    )
    def background_fixture(self, request, no_fill_):
        cxml, expected_cxml = request.param

        fill = FillFormat.from_fill_parent(element(cxml))
        # --mock must be after FillFormat call to avoid poss. contructor call
        _NoFill_ = class_mock(
            request, "pptx.dml.fill._NoFill", return_value=no_fill_, autospec=True
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

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", "p:spPr{a:b=c}/a:gradFill"),
            ("p:bgPr/a:noFill", "p:bgPr/a:gradFill"),
            ("a:tcPr/a:blipFill", "a:tcPr/a:gradFill"),
            ("a:rPr/a:grpFill", "a:rPr/a:gradFill"),
            ("a:tcPr/a:noFill", "a:tcPr/a:gradFill"),
            ("a:defRPr/a:solidFill", "a:defRPr/a:gradFill"),
            ("a:endParaRPr/a:blipFill", "a:endParaRPr/a:gradFill"),
        ]
    )
    def gradient_fixture(self, request, grad_fill_):
        cxml, expected_cxml = request.param

        fill = FillFormat.from_fill_parent(element(cxml))
        # --mock must be after FillFormat call to avoid poss. contructor call
        _GradFill_ = class_mock(
            request, "pptx.dml.fill._GradFill", return_value=grad_fill_, autospec=True
        )
        expected_xml = xml(
            expected_cxml + "{rotWithShape=1}/(a:gsLst/(a:gs{pos=0}/a:scheme"
            "Clr{val=accent1}/(a:tint{val=100000},a:shade{val=100000},a:satM"
            "od{val=130000}),a:gs{pos=100000}/a:schemeClr{val=accent1}/(a:ti"
            "nt{val=50000},a:shade{val=100000},a:satMod{val=350000})),a:lin{"
            "scaled=0})"
        )
        return fill, _GradFill_, expected_xml, grad_fill_

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

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", "p:spPr{a:b=c}/a:pattFill"),
            ("a:tcPr/a:blipFill", "a:tcPr/a:pattFill"),
            ("a:rPr/a:gradFill", "a:rPr/a:pattFill"),
            ("a:tcPr/a:grpFill", "a:tcPr/a:pattFill"),
            ("a:defRPr/a:noFill", "a:defRPr/a:pattFill"),
            ("a:endParaRPr/a:pattFill", "a:endParaRPr/a:pattFill"),
            ("a:rPr/a:solidFill", "a:rPr/a:pattFill"),
        ]
    )
    def patterned_fixture(self, request, patt_fill_):
        cxml, expected_cxml = request.param

        fill = FillFormat.from_fill_parent(element(cxml))
        # --mock must be after FillFormat call to avoid poss. contructor call
        _PattFill_ = class_mock(
            request, "pptx.dml.fill._PattFill", return_value=patt_fill_, autospec=True
        )
        expected_xml = xml(expected_cxml)
        return fill, _PattFill_, expected_xml, patt_fill_

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", "p:spPr{a:b=c}/a:solidFill"),
            ("a:tcPr/a:blipFill", "a:tcPr/a:solidFill"),
            ("a:rPr/a:gradFill", "a:rPr/a:solidFill"),
            ("a:tcPr/a:grpFill", "a:tcPr/a:solidFill"),
            ("a:defRPr/a:noFill", "a:defRPr/a:solidFill"),
            ("a:rPr/a:pattFill", "a:rPr/a:solidFill"),
            ("a:endParaRPr/a:solidFill", "a:endParaRPr/a:solidFill"),
        ]
    )
    def solid_fixture(self, request, solid_fill_):
        cxml, expected_cxml = request.param

        fill = FillFormat.from_fill_parent(element(cxml))
        # --mock must be after FillFormat call to avoid poss. contructor call
        _SolidFill_ = class_mock(
            request, "pptx.dml.fill._SolidFill", return_value=solid_fill_, autospec=True
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
    def grad_fill_(self, request):
        return instance_mock(request, _GradFill)

    @pytest.fixture
    def gradient_stops_(self, request):
        return instance_mock(request, _GradientStops)

    @pytest.fixture
    def no_fill_(self, request):
        return instance_mock(request, _NoFill)

    @pytest.fixture
    def patt_fill_(self, request):
        return instance_mock(request, _PattFill)

    @pytest.fixture
    def solid_fill_(self, request):
        return instance_mock(request, _SolidFill)

    @pytest.fixture
    def type_prop_(self, request):
        return property_mock(request, FillFormat, "type")


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
        fill = _Fill("foobar")
        exception_type = TypeError
        return fill, exception_type

    @pytest.fixture
    def fore_raise_fixture(self):
        fill = _Fill("foobar")
        exception_type = TypeError
        return fill, exception_type

    @pytest.fixture
    def pattern_raise_fixture(self):
        fill = _Fill("barfoo")
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
        xFill = element("a:blipFill")
        blip_fill = _BlipFill(xFill)
        expected_value = MSO_FILL.PICTURE
        return blip_fill, expected_value


class Describe_GradFill(object):
    def it_knows_the_gradient_angle(self, angle_fixture):
        grad_fill, expected_value = angle_fixture
        angle = grad_fill.gradient_angle
        assert angle == expected_value

    def it_can_change_the_gradient_angle(self, angle_set_fixture):
        grad_fill, angle, expected_xml = angle_set_fixture
        grad_fill.gradient_angle = angle
        assert grad_fill._gradFill.xml == expected_xml

    def it_provides_access_to_the_gradient_stops(self, stops_fixture):
        grad_fill, _GradientStops_ = stops_fixture[:2]
        expected_xml, gradient_stops_ = stops_fixture[2:]

        gradient_stops = grad_fill.gradient_stops

        gradFill = grad_fill._gradFill
        _GradientStops_.assert_called_once_with(gradFill[0])
        assert gradFill.xml == expected_xml
        assert gradient_stops is gradient_stops_

    def it_raises_on_non_linear_gradient(self):
        gradFill = element("a:gradFill/a:path")
        grad_fill = _GradFill(gradFill)
        with pytest.raises(ValueError):
            grad_fill.gradient_angle
        with pytest.raises(ValueError):
            grad_fill.gradient_angle = 43.21

    def it_knows_its_fill_type(self, fill_type_fixture):
        grad_fill, expected_value = fill_type_fixture
        fill_type = grad_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("a:gradFill", None),
            ("a:gradFill/a:lin{ang=0}", 0.0),
            ("a:gradFill/a:lin{ang=2730000}", 314.5),
            ("a:gradFill/a:lin{ang=16200000}", 90.0),
        ]
    )
    def angle_fixture(self, request):
        cxml, expected_value = request.param
        grad_fill = _GradFill(element(cxml))
        return grad_fill, expected_value

    @pytest.fixture(
        params=[
            (301.2, "a:gradFill/a:lin{ang=3528000}"),
            (31.22, "a:gradFill/a:lin{ang=19726800}"),
            (0, "a:gradFill/a:lin{ang=0}"),
            (-460.0, "a:gradFill/a:lin{ang=6000000}"),
        ]
    )
    def angle_set_fixture(self, request):
        angle, expected_cxml = request.param
        grad_fill = _GradFill(element("a:gradFill/a:lin{ang=0}"))
        expected_xml = xml(expected_cxml)
        return grad_fill, angle, expected_xml

    @pytest.fixture
    def fill_type_fixture(self):
        xFill = element("a:gradFill")
        grad_fill = _GradFill(xFill)
        expected_value = MSO_FILL.GRADIENT
        return grad_fill, expected_value

    @pytest.fixture(
        params=[
            ("a:gradFill", "a:gradFill", True),
            ("a:gradFill/a:gsLst", "a:gradFill/a:gsLst", False),
        ]
    )
    def stops_fixture(self, request, _GradientStops_, gradient_stops_):
        gradFill_cxml, expected_cxml, add_gsLst = request.param
        _GradientStops_.return_value = gradient_stops_
        grad_fill = _GradFill(element(gradFill_cxml))
        if add_gsLst:
            expected_xml = xml(
                expected_cxml + "/a:gsLst/(a:gs{pos=0}/a:schemeClr{val=accen"
                "t1}/(a:tint{val=100000},a:shade{val=100000},a:satMod{val=13"
                "0000}),a:gs{pos=100000}/a:schemeClr{val=accent1}/(a:tint{va"
                "l=50000},a:shade{val=100000},a:satMod{val=350000}))"
            )
        else:
            expected_xml = xml(expected_cxml)
        return grad_fill, _GradientStops_, expected_xml, gradient_stops_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _GradientStops_(self, request):
        return class_mock(request, "pptx.dml.fill._GradientStops")

    @pytest.fixture
    def gradient_stops_(self, request):
        return instance_mock(request, _GradientStops)


class Describe_GrpFill(object):
    def it_knows_its_fill_type(self, fill_type_fixture):
        grp_fill, expected_value = fill_type_fixture
        fill_type = grp_fill.type
        assert fill_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        xFill = element("a:grpFill")
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
        xFill = element("a:noFill")
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
        ColorFormat_from_colorchoice_parent_.assert_called_once_with(pattFill.fgClr)
        assert color is color_

    def it_provides_access_to_its_background_color(self, back_color_fixture):
        patt_fill, pattFill, expected_xml, color_ = back_color_fixture[:4]
        ColorFormat_from_colorchoice_parent_ = back_color_fixture[4]

        color = patt_fill.back_color

        assert pattFill.xml == expected_xml
        ColorFormat_from_colorchoice_parent_.assert_called_once_with(pattFill.bgClr)
        assert color is color_

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("a:pattFill", "a:pattFill/a:bgClr/a:srgbClr{val=FFFFFF}"),
            ("a:pattFill/a:bgClr", "a:pattFill/a:bgClr"),
            (
                "a:pattFill/a:bgClr/a:schemeClr{val=accent1}",
                "a:pattFill/a:bgClr/a:schemeClr{val=accent1}",
            ),
        ]
    )
    def back_color_fixture(self, request, color_, ColorFormat_from_colorchoice_parent_):
        pattFill_cxml, expected_cxml = request.param
        pattFill = element(pattFill_cxml)
        expected_xml = xml(expected_cxml)
        ColorFormat_from_colorchoice_parent_.return_value = color_

        patt_fill = _PattFill(pattFill)
        return (
            patt_fill,
            pattFill,
            expected_xml,
            color_,
            ColorFormat_from_colorchoice_parent_,
        )

    @pytest.fixture(
        params=[
            ("a:pattFill", "a:pattFill/a:fgClr/a:srgbClr{val=000000}"),
            ("a:pattFill/a:fgClr", "a:pattFill/a:fgClr"),
            (
                "a:pattFill/a:fgClr/a:schemeClr{val=accent2}",
                "a:pattFill/a:fgClr/a:schemeClr{val=accent2}",
            ),
        ]
    )
    def fore_color_fixture(self, request, color_, ColorFormat_from_colorchoice_parent_):
        pattFill_cxml, expected_cxml = request.param
        pattFill = element(pattFill_cxml)
        expected_xml = xml(expected_cxml)
        ColorFormat_from_colorchoice_parent_.return_value = color_

        patt_fill = _PattFill(pattFill)
        return (
            patt_fill,
            pattFill,
            expected_xml,
            color_,
            ColorFormat_from_colorchoice_parent_,
        )

    @pytest.fixture
    def fill_type_fixture(self):
        patt_fill = _PattFill(element("a:pattFill"))
        expected_value = MSO_FILL.PATTERNED
        return patt_fill, expected_value

    @pytest.fixture(
        params=[
            ("a:pattFill", None),
            ("a:pattFill{prst=diagCross}", MSO_PATTERN.DIAGONAL_CROSS),
            ("a:pattFill{prst=wave}", MSO_PATTERN.WAVE),
        ]
    )
    def pattern_get_fixture(self, request):
        pattFill_cxml, expected_value = request.param
        pattFill = element(pattFill_cxml)

        patt_fill = _PattFill(pattFill)
        return patt_fill, expected_value

    @pytest.fixture(
        params=[
            ("a:pattFill", MSO_PATTERN.WAVE, "a:pattFill{prst=wave}"),
            ("a:pattFill{prst=wave}", MSO_PATTERN.DIVOT, "a:pattFill{prst=divot}"),
            ("a:pattFill{prst=divot}", None, "a:pattFill"),
        ]
    )
    def pattern_set_fixture(self, request):
        pattFill_cxml, pattern, expected_cxml = request.param
        pattFill = element(pattFill_cxml)
        expected_xml = xml(expected_cxml)

        patt_fill = _PattFill(pattFill)
        return patt_fill, pattern, pattFill, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ColorFormat_from_colorchoice_parent_(self, request):
        return method_mock(request, ColorFormat, "from_colorchoice_parent")

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

        ColorFormat_from_colorchoice_parent_.assert_called_once_with(solidFill)
        assert color is color_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_type_fixture(self):
        solid_fill = _SolidFill(element("a:solidFill"))
        expected_value = MSO_FILL.SOLID
        return solid_fill, expected_value

    @pytest.fixture
    def fore_color_fixture(self, ColorFormat_from_colorchoice_parent_, color_):
        ColorFormat_from_colorchoice_parent_.return_value = color_
        solidFill = element("a:solidFill")

        solid_fill = _SolidFill(solidFill)
        return (solid_fill, solidFill, color_, ColorFormat_from_colorchoice_parent_)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ColorFormat_from_colorchoice_parent_(self, request):
        return method_mock(request, ColorFormat, "from_colorchoice_parent")

    @pytest.fixture
    def color_(self, request):
        return instance_mock(request, ColorFormat)


class Describe_GradientStops(object):
    def it_provides_access_to_its_stops(self):
        stops = _GradientStops(CT_GradientStopList.new_gsLst())

        assert len(stops) == 2
        for stop in stops:
            assert isinstance(stop, _GradientStop)


class Describe_GradientStop(object):
    def it_provides_access_to_its_color(self, request):
        gs = element("a:gs")
        ColorFormat_ = class_mock(request, "pptx.dml.fill.ColorFormat")
        color_ = instance_mock(request, ColorFormat)
        ColorFormat_.from_colorchoice_parent.return_value = color_
        stop = _GradientStop(gs)

        color = stop.color

        ColorFormat_.from_colorchoice_parent.assert_called_once_with(gs)
        assert color is color_

    def it_knows_its_position(self, pos_get_fixture):
        stop, expected_value = pos_get_fixture
        position = stop.position
        assert position == expected_value

    def it_can_change_its_position(self, pos_set_fixture):
        stop, new_value, expected_value = pos_set_fixture
        stop.position = new_value
        assert stop.position == expected_value

    def it_raises_on_position_out_of_range(self, raises_fixture):
        stop, out_of_range_value = raises_fixture
        with pytest.raises(ValueError):
            stop.position = out_of_range_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("a:gs{pos=0}", 0.0),
            ("a:gs{pos=42240}", 0.4224),
            ("a:gs{pos=100000}", 1.0),
        ]
    )
    def pos_get_fixture(self, request):
        gs_cxml, expected_value = request.param
        stop = _GradientStop(element(gs_cxml))
        return stop, expected_value

    @pytest.fixture(
        params=[
            ("a:gs{pos=0}", 0.42, 0.42),
            ("a:gs{pos=42000}", 1.0, 1.0),
            ("a:gs{pos=100000}", 0.0, 0.0),
        ]
    )
    def pos_set_fixture(self, request):
        gs_cxml, new_value, expected_value = request.param
        stop = _GradientStop(element(gs_cxml))
        return stop, new_value, expected_value

    @pytest.fixture(params=[-0.42, 1.001])
    def raises_fixture(self, request):
        out_of_range_value = request.param
        stop = _GradientStop(element("a:gs{pos=50000}"))
        return stop, out_of_range_value
