# encoding: utf-8

"""Test suite for pptx.text.bullet module."""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.enum.text import AUTO_NUMBER_SCHEME
from pptx.dml.color import ColorFormat
from pptx.util import Pt

from pptx.text.bullets import (
    TextBullet,
    _Bullet,
    _NoBullet,
    _AutoNumBullet,
    _CharBullet,
    _PictureBullet,
    TextBulletColor,
    _TextBulletColor,
    _TextBulletColorFollowText,
    _TextBulletColorSpecific,
    TextBulletSize,
    _TextBulletSize,
    _TextBulletSizeFollowText,
    _TextBulletSizePercent,
    _TextBulletSizePoints,
    TextBulletTypeface,
    _BulletTypeface,
    _BulletTypefaceFollowText,
    _BulletTypefaceSpecific,
)

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock, property_mock, method_mock

class DescribeTextBullet(object):
    """ Unit-test suite for `pptx.text.bullets.TextBullet` object. """
    def it_can_set_the_bullet_to_no_bullet(self, no_bullet_fixture):
        text_bullet, _NoBullet_, expected_xml, no_bullet_ = no_bullet_fixture

        text_bullet.no_bullet()

        assert text_bullet._parent.xml == expected_xml
        _NoBullet_.assert_called_once_with(text_bullet._parent.eg_textBullet)
        assert text_bullet._bullet is no_bullet_

    def it_can_set_the_bullet_to_auto_number(self, auto_number_fixture):
        text_bullet, _AutoNumBullet_, expected_xml, auto_num_bullet_ = auto_number_fixture

        text_bullet.auto_number()

        assert text_bullet._parent.xml == expected_xml
        _AutoNumBullet_.assert_called_once_with(text_bullet._parent.eg_textBullet)
        assert text_bullet._bullet is auto_num_bullet_

    def it_can_set_the_bullet_to_a_character(self, character_fixture):
        text_bullet, _CharBullet_, expecte_xml, char_bullet_ = character_fixture

        text_bullet.character()

        assert text_bullet._parent.xml == expecte_xml
        _CharBullet_.assert_called_once_with(text_bullet._parent.eg_textBullet)
        assert text_bullet._bullet is char_bullet_

    def it_knows_its_bullet_type(self, type_fixture):
        text_bullet, expected_value = type_fixture
        bullet_type = text_bullet.type
        assert bullet_type == expected_value

    def it_knows_its_character_type(self, auto_num_bullet_, type_prop_):
        auto_num_bullet_.char_type = AUTO_NUMBER_SCHEME.ARABIC_PERIOD
        type_prop_.return_value = "AutoNumBullet"
        text_bullet = TextBullet(None, auto_num_bullet_)
        char_type = text_bullet.char_type
        assert char_type == AUTO_NUMBER_SCHEME.ARABIC_PERIOD

    def it_can_change_its_character_type(self, auto_num_bullet_, type_prop_):
        type_prop_.return_value = "AutoNumBullet"
        text_bullet = TextBullet(None, auto_num_bullet_)
        text_bullet.char_type = AUTO_NUMBER_SCHEME.ALPHA_UPPER_CHARACTER_PAREN_BOTH
        assert auto_num_bullet_.char_type == AUTO_NUMBER_SCHEME.ALPHA_UPPER_CHARACTER_PAREN_BOTH

    def it_knows_its_start_at(self, auto_num_bullet_, type_prop_):
        auto_num_bullet_.start_at = 42
        type_prop_.return_value = "AutoNumBullet"
        text_bullet = TextBullet(None, auto_num_bullet_)
        start_at = text_bullet.start_at
        assert start_at == 42

    def it_can_change_its_start_at(self, auto_num_bullet_, type_prop_):
        type_prop_.return_value = "AutoNumBullet"
        text_bullet = TextBullet(None, auto_num_bullet_)
        text_bullet.start_at = 42
        assert auto_num_bullet_.start_at == 42

    def it_knows_its_char(self, char_bullet_, type_prop_):
        char_bullet_.char = "-"
        type_prop_.return_value = "CharBullet"
        text_bullet = TextBullet(None, char_bullet_)
        char = text_bullet.char
        assert char == "-"

    def it_can_change_its_char(self, char_bullet_, type_prop_):
        type_prop_.return_value = "CharBullet"
        text_bullet = TextBullet(None, char_bullet_)
        text_bullet.char = "|"
        assert char_bullet_.char == "|"

    

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buNone"),
            ("a:pPr/a:buChar", "a:pPr/a:buNone"),
        ]
    )  
    def no_bullet_fixture(self, request, no_bullet_):
        cxml, expected_cxml = request.param

        text_bullet = TextBullet.from_parent(element(cxml))
        _NoBullet_ = class_mock(request, "pptx.text.bullets._NoBullet", return_value=no_bullet_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet, _NoBullet_, expected_xml, no_bullet_

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buAutoNum"),
            ("a:pPr/a:buChar", "a:pPr/a:buAutoNum"),
        ]
    )
    def auto_number_fixture(self, request, auto_num_bullet_):
        cxml, expected_cxml = request.param

        text_bullet = TextBullet.from_parent(element(cxml))
        _AutoNumBullet_ = class_mock(request, "pptx.text.bullets._AutoNumBullet", return_value=auto_num_bullet_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet, _AutoNumBullet_, expected_xml, auto_num_bullet_

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buChar"),
            ("a:pPr/a:buNone", "a:pPr/a:buChar"),
            ("a:pPr/a:buAutoNum", "a:pPr/a:buChar"),
        ]
    )
    def character_fixture(self, request, char_bullet_):
        cxml, expected_cxml = request.param

        text_bullet = TextBullet.from_parent(element(cxml))
        _CharBullet_ = class_mock(request, "pptx.text.bullets._CharBullet", return_value=char_bullet_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet, _CharBullet_, expected_xml, char_bullet_

    @pytest.fixture
    def type_fixture(self, text_bullet_):
        expected_value = text_bullet_.type = 42
        text_bullet = TextBullet(None, text_bullet_)
        return text_bullet, expected_value


    # fixture components ---------------------------------------------

    @pytest.fixture
    def no_bullet_(self, request):
        return instance_mock(request, _NoBullet)

    @pytest.fixture
    def auto_num_bullet_(self, request):
        return instance_mock(request, _AutoNumBullet)

    @pytest.fixture
    def char_bullet_(self, request):
        return instance_mock(request, _CharBullet)

    @pytest.fixture
    def text_bullet_(self, request):
        return instance_mock(request, TextBullet)

    @pytest.fixture
    def type_prop_(self, request):
        return property_mock(request, TextBullet, "type")


class Describe_Bullet(object):
    def it_raises_on_char_access(self, char_raise_fixture):
        bullet, exception_type = char_raise_fixture
        with pytest.raises(exception_type):
            bullet.char

    def it_raises_on_char_type_access(self, char_type_raise_fixture):
        bullet, exception_type = char_type_raise_fixture
        with pytest.raises(exception_type):
            bullet.char_type

    def it_raises_on_start_at_access(self, start_at_raise_fixture):
        bullet, exception_type = start_at_raise_fixture
        with pytest.raises(exception_type):
            bullet.start_at

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def char_raise_fixture(self):
        bullet = _Bullet("foobar")
        exception_type = TypeError
        return bullet, exception_type

    @pytest.fixture
    def char_type_raise_fixture(self):
        bullet = _Bullet("foobar")
        exception_type = TypeError
        return bullet, exception_type

    @pytest.fixture
    def start_at_raise_fixture(self):
        bullet = _Bullet("foobar")
        exception_type = TypeError
        return bullet, exception_type


class DescribeNoBullet(object):
    """ Unit-test suite for `pptx.text.bullets._NoBullet` object. """
    def it_knows_its_bullet_type(self, bullet_type_fixture):
        no_bullet, expected_value = bullet_type_fixture
        bullet_type = no_bullet.type
        assert bullet_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def bullet_type_fixture(self):
        xBullet = element("a:buNone")
        no_bullet = _NoBullet(xBullet)
        expected_value = "NoBullet"
        return no_bullet, expected_value


class DescribeAutoNumBullet(object):
    def it_knows_its_bullet_type(self, bullet_type_fixture):
        no_bullet, expected_value = bullet_type_fixture
        bullet_type = no_bullet.type
        assert bullet_type == expected_value

    def it_knows_its_start_at(self, start_at_get_fixture):
        auto_num_bullet, expected_value = start_at_get_fixture
        start_at = auto_num_bullet.start_at
        assert start_at == expected_value

    def it_can_change_its_start_at(self, start_at_set_fixture):
        auto_num_bullet, start_at, autoNumBullet, expected_xml = start_at_set_fixture
        auto_num_bullet.start_at = start_at
        assert autoNumBullet.xml == expected_xml
    
    def it_knows_its_char_type(self, char_type_get_fixture):
        auto_num_bullet, expected_value = char_type_get_fixture
        char_type = auto_num_bullet.char_type
        assert char_type == expected_value

    def it_can_change_its_char_type(self, char_type_set_fixture):
        auto_num_bullet, char_type, autoNumBullet, expected_xml = char_type_set_fixture
        auto_num_bullet.char_type = char_type
        assert autoNumBullet.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def bullet_type_fixture(self):
        xBullet = element("a:buAutoNum")
        auto_num_bullet = _AutoNumBullet(xBullet)
        expected_value = "AutoNumBullet"
        return auto_num_bullet, expected_value

    @pytest.fixture(
        params=[
            ("a:buAutoNum", 1),
            ("a:buAutoNum{startAt=10}", 10)
        ]
    )
    def start_at_get_fixture(self, request):
        autoNumBullet_cxml, expected_value = request.param
        autoNumBullet = element(autoNumBullet_cxml)

        auto_number_bullet = _AutoNumBullet(autoNumBullet)
        return auto_number_bullet, expected_value

    @pytest.fixture(
        params=[
            ("a:buAutoNum" ,5, "a:buAutoNum{startAt=5}"),
            ("a:buAutoNum{startAt=5}", 1, "a:buAutoNum"),
            ("a:buAutoNum", 12, "a:buAutoNum{startAt=12}"),
        ]
    )
    def start_at_set_fixture(self, request):
        autoNumBullet_cxml, start_at, expected_cxml = request.param
        autoNumBullet = element(autoNumBullet_cxml)
        expected_xml = xml(expected_cxml)

        auto_number_bullet = _AutoNumBullet(autoNumBullet)
        return auto_number_bullet, start_at, autoNumBullet, expected_xml


    @pytest.fixture(
        params=[
            ("a:buAutoNum{type=alphaLcParenBoth}", AUTO_NUMBER_SCHEME.ALPHA_LOWER_CHARACTER_PAREN_BOTH),
            ("a:buAutoNum{type=arabicPeriod}", AUTO_NUMBER_SCHEME.ARABIC_PERIOD),
        ]
    )
    def char_type_get_fixture(self, request):
        autoNumBullet_cxml, expected_value = request.param
        autoNumBullet = element(autoNumBullet_cxml)

        auto_number_bullet = _AutoNumBullet(autoNumBullet)
        return auto_number_bullet, expected_value

    @pytest.fixture(
        params=[
            ("a:buAutoNum{type=alphaLcParenBoth}", AUTO_NUMBER_SCHEME.ARABIC_1_MINUS, "a:buAutoNum{type=arabic1Minus}"),
            ("a:buAutoNum{type=alphaLcParenBoth}", AUTO_NUMBER_SCHEME.ROMAN_UPPER_CHARACTER_PERIOD, "a:buAutoNum{type=romanUcPeriod}"),
        ]
    )
    def char_type_set_fixture(self, request):
        autoNumBullet_cxml, char_type, expected_cxml = request.param
        autoNumBullet = element(autoNumBullet_cxml)
        expected_xml = xml(expected_cxml)

        auto_number_bullet = _AutoNumBullet(autoNumBullet)
        return auto_number_bullet, char_type, autoNumBullet, expected_xml


class DescribeCharBullet(object):
    def it_knows_its_bullet_type(self, bullet_type_fixture):
        no_bullet, expected_value = bullet_type_fixture
        bullet_type = no_bullet.type
        assert bullet_type == expected_value

    def it_knows_its_char(self, char_get_fixture):
        char_bullet, expected_value = char_get_fixture
        char = char_bullet.char
        assert char == expected_value

    def it_can_change_its_char(self, char_set_fixture):
        char_bullet, char, charBullet, expected_xml = char_set_fixture
        char_bullet.char = char
        assert charBullet.xml == expected_xml
    

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def bullet_type_fixture(self):
        xBullet = element("a:buChar")
        char_bullet = _NoBullet(xBullet)
        expected_value = "CharBullet"
        return char_bullet, expected_value

    
    @pytest.fixture(
        params=[
            ("a:buChar{char=-}", "-"),
            ("a:buChar{char=_}", "_"),
        ]
    )
    def char_get_fixture(self, request):
        charBullet_cxml, expected_value = request.param
        charBullet = element(charBullet_cxml)

        char_bullet = _CharBullet(charBullet)
        return char_bullet, expected_value

    @pytest.fixture(
        params=[
            ("a:buChar{char=-}", "_", "a:buChar{char=_}"),
            ("a:buChar{char=_}", "-", "a:buChar{char=-}"),
        ]
    )
    def char_set_fixture(self, request):
        charBullet_cxml, char, expected_cxml = request.param
        charBullet = element(charBullet_cxml)
        expected_xml = xml(expected_cxml)

        char_bullet = _CharBullet(charBullet)
        return char_bullet, char, charBullet, expected_xml


class DescribePictureBullet(object):
    def it_knows_its_bullet_type(self, bullet_type_fixture):
        picture_bullet, expected_value = bullet_type_fixture
        bullet_type = picture_bullet.type
        assert bullet_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def bullet_type_fixture(self):
        xBullet = element("a:buBlip")
        picture_bullet = _PictureBullet(xBullet)
        expected_value = "PictureBullet"
        return picture_bullet, expected_value


class DescribeTextBulletColor(object):
    """ Unit-test suite for `pptx.text.bullets.TextBulletColor` object. """
    def it_can_set_the_color_to_follow_text(self, follow_text_fixture):
        text_bullet_color, _TextBulletColorFollowText_, expected_xml, text_bullet_color_follow_text_ = follow_text_fixture

        text_bullet_color.follow_text()

        assert text_bullet_color._parent.xml == expected_xml
        _TextBulletColorFollowText_.assert_called_once_with(text_bullet_color._parent.eg_textBulletColor)
        assert text_bullet_color._bullet_color is text_bullet_color_follow_text_

    def it_can_set_the_color_to_a_specific_color(self, specific_color_fixture):
        text_bullet_color, _TextBulletColorSpecific_, expected_xml, text_bullet_color_specific_ = specific_color_fixture

        text_bullet_color.set_color()

        assert text_bullet_color._parent.xml == expected_xml
        _TextBulletColorSpecific_.assert_called_once_with(text_bullet_color._parent.eg_textBulletColor)
        assert text_bullet_color._bullet_color is text_bullet_color_specific_

    def it_knows_its_bullet_color_type(self, type_fixture):
        text_bullet_color, expected_value = type_fixture
        bullet_color_type = text_bullet_color.type
        assert bullet_color_type == expected_value

    def it_gives_access_to_its_color(self, color_fixture):
        text_bullet_color, color_ = color_fixture
        color = text_bullet_color.color
        assert color is color_

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buClrTx"),
            ("a:pPr/a:buClr", "a:pPr/a:buClrTx")
        ]
    )
    def follow_text_fixture(self, request, text_bullet_color_follow_text_):
        cxml, expected_cxml = request.param

        text_color = TextBulletColor.from_parent(element(cxml))
        _TextBulletColorFollowText_ = class_mock(request, "pptx.text.bullets._TextBulletColorFollowText", return_value=text_bullet_color_follow_text_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_color, _TextBulletColorFollowText_, expected_xml, text_bullet_color_follow_text_

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buClr"),
            ("a:pPr/a:buClrTx", "a:pPr/a:buClr")
        ]
    )
    def specific_color_fixture(self, request, text_bullet_color_specific_):
        cxml, expected_cxml = request.param

        text_color = TextBulletColor.from_parent(element(cxml))
        _TextBulletColorSpecific_ = class_mock(request, "pptx.text.bullets._TextBulletColorSpecific", return_value=text_bullet_color_specific_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_color, _TextBulletColorSpecific_, expected_xml, text_bullet_color_specific_


    @pytest.fixture
    def type_fixture(self, text_bullet_color_):
        expected_value = text_bullet_color_.type = 42
        text_bullet_color = TextBulletColor(None, text_bullet_color_)
        return text_bullet_color, expected_value

    @pytest.fixture
    def color_fixture(self, text_bullet_color_, color_):
        text_bullet_color_.color = color_
        text_bullet_color = TextBulletColor(None, text_bullet_color_)
        return text_bullet_color, color_
    

    # fixture components ---------------------------------------------

    @pytest.fixture
    def color_(self, request):
        return instance_mock(request, ColorFormat)

    @pytest.fixture
    def text_bullet_color_(self, request):
        return instance_mock(request, TextBulletColor)

    @pytest.fixture
    def text_bullet_color_follow_text_(self, request):
        return instance_mock(request, _TextBulletColorFollowText)

    @pytest.fixture
    def text_bullet_color_specific_(self, request):
        return instance_mock(request, _TextBulletColorSpecific)


class Describe_TextBulletColor(object):
    """ Unit-test suite for `pptx.text.bullets._TextBulletColor` object. """

    def it_raises_on_color_access(self, color_raise_fixture):
        bullet_color, exception_type = color_raise_fixture
        with pytest.raises(exception_type):
            bullet_color.color

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def color_raise_fixture(self):
        bullet_color = _TextBulletColor("foobar")
        exception_type = TypeError
        return bullet_color, exception_type

class DescribeTextBullerColorFollowText(object):
    """ Unit-test suite for `pptx.text.bullets._TextBulletColorFollowText` object. """

    def it_knows_its_bullet_color_type(self, color_type_fixture):
        follow_text_color, expected_value = color_type_fixture
        color_type = follow_text_color.type
        assert color_type == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def color_type_fixture(self):
        xBulletColor = element("a:buClrTx")
        follow_text_color = _TextBulletColorFollowText(xBulletColor)
        expected_value = "TextBulletColorFollowText"
        return follow_text_color, expected_value

class DescribeTextBullerColorSpecific(object):
    """ Unit-test suite for `pptx.text.bullets._TextBulletColorSpecific` object. """

    def it_knows_its_bullet_color_type(self, color_type_fixture):
        specific_bullet_color, expected_value = color_type_fixture
        color_type = specific_bullet_color.type
        assert color_type == expected_value


    def it_gives_access_to_its_color(self, color_fixture):
        text_bullet_color_specific, textBulletColorSpecific, color_, ColorFormat_from_colorchoice_parent_ = color_fixture

        color = text_bullet_color_specific.color

        ColorFormat_from_colorchoice_parent_.assert_called_once_with(textBulletColorSpecific)
        assert color is color_


    # fixtures -------------------------------------------------------

    @pytest.fixture
    def color_type_fixture(self):
        xBulletColor = element("a:buClr")
        specific_bullet_color = _TextBulletColorSpecific(xBulletColor)
        expected_value = "TextBulletColorSpecific"
        return specific_bullet_color, expected_value

    @pytest.fixture
    def color_fixture(self, ColorFormat_from_colorchoice_parent_, color_):
        ColorFormat_from_colorchoice_parent_.return_value = color_
        textBulletColorSpecific = element("a:buClr")

        text_bullet_color_specific = _TextBulletColorSpecific(textBulletColorSpecific)
        return text_bullet_color_specific, textBulletColorSpecific, color_, ColorFormat_from_colorchoice_parent_
        

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ColorFormat_from_colorchoice_parent_(self, request):
        return method_mock(request, ColorFormat, "from_colorchoice_parent")

    @pytest.fixture
    def color_(self, request):
        return instance_mock(request, ColorFormat)

class DescribeTextBulletSize(object):
    """ Unit-test suite for `pptx.text.bullets.TextBulletSize` object. """
    def it_can_set_the_size_to_follow_text(self, follow_text_fixture):
        text_bullet_size, _TextBulletSizeFollowText_, expected_xml, follow_text_ = follow_text_fixture
        text_bullet_size.follow_text()
        assert text_bullet_size._parent.xml == expected_xml
        _TextBulletSizeFollowText_.assert_called_once_with(text_bullet_size._parent.eg_textBulletSize)

    def it_can_set_the_size_to_points(self, points_fixture):
        text_bullet_size, _TextBulletSizePoints_, expected_xml, size_points_ = points_fixture
        text_bullet_size.set_points()
        assert text_bullet_size._parent.xml == expected_xml
        _TextBulletSizePoints_.assert_called_once_with(text_bullet_size._parent.eg_textBulletSize)
    
    def it_can_set_the_size_to_percentage(self, percentage_fixture):
        text_bullet_size, _TextBulletSizePercent_, expected_xml, size_points_ = percentage_fixture
        text_bullet_size.set_percentage()
        assert text_bullet_size._parent.xml == expected_xml
        _TextBulletSizePercent_.assert_called_once_with(text_bullet_size._parent.eg_textBulletSize)
   
    def it_knows_its_points(self, points_size_, type_prop_):
        points_size_.points = Pt(12)
        type_prop_.return_value = "TextBulletSizePoints"
        text_bullet_size = TextBulletSize(None, points_size_)
        points = text_bullet_size.points
        assert points == Pt(12)
    
    def it_can_change_its_points(self, points_size_, type_prop_):
        type_prop_.return_value = "TextBulletSizePoints"
        text_bullet_size = TextBulletSize(None, points_size_)
        text_bullet_size.points = Pt(42)
        assert points_size_.points == Pt(42)

    def it_knows_its_percent(self, percentage_size_, type_prop_):
        percentage_size_.percentage = 150
        type_prop_.return_value = "TextBulletSizePercent"
        text_bullet_size = TextBulletSize(None, percentage_size_)
        percent = text_bullet_size.percentage
        assert percent == 150
    
    def it_can_change_its_percentage(self, percentage_size_, type_prop_):
        type_prop_.return_value = "TextBulletSizePercent"
        text_bullet_size = TextBulletSize(None, percentage_size_)
        text_bullet_size.percentage = 42
        assert percentage_size_.percentage == 42

    def it_knows_its_type(self, type_fixture):
        bullet_size, expected_value = type_fixture
        bullet_size_type = bullet_size.type
        assert bullet_size_type == expected_value
    
    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buSzTx"),
            ("a:pPr/a:buSzPct", "a:pPr/a:buSzTx")
        ]
    )
    def follow_text_fixture(self, request, follow_text_):
        cxml, expected_cxml = request.param
        text_bullet_size = TextBulletSize.from_parent(element(cxml))
        _TextBulletSizeFollowText_ = class_mock(request, "pptx.text.bullets._TextBulletSizeFollowText", return_value=follow_text_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet_size, _TextBulletSizeFollowText_, expected_xml, follow_text_

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buSzPts"),
            ("a:pPr/a:buSzPct", "a:pPr/a:buSzPts")
        ]
    )
    def points_fixture(self, request, points_size_):
        cxml, expected_cxml = request.param
        text_bullet_size = TextBulletSize.from_parent(element(cxml))
        _TextBulletSizePoints_ = class_mock(request, "pptx.text.bullets._TextBulletSizePoints", return_value=points_size_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet_size, _TextBulletSizePoints_, expected_xml, points_size_

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buSzPct"),
            ("a:pPr/a:buSzPts", "a:pPr/a:buSzPct")
        ]
    )
    def percentage_fixture(self, request, percentage_size_):
        cxml, expected_cxml = request.param
        text_bullet_size = TextBulletSize.from_parent(element(cxml))
        _TextBulletSizePercent_ = class_mock(request, "pptx.text.bullets._TextBulletSizePercent", return_value=percentage_size_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet_size, _TextBulletSizePercent_, expected_xml, percentage_size_

    @pytest.fixture
    def type_fixture(self, text_bullet_size_):
        expected_value = text_bullet_size_.type = 42
        text_bullet_size = TextBulletSize(None, text_bullet_size_)
        return text_bullet_size, expected_value



    # fixture components ---------------------------------------------

    @pytest.fixture
    def follow_text_(self, request):
        return instance_mock(request, _TextBulletSizeFollowText)

    @pytest.fixture
    def points_size_(self, request):
        return instance_mock(request, _TextBulletSizePoints)

    @pytest.fixture
    def percentage_size_(self, request):
        return instance_mock(request, _TextBulletSizePercent)

    @pytest.fixture
    def text_bullet_size_(self, request):
        return instance_mock(request, TextBulletSize)

    @pytest.fixture
    def type_prop_(self, request):
        return property_mock(request, TextBulletSize, "type")



class DescribeTextBulletTypeface(object):
    """ Unit-test suite for `pptx.text.bullets.TextBulletTypeface` object. """
    def it_can_set_the_size_to_follow_text(self, follow_text_fixture):
        text_bullet_typeface, _BulletTypefaceFollowText_, expected_xml, follow_text_ = follow_text_fixture
        text_bullet_typeface.follow_text()
        assert text_bullet_typeface._parent.xml == expected_xml
        _BulletTypefaceFollowText_.assert_called_once_with(text_bullet_typeface._parent.eg_textBulletTypeface)
        assert text_bullet_typeface._bullet_typeface is follow_text_


    def it_can_set_to_typeface(self, specific_typeface_fixture):
        text_bullet_typeface, _BulletTypefaceSpecific_, expected_xml, specific_typeface_ = specific_typeface_fixture
        text_bullet_typeface.set_typeface()
        assert text_bullet_typeface._parent.xml == expected_xml
        _BulletTypefaceSpecific_.assert_called_once_with(text_bullet_typeface._parent.eg_textBulletTypeface)
        assert text_bullet_typeface._bullet_typeface is specific_typeface_
    
    def it_knows_its_type(self, type_fixture):
        bullet_typeface, expected_value = type_fixture
        bullet_type = bullet_typeface.type
        assert bullet_type == expected_value


    def it_knows_its_typeface(self, text_bullet_typeface_, type_prop_):
        text_bullet_typeface_.typeface = "Foobar"
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        typeface = bullet_typeface.typeface
        assert typeface == "Foobar"
    
    def it_can_change_its_typeface(self, text_bullet_typeface_, type_prop_):
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        bullet_typeface.typeface = "Foobar"
        assert text_bullet_typeface_.typeface == "Foobar"

    def it_knows_its_pitch_family(self, text_bullet_typeface_, type_prop_):
        text_bullet_typeface_.pitch_family = "Foobar"
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        pitch_family = bullet_typeface.pitch_family
        assert pitch_family == "Foobar"
    
    def it_can_change_its_pitch_family(self, text_bullet_typeface_, type_prop_):
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        bullet_typeface.pitch_family = "Foobar"
        assert text_bullet_typeface_.pitch_family == "Foobar"

    def it_knows_its_panose(self, text_bullet_typeface_, type_prop_):
        text_bullet_typeface_.panose = "Foobar"
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        panose = bullet_typeface.panose
        assert panose == "Foobar"
    
    def it_can_change_its_panose(self, text_bullet_typeface_, type_prop_):
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        bullet_typeface.panose = "Foobar"
        assert text_bullet_typeface_.panose == "Foobar"

    def it_knows_its_charset(self, text_bullet_typeface_, type_prop_):
        text_bullet_typeface_.charset = "Foobar"
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        charset = bullet_typeface.charset
        assert charset == "Foobar"
    
    def it_can_change_its_charset(self, text_bullet_typeface_, type_prop_):
        type_prop_.return_value = "BulletTypefaceSpecific"
        bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        bullet_typeface.charset = "Foobar"
        assert text_bullet_typeface_.charset == "Foobar"

    
    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buFontTx"),
            ("a:pPr/a:buFont", "a:pPr/a:buFontTx")
        ]
    )
    def follow_text_fixture(self, request, follow_text_):
        cxml, expected_cxml = request.param
        text_bullet_typeface = TextBulletTypeface.from_parent(element(cxml))
        _BulletTypefaceFollowText_ = class_mock(request, "pptx.text.bullets._BulletTypefaceFollowText", return_value=follow_text_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet_typeface, _BulletTypefaceFollowText_, expected_xml, follow_text_

    @pytest.fixture(
        params = [
            ("a:pPr", "a:pPr/a:buFont"),
            ("a:pPr/a:buFontTx", "a:pPr/a:buFont")
        ]
    )
    def specific_typeface_fixture(self, request, specific_typeface_):
        cxml, expected_cxml = request.param
        text_bullet_typeface = TextBulletTypeface.from_parent(element(cxml))
        _BulletTypefaceSpecific_ = class_mock(request, "pptx.text.bullets._BulletTypefaceSpecific", return_value=specific_typeface_, autospec=True)
        expected_xml = xml(expected_cxml)
        return text_bullet_typeface, _BulletTypefaceSpecific_, expected_xml, specific_typeface_

    @pytest.fixture
    def type_fixture(self, text_bullet_typeface_):
        expected_value = text_bullet_typeface_.type = 42
        text_bullet_typeface = TextBulletTypeface(None, text_bullet_typeface_)
        return text_bullet_typeface, expected_value


    # fixture components ---------------------------------------------

    @pytest.fixture
    def follow_text_(self, request):
        return instance_mock(request, _BulletTypefaceFollowText)


    @pytest.fixture
    def specific_typeface_(self, request):
        return instance_mock(request, _BulletTypefaceSpecific)

    @pytest.fixture
    def text_bullet_typeface_(self, request):
        return instance_mock(request, TextBulletTypeface)

    @pytest.fixture
    def type_prop_(self, request):
        return property_mock(request, TextBulletTypeface, "type")

