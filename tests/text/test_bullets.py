# encoding: utf-8

"""Test suite for pptx.text.bullet module."""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.enum.text import AUTO_NUMBER_SCHEME
from pptx.text.bullets import (
    TextBullet,
    _Bullet,
    _NoBullet,
    _AutoNumBullet,
    _CharBullet,
    _PictureBullet,
    TextBulletColor,
    _TextBulletColorFollowText,
    _TextBulletColorSpecific,
    TextBulletSize,
    _TextBulletSizeFollowText,
    _TextBulletSizePercent,
    _TextBulletSizePoints,
    TextBulletTypeface,
    _BulletTypeface,
    _BulletTypefaceFollowText,
    _BulletTypefaceSpecific,
)

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock, property_mock

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
