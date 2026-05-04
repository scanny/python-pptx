"""Unit-test suite for `pptx.action` module."""

from __future__ import annotations

import io

import pytest

from pptx.action import ActionSetting, Hyperlink, Sound
from pptx.enum.action import PP_ACTION
from pptx.media import Audio
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.parts.slide import SlidePart
from pptx.slide import Slide

from .unitutil.cxml import element, xml
from .unitutil.mock import call, class_mock, instance_mock, method_mock, property_mock


class DescribeActionSetting(object):
    """Unit-test suite for `pptx.action.ActionSetting` objects."""

    def it_knows_its_action_type(self, action_fixture):
        action_setting, expected_action = action_fixture
        action = action_setting.action
        assert action is expected_action

    def it_provides_access_to_its_hyperlink(self, hyperlink_fixture):
        action_setting, hyperlink_, Hyperlink_ = hyperlink_fixture[:3]
        xPr, parent = hyperlink_fixture[3:]
        hyperlink = action_setting.hyperlink
        Hyperlink_.assert_called_once_with(xPr, parent, False)
        assert hyperlink is hyperlink_

    def it_threads_the_hover_flag_through_to_its_hyperlink(
        self, Hyperlink_, hyperlink_
    ):
        xPr, parent = "xPr", "parent"
        action_setting = ActionSetting(xPr, parent, hover=True)

        hyperlink = action_setting.hyperlink

        Hyperlink_.assert_called_once_with(xPr, parent, True)
        assert hyperlink is hyperlink_

    def it_reads_the_action_type_from_the_hover_hyperlink(self):
        cNvPr = element("p:cNvPr/a:hlinkHover{r:id=rId6}")
        cNvPr.hlinkHover.action = "ppaction://hlinkshowjump?jump=nextslide"

        action_setting = ActionSetting(cNvPr, None, hover=True)

        assert action_setting.action is PP_ACTION.NEXT_SLIDE

    def it_returns_NONE_for_hover_action_when_no_hlinkHover_is_present(self):
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId1}")
        action_setting = ActionSetting(cNvPr, None, hover=True)

        # -- a plain hlinkClick does not count as a hover action --
        assert action_setting.action is PP_ACTION.NONE

    def it_can_find_its_slide_jump_target(self, target_get_fixture):
        action_setting, expected_value = target_get_fixture
        target_slide = action_setting.target_slide
        assert target_slide == expected_value

    def it_can_change_its_slide_jump_target(
        self, request, _clear_click_action_, slide_, part_prop_, part_
    ):
        part_prop_.return_value = part_
        part_.relate_to.return_value = "rId42"
        slide_part_ = instance_mock(request, SlidePart)
        slide_.part = slide_part_
        action_setting = ActionSetting(element("p:cNvPr{a:b=c,r:s=t}"), None)

        action_setting.target_slide = slide_

        _clear_click_action_.assert_called_once_with(action_setting)
        part_.relate_to.assert_called_once_with(slide_part_, RT.SLIDE)
        assert action_setting._element.xml == xml(
            "p:cNvPr{a:b=c,r:s=t}/a:hlinkClick{action=ppaction://hlinksldjump,r:id=rI" "d42}",
        )

    def but_it_clears_the_target_slide_if_None_is_assigned(self, _clear_click_action_):
        action_setting = ActionSetting(element("p:cNvPr{a:b=c,r:s=t}"), None)

        action_setting.target_slide = None

        _clear_click_action_.assert_called_once_with(action_setting)
        assert action_setting._element.xml == xml("p:cNvPr{a:b=c,r:s=t}")

    def it_sets_target_slide_on_hlinkHover_when_hover_is_True(
        self, request, _clear_click_action_, slide_, part_prop_, part_
    ):
        part_prop_.return_value = part_
        part_.relate_to.return_value = "rId42"
        slide_part_ = instance_mock(request, SlidePart)
        slide_.part = slide_part_
        action_setting = ActionSetting(
            element("p:cNvPr{a:b=c,r:s=t}"), None, hover=True
        )

        action_setting.target_slide = slide_

        _clear_click_action_.assert_called_once_with(action_setting)
        part_.relate_to.assert_called_once_with(slide_part_, RT.SLIDE)
        assert action_setting._element.xml == xml(
            "p:cNvPr{a:b=c,r:s=t}/a:hlinkHover{action=ppaction://hlinksldjump,"
            "r:id=rId42}",
        )

    def it_raises_on_no_next_prev_slide(self, target_raise_fixture):
        action_setting = target_raise_fixture
        with pytest.raises(ValueError):
            action_setting.target_slide

    def it_knows_its_slide_index_to_help(self, _slide_index_fixture):
        action_setting, slides, slide, expected_value = _slide_index_fixture
        slide_index = action_setting._slide_index
        slides.index.assert_called_once_with(slide)
        assert slide_index == expected_value

    def it_clears_the_click_action_to_help(self, clear_fixture):
        action_setting, calls, expected_xml = clear_fixture

        action_setting._clear_click_action()

        assert action_setting.part.drop_rel.call_args_list == calls
        assert action_setting._element.xml == expected_xml

    def it_also_drops_an_audio_rel_when_clearing_the_click_action(self, part_prop_, part_):
        part_prop_.return_value = part_
        cNvPr = element(
            "p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{r:id=rId3}"
            "/a:snd{r:embed=rId4,name=applause.wav}"
        )
        action_setting = ActionSetting(cNvPr, None)

        action_setting._clear_click_action()

        assert part_.drop_rel.call_args_list == [call("rId3"), call("rId4")]
        assert action_setting._element.xml == xml("p:cNvPr{a:a=a,r:r=r}")

    def it_clears_the_hlinkHover_element_when_hover_is_True(self, part_prop_, part_):
        part_prop_.return_value = part_
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}/a:hlinkHover{r:id=rId5}")
        action_setting = ActionSetting(cNvPr, None, hover=True)

        action_setting._clear_click_action()

        assert part_.drop_rel.call_args_list == [call("rId5")]
        assert action_setting._element.xml == xml("p:cNvPr{a:a=a,r:r=r}")

    def it_returns_None_sound_when_no_hyperlink_is_present(self):
        action_setting = ActionSetting(element("p:cNvPr"), None)
        assert action_setting.sound is None

    def it_returns_None_sound_when_hyperlink_has_no_snd(self):
        action_setting = ActionSetting(element("p:cNvPr/a:hlinkClick{r:id=rId1}"), None)
        assert action_setting.sound is None

    def it_provides_a_Sound_object_when_snd_is_present(self, part_prop_, slide_part_):
        part_prop_.return_value = slide_part_
        cNvPr = element(
            "p:cNvPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId2,name=applause.wav}"
        )
        action_setting = ActionSetting(cNvPr, None)

        sound = action_setting.sound

        assert isinstance(sound, Sound)
        assert sound.rId == "rId2"
        assert sound.name == "applause.wav"

    def it_can_set_a_sound_from_a_file_like(self, part_prop_, slide_part_):
        part_prop_.return_value = slide_part_
        slide_part_.get_or_add_sound_media_part.return_value = "rId9"
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None)

        sound = action_setting.set_sound(io.BytesIO(b"RIFFWAV"), name="boing.wav")

        assert slide_part_.get_or_add_sound_media_part.call_count == 1
        assert sound.rId == "rId9"
        assert sound.name == "boing.wav"
        assert action_setting._element.xml == xml(
            "p:cNvPr{a:a=a,r:r=r}/a:hlinkClick/a:snd{r:embed=rId9,name=boing.wav}"
        )

    def it_can_set_a_sound_on_a_hover_action(self, part_prop_, slide_part_):
        part_prop_.return_value = slide_part_
        slide_part_.get_or_add_sound_media_part.return_value = "rId9"
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None, hover=True)

        action_setting.set_sound(io.BytesIO(b"RIFFWAV"), name="boing.wav")

        assert action_setting._element.xml == xml(
            "p:cNvPr{a:a=a,r:r=r}/a:hlinkHover/a:snd{r:embed=rId9,name=boing.wav}"
        )

    def it_replaces_an_existing_sound_when_setting_a_new_one(
        self, part_prop_, slide_part_
    ):
        part_prop_.return_value = slide_part_
        slide_part_.get_or_add_sound_media_part.return_value = "rId99"
        cNvPr = element(
            "p:cNvPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId5,name=old.wav}"
        )
        action_setting = ActionSetting(cNvPr, None)

        sound = action_setting.set_sound(io.BytesIO(b"RIFFWAV"), name="new.wav")

        # -- the original audio rel is dropped before the new one is attached
        assert slide_part_.drop_rel.call_args_list == [call("rId5")]
        assert sound.rId == "rId99"
        assert action_setting._element.xml == xml(
            "p:cNvPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId99,name=new.wav}"
        )

    def it_defaults_sound_name_to_the_audio_filename(
        self, part_prop_, slide_part_, tmp_path
    ):
        wav_path = tmp_path / "bell.wav"
        wav_path.write_bytes(b"RIFFWAV")
        part_prop_.return_value = slide_part_
        slide_part_.get_or_add_sound_media_part.return_value = "rId7"
        cNvPr = element("p:cNvPr")
        action_setting = ActionSetting(cNvPr, None)

        sound = action_setting.set_sound(str(wav_path))

        assert sound.name == "bell.wav"

    def it_accepts_an_Audio_instance_directly_without_rereading(
        self, part_prop_, slide_part_
    ):
        part_prop_.return_value = slide_part_
        slide_part_.get_or_add_sound_media_part.return_value = "rId8"
        audio = Audio.from_blob(b"RIFFWAV", None, "zap.wav")
        action_setting = ActionSetting(element("p:cNvPr"), None)

        action_setting.set_sound(audio)

        # -- the Audio instance is passed through untouched to the media-part helper
        (args, _) = slide_part_.get_or_add_sound_media_part.call_args
        assert args[0] is audio

    def it_removes_a_sound_and_drops_the_audio_rel(self, part_prop_, slide_part_):
        part_prop_.return_value = slide_part_
        cNvPr = element(
            "p:cNvPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId2,name=applause.wav}"
        )
        action_setting = ActionSetting(cNvPr, None)

        action_setting.remove_sound()

        assert slide_part_.drop_rel.call_args_list == [call("rId2")]
        assert action_setting._element.xml == xml("p:cNvPr/a:hlinkClick{r:id=rId1}")

    def it_silently_ignores_remove_sound_when_none_is_present(
        self, part_prop_, slide_part_
    ):
        part_prop_.return_value = slide_part_
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId1}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.remove_sound()

        slide_part_.drop_rel.assert_not_called()
        assert action_setting._element.xml == xml("p:cNvPr/a:hlinkClick{r:id=rId1}")

    def it_returns_None_screen_tip_when_no_hyperlink_is_present(self):
        action_setting = ActionSetting(element("p:cNvPr"), None)
        assert action_setting.screen_tip is None

    def it_returns_None_screen_tip_when_hyperlink_has_no_tooltip(self):
        action_setting = ActionSetting(element("p:cNvPr/a:hlinkClick{r:id=rId1}"), None)
        assert action_setting.screen_tip is None

    def it_reads_the_tooltip_attribute_when_present(self):
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId1,tooltip=Click me}")
        action_setting = ActionSetting(cNvPr, None)
        assert action_setting.screen_tip == "Click me"

    def it_reads_the_tooltip_from_an_hlinkHover_element(self):
        cNvPr = element("p:cNvPr/a:hlinkHover{r:id=rId1,tooltip=Hover}")
        action_setting = ActionSetting(cNvPr, None, hover=True)
        assert action_setting.screen_tip == "Hover"

    def it_creates_a_hlinkClick_when_setting_a_tip_with_no_hyperlink(self):
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.screen_tip = "Hello"

        assert action_setting._element.xml == xml(
            "p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{tooltip=Hello}"
        )

    def it_adds_the_tooltip_attribute_to_an_existing_hyperlink(self):
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId1}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.screen_tip = "Visit site"

        assert action_setting._element.xml == xml(
            "p:cNvPr/a:hlinkClick{r:id=rId1,tooltip=Visit site}"
        )

    def it_replaces_an_existing_tooltip(self):
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId1,tooltip=old}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.screen_tip = "new"

        assert action_setting._element.xml == xml(
            "p:cNvPr/a:hlinkClick{r:id=rId1,tooltip=new}"
        )

    def it_sets_the_tooltip_on_hlinkHover_when_hover_is_True(self):
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None, hover=True)

        action_setting.screen_tip = "Hovered"

        assert action_setting._element.xml == xml(
            "p:cNvPr{a:a=a,r:r=r}/a:hlinkHover{tooltip=Hovered}"
        )

    def it_removes_the_tooltip_when_assigned_None(self):
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId1,tooltip=Clickme}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.screen_tip = None

        # the hyperlink element itself is preserved (may carry URL/action)
        assert action_setting._element.xml == xml("p:cNvPr/a:hlinkClick{r:id=rId1}")

    def it_removes_the_tooltip_when_assigned_empty_string(self):
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId1,tooltip=Clickme}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.screen_tip = ""

        assert action_setting._element.xml == xml("p:cNvPr/a:hlinkClick{r:id=rId1}")

    def it_silently_ignores_removing_tooltip_when_no_hyperlink_is_present(self):
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.screen_tip = None

        # no hyperlink is created just to clear a non-existent tooltip
        assert action_setting._element.xml == xml("p:cNvPr{a:a=a,r:r=r}")

    # -- macro (ppaction://macro) ----------------------------------

    def it_returns_None_macro_when_no_hyperlink_is_present(self):
        action_setting = ActionSetting(element("p:cNvPr"), None)
        assert action_setting.macro is None

    def it_returns_None_macro_when_hyperlink_has_no_action_verb(self):
        action_setting = ActionSetting(
            element("p:cNvPr/a:hlinkClick{r:id=rId1}"), None
        )
        assert action_setting.macro is None

    def it_returns_None_macro_when_action_verb_is_not_macro(self):
        cNvPr = element("p:cNvPr/a:hlinkClick")
        cNvPr.hlinkClick.action = "ppaction://hlinkshowjump?jump=nextslide"
        action_setting = ActionSetting(cNvPr, None)
        assert action_setting.macro is None

    def it_returns_None_macro_when_macro_action_has_no_name_field(self):
        cNvPr = element("p:cNvPr/a:hlinkClick")
        cNvPr.hlinkClick.action = "ppaction://macro"
        action_setting = ActionSetting(cNvPr, None)
        assert action_setting.macro is None

    def it_reads_the_macro_name_from_the_action_attribute(self):
        cNvPr = element("p:cNvPr/a:hlinkClick")
        cNvPr.hlinkClick.action = "ppaction://macro?name=Module1.MySub"
        action_setting = ActionSetting(cNvPr, None)
        assert action_setting.macro == "Module1.MySub"

    def it_reads_the_macro_name_from_the_hover_hyperlink(self):
        cNvPr = element("p:cNvPr/a:hlinkHover")
        cNvPr.hlinkHover.action = "ppaction://macro?name=HoverMacro"
        action_setting = ActionSetting(cNvPr, None, hover=True)
        assert action_setting.macro == "HoverMacro"

    def it_creates_a_hlinkClick_when_setting_macro_with_no_hyperlink(self):
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.macro = "Module1.Go"

        # -- verify the action attribute directly; cxml attr-value grammar
        # -- does not accept the '?' and '=' characters in ppaction URLs
        hlinkClick = cNvPr.hlinkClick
        assert hlinkClick is not None
        assert hlinkClick.action == "ppaction://macro?name=Module1.Go"
        assert action_setting.action is PP_ACTION.RUN_MACRO
        assert action_setting.macro == "Module1.Go"

    def it_sets_the_macro_name_on_hlinkHover_when_hover_is_True(self):
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None, hover=True)

        action_setting.macro = "HoverSub"

        assert cNvPr.hlinkClick is None
        assert cNvPr.hlinkHover is not None
        assert cNvPr.hlinkHover.action == "ppaction://macro?name=HoverSub"

    def it_replaces_an_existing_macro_with_a_new_name(self, part_prop_, part_):
        part_prop_.return_value = part_
        cNvPr = element("p:cNvPr")
        cNvPr.get_or_add_hlinkClick().action = "ppaction://macro?name=Old"
        action_setting = ActionSetting(cNvPr, None)

        action_setting.macro = "New"

        assert cNvPr.hlinkClick.action == "ppaction://macro?name=New"

    def it_replaces_an_existing_hyperlink_with_a_macro(self, part_prop_, part_):
        part_prop_.return_value = part_
        # -- start with a URL hyperlink carrying an external relationship
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId7}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.macro = "ReplaceMe"

        # -- the URL relationship is dropped, action attribute rewritten
        assert part_.drop_rel.call_args_list == [call("rId7")]
        hlinkClick = cNvPr.hlinkClick
        assert hlinkClick is not None
        assert hlinkClick.action == "ppaction://macro?name=ReplaceMe"
        # -- the original r:id is gone (rel was dropped and hlink replaced)
        assert not hlinkClick.rId

    def it_removes_the_macro_when_assigned_None(self, part_prop_, part_):
        part_prop_.return_value = part_
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        cNvPr.get_or_add_hlinkClick().action = "ppaction://macro?name=Gone"
        action_setting = ActionSetting(cNvPr, None)

        action_setting.macro = None

        # the hlinkClick itself is removed because it carried nothing but the macro
        assert action_setting._element.xml == xml("p:cNvPr{a:a=a,r:r=r}")

    def it_removes_the_macro_when_assigned_empty_string(self, part_prop_, part_):
        part_prop_.return_value = part_
        cNvPr = element("p:cNvPr")
        cNvPr.get_or_add_hlinkClick().action = "ppaction://macro?name=Gone"
        action_setting = ActionSetting(cNvPr, None)

        action_setting.macro = ""

        assert action_setting._element.xml == xml("p:cNvPr")

    def it_leaves_other_actions_alone_when_clearing_macro(self, part_prop_, part_):
        part_prop_.return_value = part_
        # -- the action is a slide-jump, not a macro; assigning None must
        # -- not clobber it
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId9}")
        cNvPr.hlinkClick.action = "ppaction://hlinksldjump"
        action_setting = ActionSetting(cNvPr, None)

        action_setting.macro = None

        part_.drop_rel.assert_not_called()
        # -- the hlinkClick survived intact
        assert cNvPr.hlinkClick is not None
        assert cNvPr.hlinkClick.rId == "rId9"
        assert cNvPr.hlinkClick.action == "ppaction://hlinksldjump"

    def it_silently_ignores_clear_macro_when_no_hyperlink_is_present(self):
        cNvPr = element("p:cNvPr{a:a=a,r:r=r}")
        action_setting = ActionSetting(cNvPr, None)

        action_setting.macro = None

        assert action_setting._element.xml == xml("p:cNvPr{a:a=a,r:r=r}")

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("p:cNvPr", None, PP_ACTION.NONE),
            ("p:cNvPr/a:hlinkClick", None, PP_ACTION.HYPERLINK),
            (
                "p:cNvPr/a:hlinkClick",
                "ppaction://hlinkshowjump?jump=firstslide",
                PP_ACTION.FIRST_SLIDE,
            ),
            (
                "p:cNvPr/a:hlinkClick",
                "ppaction://hlinkshowjump?jump=lastslide",
                PP_ACTION.LAST_SLIDE,
            ),
            (
                "p:cNvPr/a:hlinkClick",
                "ppaction://hlinkshowjump?jump=nextslide",
                PP_ACTION.NEXT_SLIDE,
            ),
            (
                "p:cNvPr/a:hlinkClick",
                "ppaction://hlinkshowjump?jump=previousslide",
                PP_ACTION.PREVIOUS_SLIDE,
            ),
            (
                "p:cNvPr/a:hlinkClick",
                "ppaction://hlinkshowjump?jump=endshow",
                PP_ACTION.END_SHOW,
            ),
            ("p:cNvPr/a:hlinkClick", "ppaction://hlinksldjump", PP_ACTION.NAMED_SLIDE),
            ("p:cNvPr/a:hlinkClick", "ppaction://hlinkfile", PP_ACTION.OPEN_FILE),
            ("p:cNvPr/a:hlinkClick", "ppaction://hlinkpres", PP_ACTION.PLAY),
            (
                "p:cNvPr/a:hlinkClick",
                "ppaction://customshow",
                PP_ACTION.NAMED_SLIDE_SHOW,
            ),
            ("p:cNvPr/a:hlinkClick", "ppaction://ole", PP_ACTION.OLE_VERB),
            ("p:cNvPr/a:hlinkClick", "ppaction://macro", PP_ACTION.RUN_MACRO),
            ("p:cNvPr/a:hlinkClick", "ppaction://program", PP_ACTION.RUN_PROGRAM),
            (
                "p:cNvPr/a:hlinkClick",
                "ppaction://hlinkshowjump?jump=lastslideviewed",
                PP_ACTION.LAST_SLIDE_VIEWED,
            ),
            ("p:cNvPr/a:hlinkClick", "ppaction://media", PP_ACTION.NONE),
        ]
    )
    def action_fixture(self, request):
        cNvPr_cxml, action_text, expected_action = request.param
        cNvPr = element(cNvPr_cxml)
        if action_text is not None:
            cNvPr.hlinkClick.action = action_text
        action_setting = ActionSetting(cNvPr, None)
        return action_setting, expected_action

    @pytest.fixture(
        params=[
            ("p:cNvPr", None, "p:cNvPr"),
            (
                "p:cNvPr{a:b=c,r:s=t}/a:hlinkClick{r:id=rId42}",
                "rId42",
                "p:cNvPr{a:b=c,r:s=t}",
            ),
        ]
    )
    def clear_fixture(self, request, part_prop_, part_):
        xPr_cxml, rId, expected_cxml = request.param
        action_setting = ActionSetting(element(xPr_cxml), None)

        part_prop_.return_value = part_

        calls = [call(rId)] if rId else []
        expected_xml = xml(expected_cxml)
        return action_setting, calls, expected_xml

    @pytest.fixture
    def hyperlink_fixture(self, Hyperlink_, hyperlink_):
        xPr, parent = "xPr", "parent"
        action_setting = ActionSetting(xPr, parent)
        return action_setting, hyperlink_, Hyperlink_, xPr, parent

    @pytest.fixture
    def _slide_index_fixture(self, request, part_prop_):
        action_setting = ActionSetting(None, None)
        slide_part = part_prop_.return_value
        slides = slide_part.package.presentation_part.presentation.slides
        slide = slide_part.slide
        expected_value = 123
        slides.index.return_value = expected_value
        return action_setting, slides, slide, expected_value

    @pytest.fixture(
        params=[
            (PP_ACTION.NONE, None),
            (PP_ACTION.HYPERLINK, None),
            (PP_ACTION.FIRST_SLIDE, 0),
            (PP_ACTION.LAST_SLIDE, 5),
            (PP_ACTION.NEXT_SLIDE, 3),
            (PP_ACTION.PREVIOUS_SLIDE, 1),
            (PP_ACTION.END_SHOW, None),
            (PP_ACTION.NAMED_SLIDE, 4),
            (PP_ACTION.OPEN_FILE, None),
            (PP_ACTION.PLAY, None),
            (PP_ACTION.NAMED_SLIDE_SHOW, None),
            (PP_ACTION.OLE_VERB, None),
            (PP_ACTION.RUN_MACRO, None),
            (PP_ACTION.RUN_PROGRAM, None),
            (PP_ACTION.LAST_SLIDE_VIEWED, None),
        ]
    )
    def target_get_fixture(self, request, action_prop_, _slide_index_prop_, part_prop_):
        action_type, expected_value = request.param
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId6}")
        action_setting = ActionSetting(cNvPr, None)
        action_prop_.return_value = action_type
        _slide_index_prop_.return_value = 2
        # this becomes the return value of ActionSetting._slides
        prs_part_ = part_prop_.return_value.package.presentation_part
        prs_part_.presentation.slides = [0, 1, 2, 3, 4, 5]
        related_part_ = part_prop_.return_value.related_part
        related_part_.return_value.slide = 4
        return action_setting, expected_value

    @pytest.fixture(params=[(PP_ACTION.NEXT_SLIDE, 2), (PP_ACTION.PREVIOUS_SLIDE, 0)])
    def target_raise_fixture(self, request, action_prop_, part_prop_, _slide_index_prop_):
        action_type, slide_idx = request.param
        action_setting = ActionSetting(None, None)
        action_prop_.return_value = action_type
        # this becomes the return value of ActionSetting._slides
        part_prop_.return_value.package.presentation.slides = [0, 1, 2]
        _slide_index_prop_.return_value = slide_idx
        return action_setting

    # fixture components ---------------------------------------------

    @pytest.fixture
    def action_prop_(self, request):
        return property_mock(request, ActionSetting, "action")

    @pytest.fixture
    def _clear_click_action_(self, request):
        return method_mock(request, ActionSetting, "_clear_click_action")

    @pytest.fixture
    def Hyperlink_(self, request, hyperlink_):
        return class_mock(request, "pptx.action.Hyperlink", return_value=hyperlink_)

    @pytest.fixture
    def hyperlink_(self, request):
        return instance_mock(request, Hyperlink)

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, XmlPart)

    @pytest.fixture
    def part_prop_(self, request):
        return property_mock(request, ActionSetting, "part")

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)

    @pytest.fixture
    def _slide_index_prop_(self, request):
        return property_mock(request, ActionSetting, "_slide_index")


class DescribeHyperlink(object):
    """Unit-test suite for `pptx.action.Hyperlink` objects."""

    def it_knows_the_target_url_of_the_hyperlink(self, address_fixture):
        hyperlink, rId, expected_address = address_fixture
        address = hyperlink.address
        hyperlink.part.target_ref.assert_called_once_with(rId)
        assert address == expected_address

    def it_knows_when_theres_no_url(self, no_address_fixture):
        hyperlink = no_address_fixture
        assert hyperlink.address is None

    def it_can_remove_its_url(self, remove_fixture):
        hyperlink, calls, expected_xml = remove_fixture

        hyperlink.address = None

        assert hyperlink.part.drop_rel.call_args_list == calls
        assert hyperlink._element.xml == expected_xml

    def it_can_set_its_target_url(self, update_fixture):
        hyperlink, url, calls, expected_xml = update_fixture

        hyperlink.address = url

        assert hyperlink.part.relate_to.call_args_list == calls
        assert hyperlink._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def address_fixture(self, part_prop_):
        cNvPr_cxml, rId = "p:cNvPr/a:hlinkClick{r:id=rId1}", "rId1"
        expected_address = "http://foobar.com"
        cNvPr = element(cNvPr_cxml)
        hyperlink = Hyperlink(cNvPr, None)
        part_prop_.return_value.target_ref.return_value = expected_address
        return hyperlink, rId, expected_address

    @pytest.fixture(params=["p:cNvPr", "p:cNvPr/a:hlinkClick"])
    def no_address_fixture(self, request):
        cNvPr_cxml = request.param
        cNvPr = element(cNvPr_cxml)
        hyperlink = Hyperlink(cNvPr, None)
        return hyperlink

    @pytest.fixture(
        params=[
            ("p:cNvPr{a:a=a,r:r=r}", []),
            ("p:cNvPr{a:a=a,r:r=r}/a:hlinkClick", []),
            ("p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{r:id=rId3}", [call("rId3")]),
        ]
    )
    def remove_fixture(self, request, part_prop_):
        cNvPr_cxml, calls = request.param
        cNvPr = element(cNvPr_cxml)
        expected_xml = xml("p:cNvPr{a:a=a,r:r=r}")
        hyperlink = Hyperlink(cNvPr, None)
        return hyperlink, calls, expected_xml

    @pytest.fixture(
        params=[
            (
                "p:cNvPr{a:a=a,r:r=r}",
                "http://foo.com",
                False,
                True,
                "p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{r:id=rId3}",
            ),
            (
                "p:cNvPr{a:a=a,r:r=r}",
                "http://bar.com",
                True,
                True,
                "p:cNvPr{a:a=a,r:r=r}/a:hlinkHover{r:id=rId3}",
            ),
            (
                "p:cNvPr/a:hlinkClick{r:id=rId6}",
                "http://baz.com",
                False,
                True,
                "p:cNvPr/a:hlinkClick{r:id=rId3}",
            ),
            (
                "p:cNvPr/a:hlinkHover{r:id=rId6}",
                "http://zab.com",
                True,
                True,
                "p:cNvPr/a:hlinkHover{r:id=rId3}",
            ),
            (
                "p:cNvPr/a:hlinkHover{r:id=rId6}",
                "http://boo.com",
                False,
                True,
                "p:cNvPr/(a:hlinkClick{r:id=rId3},a:hlinkHover{r:id=rId6})",
            ),
            (
                "p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{r:id=rId6}",
                None,
                False,
                False,
                "p:cNvPr{a:a=a,r:r=r}",
            ),
        ]
    )
    def update_fixture(self, request, part_prop_):
        cNvPr_cxml, url, hover, called, expected_cNvPr_cxml = request.param
        cNvPr = element(cNvPr_cxml)
        hyperlink = Hyperlink(cNvPr, None, hover)
        calls = [call(url, RT.HYPERLINK, is_external=True)] if called else []
        part_prop_.return_value.relate_to.return_value = "rId3"
        expected_xml = xml(expected_cNvPr_cxml)
        return hyperlink, url, calls, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def part_prop_(self, request):
        return property_mock(request, Hyperlink, "part")


class DescribeSound(object):
    """Unit-test suite for `pptx.action.Sound` objects."""

    def it_exposes_name_and_rId_of_the_snd_element(self, request):
        snd = element("a:snd{r:embed=rId2,name=boo.wav}")
        part_ = instance_mock(request, SlidePart)
        sound = Sound(snd, part_)

        assert sound.rId == "rId2"
        assert sound.name == "boo.wav"

    def it_defaults_name_to_empty_string_when_absent(self, request):
        snd = element("a:snd{r:embed=rId2}")
        part_ = instance_mock(request, SlidePart)
        sound = Sound(snd, part_)

        assert sound.name == ""

    def it_returns_the_audio_blob_from_the_related_media_part(self, request):
        snd = element("a:snd{r:embed=rId2,name=a.wav}")
        media_part_ = instance_mock(request, SlidePart)
        media_part_.blob = b"RIFFWAVE..."
        part_ = instance_mock(request, SlidePart)
        part_.related_part.return_value = media_part_

        sound = Sound(snd, part_)

        assert sound.blob == b"RIFFWAVE..."
        part_.related_part.assert_called_once_with("rId2")


class Describe_Issue782_SoundAccess(object):
    """Regression suite pinning the issue #782 programmatic-sound-access contract.

    Issue #782 asked for the ability to *access* click/hover-action sounds
    programmatically (read name + blob, attach a new sound, remove one). Wave 2
    #734 shipped that surface (``ActionSetting.sound``, ``.set_sound()``,
    ``.remove_sound()`` plus the ``Sound`` view and ``Audio`` value object).
    This suite round-trips those entry points at the proxy-object level so a
    future refactor cannot silently regress the user-visible API.
    """

    def it_reads_an_existing_snd_via_ActionSetting_sound(self, request):
        # -- read path: `<a:snd>` under `<a:hlinkClick>` surfaces as a Sound --
        part_ = instance_mock(request, SlidePart)
        media_part_ = instance_mock(request, SlidePart)
        media_part_.blob = b"RIFFWAV_payload"
        part_.related_part.return_value = media_part_
        property_mock(request, ActionSetting, "part").return_value = part_
        cNvPr = element(
            "p:cNvPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId2,name=applause.wav}"
        )
        action_setting = ActionSetting(cNvPr, None)

        sound = action_setting.sound

        assert sound is not None
        assert sound.name == "applause.wav"
        assert sound.rId == "rId2"
        assert sound.blob == b"RIFFWAV_payload"

    def it_attaches_a_sound_via_ActionSetting_set_sound(self, request):
        # -- write path: set_sound embeds the blob and writes an `<a:snd>` child --
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.get_or_add_sound_media_part.return_value = "rId42"
        property_mock(request, ActionSetting, "part").return_value = slide_part_
        # -- {a:a=a,r:r=r} forces the `a:` / `r:` namespace declarations onto
        # -- the root `p:cNvPr` so the round-tripped XML comparison below is
        # -- stable regardless of where lxml chooses to place them.
        action_setting = ActionSetting(element("p:cNvPr{a:a=a,r:r=r}"), None)

        audio = Audio.from_blob(b"RIFFWAVexplicit", "audio/wav", "ding.wav")
        sound = action_setting.set_sound(audio)

        # -- Audio is threaded through untouched to the media-part helper --
        (args, _) = slide_part_.get_or_add_sound_media_part.call_args
        assert args[0] is audio
        assert sound.name == "ding.wav"
        assert sound.rId == "rId42"
        assert action_setting._element.xml == xml(
            "p:cNvPr{a:a=a,r:r=r}/a:hlinkClick/a:snd{r:embed=rId42,name=ding.wav}"
        )

    def it_round_trips_read_replace_remove_on_the_same_ActionSetting(self, request):
        # -- composite scenario: read -> replace -> remove leaves the shape
        # -- with an empty `<a:hlinkClick>` (hyperlink element preserved) --
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.get_or_add_sound_media_part.return_value = "rId77"
        property_mock(request, ActionSetting, "part").return_value = slide_part_
        cNvPr = element(
            "p:cNvPr/a:hlinkClick{r:id=rId1}/a:snd{r:embed=rId9,name=old.wav}"
        )
        action_setting = ActionSetting(cNvPr, None)

        # (1) read the original sound -------------------------------------
        original = action_setting.sound
        assert original is not None
        assert original.name == "old.wav"
        assert original.rId == "rId9"

        # (2) replace it with a new sound --------------------------------
        replacement = action_setting.set_sound(
            Audio.from_blob(b"RIFF_new", "audio/wav", "new.wav")
        )
        assert replacement.rId == "rId77"
        # -- the prior audio relationship is dropped as part of replace --
        assert call("rId9") in slide_part_.drop_rel.call_args_list

        # (3) remove the sound completely --------------------------------
        action_setting.remove_sound()
        assert action_setting.sound is None
        # -- hlinkClick itself is retained so any URL/action survives --
        assert action_setting._element.xml == xml(
            "p:cNvPr/a:hlinkClick{r:id=rId1}"
        )

    def it_exposes_sound_on_a_hover_action_through_the_same_surface(self, request):
        # -- the same property reads from `<a:hlinkHover>` when hover=True --
        part_ = instance_mock(request, SlidePart)
        property_mock(request, ActionSetting, "part").return_value = part_
        cNvPr = element(
            "p:cNvPr/a:hlinkHover{r:id=rId1}/a:snd{r:embed=rId3,name=hover.wav}"
        )
        action_setting = ActionSetting(cNvPr, None, hover=True)

        sound = action_setting.sound

        assert sound is not None
        assert sound.name == "hover.wav"
        assert sound.rId == "rId3"

    def it_returns_None_for_an_action_that_has_no_snd(self):
        # -- and, critically, access is safe when no sound is present --
        no_hlink = ActionSetting(element("p:cNvPr"), None)
        hlink_no_snd = ActionSetting(
            element("p:cNvPr/a:hlinkClick{r:id=rId1}"), None
        )

        assert no_hlink.sound is None
        assert hlink_no_snd.sound is None
