# encoding: utf-8

"""
Test suite for pptx.action module
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.action import ActionSetting
from pptx.enum.action import PP_ACTION

from .unitutil.cxml import element


class DescribeActionSetting(object):

    def it_knows_its_action_type(self, action_fixture):
        action_setting, expected_action = action_fixture
        action = action_setting.action
        assert action is expected_action

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('p:cNvPr',              None,
         PP_ACTION.NONE),
        ('p:cNvPr/a:hlinkClick', None,
         PP_ACTION.HYPERLINK),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinkshowjump?jump=firstslide',
         PP_ACTION.FIRST_SLIDE),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinkshowjump?jump=lastslide',
         PP_ACTION.LAST_SLIDE),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinkshowjump?jump=nextslide',
         PP_ACTION.NEXT_SLIDE),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinkshowjump?jump=previousslide',
         PP_ACTION.PREVIOUS_SLIDE),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinkshowjump?jump=endshow',
         PP_ACTION.END_SHOW),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinksldjump',
         PP_ACTION.NAMED_SLIDE),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinkfile',
         PP_ACTION.OPEN_FILE),
        ('p:cNvPr/a:hlinkClick', 'ppaction://hlinkpres',
         PP_ACTION.PLAY),
        ('p:cNvPr/a:hlinkClick', 'ppaction://customshow',
         PP_ACTION.NAMED_SLIDE_SHOW),
        ('p:cNvPr/a:hlinkClick', 'ppaction://ole',
         PP_ACTION.OLE_VERB),
        ('p:cNvPr/a:hlinkClick', 'ppaction://macro',
         PP_ACTION.RUN_MACRO),
        ('p:cNvPr/a:hlinkClick', 'ppaction://program',
         PP_ACTION.RUN_PROGRAM),
        ('p:cNvPr/a:hlinkClick',
         'ppaction://hlinkshowjump?jump=lastslideviewed',
         PP_ACTION.LAST_SLIDE_VIEWED),
    ])
    def action_fixture(self, request):
        cNvPr_cxml, action_text, expected_action = request.param
        cNvPr = element(cNvPr_cxml)
        if action_text is not None:
            cNvPr.hlinkClick.action = action_text
        action_setting = ActionSetting(cNvPr, None)
        return action_setting, expected_action
