# encoding: utf-8

"""Test suite for pptx.shapes.group module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.shapes.group import GroupShape


class DescribeGroupShape(object):

    def it_raises_on_access_click_action(self, click_action_fixture):
        group = click_action_fixture
        with pytest.raises(TypeError):
            group.click_action

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def click_action_fixture(self):
        return GroupShape(None, None)
