# encoding: utf-8

"""Test suite for pptx.shapes.group module."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.group import GroupShape


class DescribeGroupShape(object):

    def it_raises_on_access_click_action(self, click_action_fixture):
        group = click_action_fixture
        with pytest.raises(TypeError):
            group.click_action

    def it_knows_its_shape_type(self, shape_type_fixture):
        group = shape_type_fixture
        assert group.shape_type == MSO_SHAPE_TYPE.GROUP

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def click_action_fixture(self):
        return GroupShape(None, None)

    @pytest.fixture
    def shape_type_fixture(self):
        return GroupShape(None, None)
