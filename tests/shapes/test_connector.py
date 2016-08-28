# encoding: utf-8

"""
Unit test suite for pptx.shapes.connector module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.shapes.connector import Connector
from pptx.util import Emu

from ..unitutil.cxml import element


class DescribeConnector(object):

    def it_knows_its_begin_point_x_location(self, begin_x_get_fixture):
        connector, expected_value = begin_x_get_fixture
        begin_x = connector.begin_x
        assert isinstance(begin_x, Emu)
        assert connector.begin_x == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (42, 24, False, 42),
        (24, 42, True,  66),
    ])
    def begin_x_get_fixture(self, request):
        x, cx, flipH, expected_value = request.param
        cxnSp = element(
            'p:cxnSp/p:spPr/a:xfrm{flipH=%d}/(a:off{x=%d,y=6},a:ext{cx=%d,cy'
            '=32})' % (flipH, x, cx)
        )
        connector = Connector(cxnSp, None)
        return connector, expected_value
