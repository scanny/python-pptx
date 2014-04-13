# encoding: utf-8

"""
Test suite for pptx.dml.line module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat
from pptx.oxml.shapes.shared import CT_LineProperties
from pptx.shapes.autoshape import Shape

from ..unitutil import class_mock, instance_mock


class DescribeLineFormat(object):

    def it_has_a_fill(self, fill_fixture):
        line, FillFormat_, ln_, fill_ = fill_fixture
        fill = line.fill
        FillFormat_.from_fill_parent.assert_called_once_with(ln_)
        assert fill is fill_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def fill_fixture(self, line, FillFormat_, ln_, fill_):
        return line, FillFormat_, ln_, fill_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def fill_(self, request):
        return instance_mock(request, FillFormat)

    @pytest.fixture
    def FillFormat_(self, request, fill_):
        FillFormat_ = class_mock(request, 'pptx.dml.line.FillFormat')
        FillFormat_.from_fill_parent.return_value = fill_
        return FillFormat_

    @pytest.fixture
    def line(self, shape_):
        return LineFormat(shape_)

    @pytest.fixture
    def ln_(self, request):
        return instance_mock(request, CT_LineProperties)

    @pytest.fixture
    def shape_(self, request, ln_):
        shape_ = instance_mock(request, Shape)
        shape_.get_or_add_ln.return_value = ln_
        return shape_
