# encoding: utf-8

"""
Test suite for pptx.shapes.placeholder module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.shapes.shared import (
    BaseShapeElement, ST_Direction, ST_PlaceholderSize
)
from pptx.shapes.placeholder import BasePlaceholder, BasePlaceholders

from ..oxml.unitdata.shape import (
    a_graphicFrame, a_ph, an_nvGraphicFramePr, an_nvPicPr, an_nvPr,
    an_nvSpPr, an_sp
)
from ..unitutil.mock import instance_mock


class DescribeBasePlaceholders(object):

    def it_contains_only_placeholder_shapes(self, member_fixture):
        shape_elm_, is_ph_shape = member_fixture
        _is_ph_shape = BasePlaceholders._is_member_elm(shape_elm_)
        assert _is_ph_shape == is_ph_shape

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[True, False])
    def member_fixture(self, request, shape_elm_):
        is_ph_shape = request.param
        shape_elm_.has_ph_elm = is_ph_shape
        return shape_elm_, is_ph_shape

    # fixture components ---------------------------------------------

    @pytest.fixture
    def shape_elm_(self, request):
        return instance_mock(request, BaseShapeElement)


class DescribeBasePlaceholder(object):

    def it_knows_its_idx_value(self, idx_fixture):
        placeholder, idx = idx_fixture
        assert placeholder.idx == idx

    def it_knows_its_orient_value(self, orient_fixture):
        placeholder, orient = orient_fixture
        assert placeholder.orient == orient

    def it_raises_on_ph_orient_when_not_a_placeholder(
            self, orient_raise_fixture):
        shape = orient_raise_fixture
        with pytest.raises(ValueError):
            shape.orient

    def it_knows_its_sz_value(self, sz_fixture):
        placeholder, sz = sz_fixture
        assert placeholder.sz == sz

    def it_knows_its_placeholder_type(self, type_fixture):
        placeholder, ph_type = type_fixture
        assert placeholder.ph_type == ph_type

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('sp',           None, 'title', 0),
        ('sp',           0,    'title', 0),
        ('sp',           3,    'body',  3),
        ('pic',          6,    'pic',   6),
        ('graphicFrame', 9,    'tbl',   9),
    ])
    def idx_fixture(self, request):
        tagname, idx, ph_type, expected_idx = request.param
        shape_elm = self.shape_elm_factory(tagname, ph_type, idx)
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_idx

    @pytest.fixture(params=[
        (None, ST_Direction.HORZ),
        (ST_Direction.VERT, ST_Direction.VERT),
    ])
    def orient_fixture(self, request):
        orient, expected_orient = request.param
        ph_bldr = a_ph()
        if orient is not None:
            ph_bldr.with_orient(orient)
        shape_elm = (
            an_sp().with_nsdecls('p').with_child(
                an_nvSpPr().with_child(
                    an_nvPr().with_child(
                        ph_bldr)))
        ).element
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_orient

    @pytest.fixture
    def orient_raise_fixture(self):
        shape_elm = (
            an_sp().with_nsdecls('p').with_child(
                an_nvSpPr().with_child(
                    an_nvPr()))
        ).element
        shape = BasePlaceholder(shape_elm, None)
        return shape

    @pytest.fixture(params=[
        (None, ST_PlaceholderSize.FULL),
        (ST_PlaceholderSize.HALF, ST_PlaceholderSize.HALF),
    ])
    def sz_fixture(self, request):
        sz, expected_sz = request.param
        ph_bldr = a_ph()
        if sz is not None:
            ph_bldr.with_sz(sz)
        shape_elm = (
            an_sp().with_nsdecls('p').with_child(
                an_nvSpPr().with_child(
                    an_nvPr().with_child(
                        ph_bldr)))
        ).element
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_sz

    @pytest.fixture(params=[
        ('sp',           None,    1, PP_PLACEHOLDER.OBJECT),
        ('sp',           'title', 0, PP_PLACEHOLDER.TITLE),
        ('pic',          'pic',   6, PP_PLACEHOLDER.PICTURE),
        ('graphicFrame', 'tbl',   9, PP_PLACEHOLDER.TABLE),
    ])
    def type_fixture(self, request):
        tagname, ph_type, idx, expected_ph_type = request.param
        shape_elm = self.shape_elm_factory(tagname, ph_type, idx)
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_ph_type

    # fixture components ---------------------------------------------

    @staticmethod
    def shape_elm_factory(tagname, ph_type, idx):
        root_bldr, nvXxPr_bldr = {
            'sp':           (an_sp().with_nsdecls('p'), an_nvSpPr()),
            'pic':          (an_sp().with_nsdecls('p'), an_nvPicPr()),
            'graphicFrame': (a_graphicFrame().with_nsdecls('p'),
                             an_nvGraphicFramePr()),
        }[tagname]
        ph_bldr = {
            None:    a_ph().with_idx(idx),
            'obj':   a_ph().with_idx(idx),
            'title': a_ph().with_type('title'),
            'body':  a_ph().with_type('body').with_idx(idx),
            'pic':   a_ph().with_type('pic').with_idx(idx),
            'tbl':   a_ph().with_type('tbl').with_idx(idx),
        }[ph_type]
        return (
            root_bldr.with_child(
                nvXxPr_bldr.with_child(
                    an_nvPr().with_child(
                        ph_bldr)))
        ).element
