# encoding: utf-8

"""
Test suite for pptx.shapes.placeholder module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.oxml.ns import _nsmap as nsmap
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.shapes.placeholder import (
    BasePlaceholder, BasePlaceholders, Placeholder
)
from pptx.shapes.shapetree import ShapeCollection
from pptx.spec import (
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL, PH_ORIENT_HORZ,
    PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER
)

from ..oxml.unitdata.shape import (
    a_graphicFrame, a_ph, an_nvGraphicFramePr, an_nvPicPr, an_nvPr,
    an_nvSpPr, an_sp
)
from ..unitutil import absjoin, instance_mock, parse_xml_file, test_file_dir


class DescribePlaceholder(object):

    def it_knows_the_placeholder_type(self, type_fixture):
        placeholder, expected_type = type_fixture
        assert placeholder.type == expected_type

    def it_knows_the_placeholder_orientation(self, orient_fixture):
        placeholder, expected_orient = orient_fixture
        assert placeholder.orient == expected_orient

    def it_knows_the_placeholder_size_name(self, sz_fixture):
        placeholder, expected_sz = sz_fixture
        assert placeholder.sz == expected_sz

    def it_knows_the_placeholder_idx(self, idx_fixture):
        placeholder, expected_idx = idx_fixture
        assert placeholder.idx == expected_idx

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[0, 1, 2, 3, 4, 5])
    def idx_fixture(self, request, layout_shapes):
        sp_idx = request.param
        shape = layout_shapes[sp_idx]
        expected_idx = (0, 10, 1, 14, 12, 11)[sp_idx]
        placeholder = Placeholder(shape)
        return placeholder, expected_idx

    @pytest.fixture(scope="module")
    def layout_shapes(self):
        path = absjoin(test_file_dir, 'slideLayout1.xml')
        sldLayout = parse_xml_file(path).getroot()
        spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        shapes = ShapeCollection(spTree)
        return shapes

    @pytest.fixture(params=[0, 1, 2, 3, 4, 5])
    def orient_fixture(self, request, layout_shapes):
        sp_idx = request.param
        shape = layout_shapes[sp_idx]
        expected_orient = (
            PH_ORIENT_HORZ, PH_ORIENT_HORZ, PH_ORIENT_VERT,
            PH_ORIENT_HORZ, PH_ORIENT_HORZ, PH_ORIENT_HORZ,
        )[sp_idx]
        placeholder = Placeholder(shape)
        return placeholder, expected_orient

    @pytest.fixture(params=[0, 1, 2, 3, 4, 5])
    def sz_fixture(self, request, layout_shapes):
        sp_idx = request.param
        shape = layout_shapes[sp_idx]
        expected_sz = (
            PH_SZ_FULL,    PH_SZ_HALF,    PH_SZ_FULL,
            PH_SZ_QUARTER, PH_SZ_QUARTER, PH_SZ_QUARTER,
        )[sp_idx]
        placeholder = Placeholder(shape)
        return placeholder, expected_sz

    @pytest.fixture(params=[0, 1, 2, 3, 4, 5])
    def type_fixture(self, request, layout_shapes):
        sp_idx = request.param
        shape = layout_shapes[sp_idx]
        expected_type = (
            PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_SUBTITLE, PH_TYPE_TBL,
            PH_TYPE_SLDNUM, PH_TYPE_FTR
        )[sp_idx]
        placeholder = Placeholder(shape)
        return placeholder, expected_type


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

    def it_knows_its_placeholder_type(self, type_fixture):
        placeholder, ph_type = type_fixture
        assert placeholder.ph_type == ph_type

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('sp',           None,    'obj'),
        ('sp',           'title', 'title'),
        ('pic',          'pic',   'pic'),
        ('graphicFrame', 'tbl',   'tbl'),
    ])
    def type_fixture(self, request):
        tagname, ph_type, expected_ph_type = request.param
        shape_elm = self.shape_elm_factory(tagname, ph_type)
        print(shape_elm.xml)
        placeholder = BasePlaceholder(shape_elm, None)
        return placeholder, expected_ph_type

    # fixture components ---------------------------------------------

    @staticmethod
    def shape_elm_factory(tagname, ph_type):
        root_bldr, nvXxPr_bldr = {
            'sp':           (an_sp().with_nsdecls('p'), an_nvSpPr()),
            'pic':          (an_sp().with_nsdecls('p'), an_nvPicPr()),
            'graphicFrame': (a_graphicFrame().with_nsdecls('p'),
                             an_nvGraphicFramePr()),
        }[tagname]
        ph_bldr = {
            None:    a_ph().with_idx(1),
            'body':  a_ph().with_idx(1),
            'title': a_ph().with_type('title'),
            'pic':   a_ph().with_idx(6).with_type('pic'),
            'tbl':   a_ph().with_idx(7).with_type('tbl'),
        }[ph_type]
        return (
            root_bldr.with_child(
                nvXxPr_bldr.with_child(
                    an_nvPr().with_child(
                        ph_bldr)))
        ).element
