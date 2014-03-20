# encoding: utf-8

"""
Test suite for pptx.parts.slidemaster module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.parts.slidelayout import SlideLayout
from pptx.parts.slidemaster import (
    _MasterShapeTree, _SlideLayouts, SlideMaster
)
from pptx.oxml.slidemaster import CT_SlideLayoutIdList

from ..oxml.unitdata.slides import (
    a_sldLayoutId, a_sldLayoutIdLst, a_sldMaster
)
from ..unitutil import instance_mock, method_mock, property_mock


class DescribeSlideMaster(object):

    def it_provides_access_to_its_shapes(self, slide_master):
        shapes = slide_master.shapes
        assert isinstance(shapes, _MasterShapeTree)
        assert shapes._slide is slide_master

    def it_provides_access_to_its_slide_layouts(self, slide_master):
        slide_layouts = slide_master.slide_layouts
        assert isinstance(slide_layouts, _SlideLayouts)

    def it_provides_access_to_its_sldLayoutIdLst(self, slide_master):
        sldLayoutIdLst = slide_master.sldLayoutIdLst
        assert isinstance(sldLayoutIdLst, CT_SlideLayoutIdList)

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[True, False])
    def sldMaster(self, request):
        has_sldLayoutIdLst = request.param
        sldMaster_bldr = a_sldMaster().with_nsdecls()
        if has_sldLayoutIdLst:
            sldMaster_bldr.with_child(a_sldLayoutIdLst())
        return sldMaster_bldr.element

    @pytest.fixture
    def slide_master(self, sldMaster):
        return SlideMaster(None, None, sldMaster, None)


class DescribeSlideLayouts(object):

    def it_knows_how_many_layouts_it_contains(self, len_fixture):
        slide_layouts, expected_count = len_fixture
        slide_layout_count = len(slide_layouts)
        assert slide_layout_count == expected_count

    def it_can_iterate_over_the_slide_layouts(self, iter_fixture):
        slide_layouts, slide_layout_, slide_layout_2_ = iter_fixture
        assert [s for s in slide_layouts] == [slide_layout_, slide_layout_2_]

    def it_iterates_over_rIds_to_help__iter__(self, iter_rIds_fixture):
        slide_layouts, expected_rIds = iter_rIds_fixture
        assert [rId for rId in slide_layouts._iter_rIds()] == expected_rIds

    def it_supports_indexed_access(self, getitem_fixture):
        slide_layouts, idx, slide_layout_ = getitem_fixture
        slide_layout = slide_layouts[idx]
        assert slide_layout is slide_layout_

    def it_raises_on_slide_layout_index_out_of_range(self, getitem_fixture):
        slide_layouts = getitem_fixture[0]
        with pytest.raises(IndexError):
            slide_layouts[2]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getitem_fixture(self, slide_master, related_parts_, slide_layout_):
        slide_layouts = _SlideLayouts(slide_master)
        related_parts_.return_value = {'rId1': None, 'rId2': slide_layout_}
        idx = 1
        return slide_layouts, idx, slide_layout_

    @pytest.fixture
    def iter_fixture(
            self, slide_master_, _iter_rIds_, slide_layout_, slide_layout_2_):
        slide_master_.related_parts = {
            'rId1': slide_layout_, 'rId2': slide_layout_2_
        }
        slide_layouts = _SlideLayouts(slide_master_)
        return slide_layouts, slide_layout_, slide_layout_2_

    @pytest.fixture
    def iter_rIds_fixture(self, slide_master):
        slide_layouts = _SlideLayouts(slide_master)
        expected_rIds = ['rId1', 'rId2']
        return slide_layouts, expected_rIds

    @pytest.fixture
    def len_fixture(self, slide_master_):
        slide_layouts = _SlideLayouts(slide_master_)
        slide_master_.sldLayoutIdLst = [1, 2]
        expected_count = 2
        return slide_layouts, expected_count

    # fixture components -----------------------------------

    @pytest.fixture
    def _iter_rIds_(self, request):
        return method_mock(
            request, _SlideLayouts, '_iter_rIds',
            return_value=iter(['rId1', 'rId2'])
        )

    @pytest.fixture
    def related_parts_(self, request):
        return property_mock(request, SlideMaster, 'related_parts')

    @pytest.fixture
    def sldMaster(self, request):
        sldMaster_bldr = (
            a_sldMaster().with_nsdecls('p', 'r').with_child(
                a_sldLayoutIdLst().with_child(
                    a_sldLayoutId().with_rId('rId1')).with_child(
                    a_sldLayoutId().with_rId('rId2'))
            )
        )
        return sldMaster_bldr.element

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_layout_2_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_master(self, sldMaster):
        return SlideMaster(None, None, sldMaster, None)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)
