# encoding: utf-8

"""
Test suite for pptx.parts.slidemaster module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.oxml.parts.slidemaster import CT_SlideLayoutIdList
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.parts.slidelayout import SlideLayout
from pptx.parts.slidemaster import (
    _MasterPlaceholders, _MasterShapeFactory, _MasterShapeTree,
    _SlideLayouts, SlideMaster
)
from pptx.shapes.base import BaseShape
from pptx.shapes.placeholder import MasterPlaceholder

from ..oxml.unitdata.shape import a_ph, a_pic, an_nvPr, an_nvSpPr, an_sp
from ..oxml.unitdata.slides import (
    a_sldLayoutId, a_sldLayoutIdLst, a_sldMaster
)
from ..unitutil.mock import (
    class_mock, function_mock, instance_mock, method_mock, property_mock
)


class DescribeSlideMaster(object):

    def it_provides_access_to_its_placeholders(self, slide_master):
        placeholders = slide_master.placeholders
        assert isinstance(placeholders, _MasterPlaceholders)
        assert placeholders._slide is slide_master

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


class Describe_MasterShapeFactory(object):

    def it_constructs_a_master_placeholder_for_a_shape_element(
            self, factory_fixture):
        shape_elm, parent_, ShapeConstructor_, shape_ = factory_fixture
        shape = _MasterShapeFactory(shape_elm, parent_)
        ShapeConstructor_.assert_called_once_with(shape_elm, parent_)
        assert shape is shape_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=['ph', 'sp', 'pic'])
    def factory_fixture(
            self, request, ph_bldr, slide_master_, _MasterPlaceholder_,
            master_placeholder_, BaseShapeFactory_, base_shape_):
        shape_bldr, ShapeConstructor_, shape_mock = {
            'ph':  (ph_bldr, _MasterPlaceholder_, master_placeholder_),
            'sp':  (an_sp(), BaseShapeFactory_,   base_shape_),
            'pic': (a_pic(), BaseShapeFactory_,   base_shape_),
        }[request.param]
        shape_elm = shape_bldr.with_nsdecls().element
        return shape_elm, slide_master_, ShapeConstructor_, shape_mock

    # fixture components -----------------------------------

    @pytest.fixture
    def BaseShapeFactory_(self, request, base_shape_):
        return function_mock(
            request, 'pptx.parts.slidemaster.BaseShapeFactory',
            return_value=base_shape_
        )

    @pytest.fixture
    def base_shape_(self, request):
        return instance_mock(request, BaseShape)

    @pytest.fixture
    def _MasterPlaceholder_(self, request, master_placeholder_):
        return class_mock(
            request, 'pptx.parts.slidemaster.MasterPlaceholder',
            return_value=master_placeholder_
        )

    @pytest.fixture
    def master_placeholder_(self, request):
        return instance_mock(request, MasterPlaceholder)

    @pytest.fixture
    def ph_bldr(self):
        return (
            an_sp().with_child(
                an_nvSpPr().with_child(
                    an_nvPr().with_child(
                        a_ph().with_idx(1))))
        )

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)


class Describe_MasterShapeTree(object):

    def it_constructs_a_master_placeholder_for_a_placeholder_element(
            self, factory_fixture):
        master_shapes, ph_elm_, _MasterShapeFactory_, master_placeholder_ = (
            factory_fixture
        )
        master_placeholder = master_shapes._shape_factory(ph_elm_)
        _MasterShapeFactory_.assert_called_once_with(ph_elm_, master_shapes)
        assert master_placeholder is master_placeholder_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def factory_fixture(
            self, ph_elm_, _MasterShapeFactory_, master_placeholder_):
        master_shapes = _MasterShapeTree(None)
        return (
            master_shapes, ph_elm_, _MasterShapeFactory_, master_placeholder_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def master_placeholder_(self, request):
        return instance_mock(request, MasterPlaceholder)

    @pytest.fixture
    def _MasterShapeFactory_(self, request, master_placeholder_):
        return function_mock(
            request, 'pptx.parts.slidemaster._MasterShapeFactory',
            return_value=master_placeholder_
        )

    @pytest.fixture
    def ph_elm_(self, request):
        return instance_mock(request, CT_Shape)


class Describe_MasterPlaceholders(object):

    def it_uses_master_shape_factory_to_construct_placeholder_shapes(
            self, factory_fixture):
        master_placeholders, shape_elm_ = factory_fixture[:2]
        _MasterShapeFactory_, placeholder_ = factory_fixture[2:]
        placeholder = master_placeholders._shape_factory(shape_elm_)
        _MasterShapeFactory_.assert_called_once_with(
            shape_elm_, master_placeholders
        )
        assert placeholder is placeholder_

    def it_can_find_a_placeholder_by_type(self, get_fixture):
        master_placeholders, ph_type, placeholder_ = get_fixture
        placeholder = master_placeholders.get(ph_type)
        assert placeholder is placeholder_

    def it_returns_default_if_placeholder_of_type_not_found(
            self, default_fixture):
        master_placeholders = default_fixture
        default = 'barfoo'
        placeholder = master_placeholders.get('foobar', default)
        assert placeholder is default

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def default_fixture(self, _iter_):
        master_placeholders = _MasterPlaceholders(None)
        return master_placeholders

    @pytest.fixture
    def factory_fixture(
            self, ph_elm_, _MasterShapeFactory_, placeholder_):
        master_placeholders = _MasterPlaceholders(None)
        return (
            master_placeholders, ph_elm_, _MasterShapeFactory_, placeholder_
        )

    @pytest.fixture(params=['title', 'body'])
    def get_fixture(self, request, _iter_, placeholder_, placeholder_2_):
        master_placeholders = _MasterPlaceholders(None)
        ph_type = request.param
        ph_shape_ = {
            'title': placeholder_, 'body': placeholder_2_
        }[request.param]
        return master_placeholders, ph_type, ph_shape_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _iter_(self, request, placeholder_, placeholder_2_):
        return method_mock(
            request, _MasterPlaceholders, '__iter__',
            return_value=iter([placeholder_, placeholder_2_])
        )

    @pytest.fixture
    def _MasterShapeFactory_(self, request, placeholder_):
        return function_mock(
            request, 'pptx.parts.slidemaster._MasterShapeFactory',
            return_value=placeholder_
        )

    @pytest.fixture
    def ph_elm_(self, request):
        return instance_mock(request, CT_Shape)

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, MasterPlaceholder, ph_type='title')

    @pytest.fixture
    def placeholder_2_(self, request):
        return instance_mock(request, MasterPlaceholder, ph_type='body')
