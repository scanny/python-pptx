# encoding: utf-8

"""
Test suite for pptx.parts.slidemaster module
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.parts.slidemaster import (
    _MasterPlaceholders, _MasterShapeFactory, _MasterShapeTree,
    SlideMasterPart
)
from pptx.shapes.base import BaseShape
from pptx.shapes.placeholder import MasterPlaceholder
from pptx.slide import SlideLayout, SlideLayouts

from ..oxml.unitdata.shape import a_ph, a_pic, an_nvPr, an_nvSpPr, an_sp
from ..unitutil.cxml import element
from ..unitutil.mock import (
    class_mock, function_mock, instance_mock, method_mock, property_mock
)


class DescribeSlideMasterPart(object):

    def it_provides_access_to_its_placeholders(self, placeholders_fixture):
        slide_master, _MasterPlaceholders_, spTree, placeholders_ = (
            placeholders_fixture
        )
        placeholders = slide_master.placeholders
        _MasterPlaceholders_.assert_called_once_with(spTree, slide_master)
        assert placeholders is placeholders_

    def it_provides_access_to_its_shapes(self, shapes_fixture):
        slide_master, _MasterShapeTree_, spTree, shapes_ = shapes_fixture
        shapes = slide_master.shapes
        _MasterShapeTree_.assert_called_once_with(spTree, slide_master)
        assert shapes is shapes_

    def it_provides_access_to_its_slide_layouts(self, layouts_fixture):
        slide_master, SlideLayouts_, sldLayoutIdLst, slide_layouts_ = (
            layouts_fixture
        )
        slide_layouts = slide_master.slide_layouts
        SlideLayouts_.assert_called_once_with(sldLayoutIdLst, slide_master)
        assert slide_layouts is slide_layouts_

    def it_provides_access_to_a_related_slide_layout(self, related_fixture):
        slide_master_part, rId, getitem_, slide_layout_ = related_fixture
        slide_layout = slide_master_part.related_slide_layout(rId)
        getitem_.assert_called_once_with(rId)
        assert slide_layout is slide_layout_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def layouts_fixture(self, SlideLayouts_, slide_layouts_):
        sldMaster = element('p:sldMaster/p:sldLayoutIdLst')
        slide_master = SlideMasterPart(None, None, sldMaster)
        sldMasterIdLst = sldMaster.sldLayoutIdLst
        return slide_master, SlideLayouts_, sldMasterIdLst, slide_layouts_

    @pytest.fixture
    def related_fixture(self, slide_layout_, related_parts_prop_):
        slide_master_part = SlideMasterPart(None, None, None, None)
        rId = 'rId42'
        getitem_ = related_parts_prop_.return_value.__getitem__
        getitem_.return_value.slide_layout = slide_layout_
        return slide_master_part, rId, getitem_, slide_layout_

    @pytest.fixture
    def placeholders_fixture(self, _MasterPlaceholders_, placeholders_):
        sldMaster = element('p:sldMaster/p:cSld/p:spTree')
        slide_master = SlideMasterPart(None, None, sldMaster)
        spTree = sldMaster.xpath('//p:spTree')[0]
        return slide_master, _MasterPlaceholders_, spTree, placeholders_

    @pytest.fixture
    def shapes_fixture(self, _MasterShapeTree_, shapes_):
        sldMaster = element('p:sldMaster/p:cSld/p:spTree')
        slide_master = SlideMasterPart(None, None, sldMaster)
        spTree = sldMaster.xpath('//p:spTree')[0]
        return slide_master, _MasterShapeTree_, spTree, shapes_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _MasterPlaceholders_(self, request, placeholders_):
        return class_mock(
            request, 'pptx.parts.slidemaster._MasterPlaceholders',
            return_value=placeholders_
        )

    @pytest.fixture
    def _MasterShapeTree_(self, request, shapes_):
        return class_mock(
            request, 'pptx.parts.slidemaster._MasterShapeTree',
            return_value=shapes_
        )

    @pytest.fixture
    def placeholders_(self, request):
        return instance_mock(request, _MasterPlaceholders)

    @pytest.fixture
    def related_parts_prop_(self, request):
        return property_mock(request, SlideMasterPart, 'related_parts')

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, _MasterShapeTree)

    @pytest.fixture
    def SlideLayouts_(self, request, slide_layouts_):
        return class_mock(
            request, 'pptx.parts.slidemaster.SlideLayouts',
            return_value=slide_layouts_
        )

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_layouts_(self, request):
        return instance_mock(request, SlideLayouts)


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
        return instance_mock(request, SlideMasterPart)


class Describe_MasterShapeTree(object):

    def it_provides_access_to_its_shape_factory(self, factory_fixture):
        shapes, sp, _MasterShapeFactory_, placeholder_ = factory_fixture
        placeholder = shapes._shape_factory(sp)
        _MasterShapeFactory_.assert_called_once_with(sp, shapes)
        assert placeholder is placeholder_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def factory_fixture(self, _MasterShapeFactory_, placeholder_):
        shapes = _MasterShapeTree(None, None)
        sp = element('p:sp')
        return shapes, sp, _MasterShapeFactory_, placeholder_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _MasterShapeFactory_(self, request, placeholder_):
        return function_mock(
            request, 'pptx.parts.slidemaster._MasterShapeFactory',
            return_value=placeholder_, autospec=True
        )

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, MasterPlaceholder)


class Describe_MasterPlaceholders(object):

    def it_provides_access_to_its_shape_factory(self, factory_fixture):
        placeholders, sp, _MasterShapeFactory_, placeholder_ = factory_fixture
        placeholder = placeholders._shape_factory(sp)
        _MasterShapeFactory_.assert_called_once_with(sp, placeholders)
        assert placeholder is placeholder_

    def it_can_find_a_placeholder_by_type(self, get_fixture):
        placeholders, ph_type, placeholder_ = get_fixture
        assert placeholders.get(ph_type) is placeholder_

    def it_returns_default_on_ph_type_not_found(self, default_fixture):
        placeholders, default = default_fixture
        assert placeholders.get(42, default) is default

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def default_fixture(self, _iter_):
        placeholders = _MasterPlaceholders(None, None)
        default = 'barfoo'
        return placeholders, default

    @pytest.fixture
    def factory_fixture(self, _MasterShapeFactory_, placeholder_):
        placeholders = _MasterPlaceholders(None, None)
        sp = element('p:sp')
        return placeholders, sp, _MasterShapeFactory_, placeholder_

    @pytest.fixture(params=['title', 'body'])
    def get_fixture(self, request, _iter_, placeholder_, placeholder_2_):
        ph_type = request.param
        placeholders = _MasterPlaceholders(None, None)
        _placeholder_ = {
            'title': placeholder_, 'body': placeholder_2_
        }[ph_type]
        return placeholders, ph_type, _placeholder_

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
            return_value=placeholder_, autospec=True
        )

    @pytest.fixture
    def placeholder_(self, request):
        return instance_mock(request, MasterPlaceholder, ph_type='title')

    @pytest.fixture
    def placeholder_2_(self, request):
        return instance_mock(request, MasterPlaceholder, ph_type='body')
