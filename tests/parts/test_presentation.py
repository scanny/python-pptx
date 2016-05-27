# encoding: utf-8

"""
Test suite for pptx.parts.presentation module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.package import Package
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import SlidePart
from pptx.parts.slidelayout import SlideLayoutPart
from pptx.presentation import Presentation
from pptx.slide import Slide

from ..unitutil.cxml import element
from ..unitutil.mock import (
    call, class_mock, instance_mock, method_mock, property_mock
)


class DescribePresentationPart(object):

    def it_provides_access_to_its_presentation(self, prs_fixture):
        prs_part, Presentation_, prs_elm, prs_ = prs_fixture
        prs = prs_part.presentation
        Presentation_.assert_called_once_with(prs_elm, prs_part)
        assert prs is prs_

    def it_provides_access_to_its_core_properties(self, core_props_fixture):
        prs_part, core_properties_ = core_props_fixture
        core_properties = prs_part.core_properties
        assert core_properties is core_properties_

    def it_provides_access_to_a_related_slide(self, slide_fixture):
        prs_part, rId, slide_ = slide_fixture
        slide = prs_part.related_slide(rId)
        prs_part.related_parts.__getitem__.assert_called_once_with(rId)
        assert slide is slide_

    def it_can_rename_related_slide_parts(self, rename_fixture):
        prs_part, rIds, getitem_ = rename_fixture[:3]
        calls, slide_parts, expected_names = rename_fixture[3:]
        prs_part.rename_slide_parts(rIds)
        assert getitem_.call_args_list == calls
        assert [sp.partname for sp in slide_parts] == expected_names

    def it_can_save_the_package_to_a_file(self, save_fixture):
        prs_part, file_, package_ = save_fixture
        prs_part.save(file_)
        package_.save.assert_called_once_with(file_)

    def it_can_add_a_new_slide(self, add_slide_fixture):
        prs_part, slide_layout_, SlidePart_, partname = add_slide_fixture[:4]
        package_, slide_part_, rId_, slide_ = add_slide_fixture[4:]

        rId, slide = prs_part.add_slide(slide_layout_)

        SlidePart_.new.assert_called_once_with(
            partname, package_, slide_layout_
        )
        prs_part.relate_to.assert_called_once_with(
            prs_part, slide_part_, RT.SLIDE
        )
        assert rId is rId_
        assert slide is slide_

    def it_knows_the_next_slide_partname_to_help(self, next_fixture):
        prs_part, partname = next_fixture
        assert prs_part._next_slide_partname == partname

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_slide_fixture(
            self, package_, slide_layout_, SlidePart_, slide_part_, slide_,
            _next_slide_partname_prop_, relate_to_):
        prs_part = PresentationPart(None, None, None, package_)
        partname = _next_slide_partname_prop_.return_value
        rId_ = 'rId42'
        SlidePart_.new.return_value = slide_part_
        relate_to_.return_value = rId_
        slide_part_.slide = slide_
        return (
            prs_part, slide_layout_, SlidePart_, partname, package_,
            slide_part_, rId_, slide_
        )

    @pytest.fixture
    def core_props_fixture(self, package_, core_properties_):
        prs_part = PresentationPart(None, None, None, package_)
        package_.core_properties = core_properties_
        return prs_part, core_properties_

    @pytest.fixture
    def next_fixture(self):
        prs_elm = element('p:presentation/p:sldIdLst/(p:sldId,p:sldId)')
        prs_part = PresentationPart(None, None, prs_elm)
        partname = PackURI('/ppt/slides/slide3.xml')
        return prs_part, partname

    @pytest.fixture
    def prs_fixture(self, Presentation_, prs_):
        prs_elm = element('p:presentation')
        prs_part = PresentationPart(None, None, prs_elm)
        return prs_part, Presentation_, prs_elm, prs_

    @pytest.fixture
    def rename_fixture(self, related_parts_prop_):
        prs_part = PresentationPart(None, None, None)
        rIds = ('rId1', 'rId2')
        getitem_ = related_parts_prop_.return_value.__getitem__
        calls = [call('rId1'), call('rId2')]
        slide_parts = [
            SlidePart(None, None, None),
            SlidePart(None, None, None),
        ]
        expected_names = [
            PackURI('/ppt/slides/slide1.xml'),
            PackURI('/ppt/slides/slide2.xml'),
        ]
        getitem_.side_effect = slide_parts
        return (
            prs_part, rIds, getitem_, calls, slide_parts,
            expected_names
        )

    @pytest.fixture
    def save_fixture(self, package_):
        prs_part = PresentationPart(None, None, None, package_)
        file_ = 'foobar.docx'
        return prs_part, file_, package_

    @pytest.fixture
    def slide_fixture(self, slide_, related_parts_prop_):
        prs_part = PresentationPart(None, None, None, None)
        rId = 'rId42'
        related_parts_ = related_parts_prop_.return_value
        related_parts_.__getitem__.return_value.slide = slide_
        return prs_part, rId, slide_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def core_properties_(self, request):
        return instance_mock(request, CorePropertiesPart)

    @pytest.fixture
    def _next_slide_partname_prop_(self, request):
        return property_mock(
            request, PresentationPart, '_next_slide_partname'
        )

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def Presentation_(self, request, prs_):
        return class_mock(
            request, 'pptx.parts.presentation.Presentation',
            return_value=prs_
        )

    @pytest.fixture
    def prs_(self, request):
        return instance_mock(request, Presentation)

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(
            request, PresentationPart, 'relate_to', autospec=True
        )

    @pytest.fixture
    def related_parts_prop_(self, request):
        return property_mock(request, PresentationPart, 'related_parts')

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayoutPart)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)

    @pytest.fixture
    def SlidePart_(self, request):
        return class_mock(request, 'pptx.parts.presentation.SlidePart')
