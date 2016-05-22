# encoding: utf-8

"""
Test suite for pptx.parts.presentation module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.package import Package
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.presentation import PresentationPart
from pptx.presentation import Presentation

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, instance_mock


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

    def it_can_save_the_package_to_a_file(self, save_fixture):
        prs_part, file_, package_ = save_fixture
        prs_part.save(file_)
        package_.save.assert_called_once_with(file_)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def core_props_fixture(self, package_, core_properties_):
        prs_part = PresentationPart(None, None, None, package_)
        package_.core_properties = core_properties_
        return prs_part, core_properties_

    @pytest.fixture
    def prs_fixture(self, Presentation_, prs_):
        prs_elm = element('p:presentation')
        prs_part = PresentationPart(None, None, prs_elm)
        # prs_cxml, expected_value = request.param
        # prs = Presentation(element(prs_cxml), None)
        return prs_part, Presentation_, prs_elm, prs_

    @pytest.fixture
    def save_fixture(self, package_):
        prs_part = PresentationPart(None, None, None, package_)
        file_ = 'foobar.docx'
        return prs_part, file_, package_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def core_properties_(self, request):
        return instance_mock(request, CorePropertiesPart)

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
