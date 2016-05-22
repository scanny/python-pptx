# encoding: utf-8

"""
Test suite for pptx.api module
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import os

import pytest

from pptx.api import Presentation
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.parts.presentation import PresentationPart

from .unitutil.mock import class_mock, instance_mock


class DescribePresentation(object):

    def it_opens_default_template_on_no_path_provided(self, call_fixture):
        Package_, path, prs_ = call_fixture
        prs = Presentation()
        Package_.open.assert_called_once_with(path)
        assert prs is prs_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def call_fixture(self, Package_, prs_, prs_part_):
        path = os.path.abspath(
            os.path.join(
                os.path.split(__file__)[0], '../pptx/templates',
                'default.pptx'
            )
        )
        Package_.open.return_value.main_document_part = prs_part_
        prs_part_.content_type = CT.PML_PRESENTATION_MAIN
        prs_part_.presentation = prs_
        return Package_, path, prs_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Package_(self, request):
        return class_mock(request, 'pptx.api.Package')

    @pytest.fixture
    def prs_(self, request):
        return instance_mock(request, Presentation)

    @pytest.fixture
    def prs_part_(self, request):
        return instance_mock(request, PresentationPart)
