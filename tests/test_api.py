"""Unit-test suite for `pptx.api` module."""

from __future__ import annotations

import os

import pytest

import pptx
from pptx.api import Presentation, _is_pptx_package
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.parts.presentation import PresentationPart

from .unitutil.mock import class_mock, instance_mock


class DescribePresentation(object):
    def it_opens_default_template_on_no_path_provided(self, call_fixture):
        Package_, path, prs_ = call_fixture
        prs = Presentation()
        Package_.open.assert_called_once_with(path)
        assert prs is prs_

    @pytest.mark.parametrize(
        "content_type",
        [
            CT.PML_PRESENTATION_MAIN,
            CT.PML_PRES_MACRO_MAIN,
            CT.PML_TEMPLATE_MAIN,
            CT.PML_SLIDESHOW_MAIN,
        ],
    )
    def it_opens_any_presentationml_main_content_type(
        self, content_type, Package_, prs_, prs_part_
    ):
        Package_.open.return_value.main_document_part = prs_part_
        prs_part_.content_type = content_type
        prs_part_.presentation = prs_

        prs = Presentation("foobar.pptx")

        assert prs is prs_

    def it_raises_on_non_presentationml_main_content_type(self, Package_, prs_part_):
        Package_.open.return_value.main_document_part = prs_part_
        prs_part_.content_type = "application/vnd.ms-word"

        with pytest.raises(ValueError, match="is not a PowerPoint file"):
            Presentation("foobar.docx")

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def call_fixture(self, Package_, prs_, prs_part_):
        path = os.path.abspath(
            os.path.join(os.path.split(pptx.__file__)[0], "templates", "default.pptx")
        )
        Package_.open.return_value.main_document_part = prs_part_
        prs_part_.content_type = CT.PML_PRESENTATION_MAIN
        prs_part_.presentation = prs_
        return Package_, path, prs_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Package_(self, request):
        return class_mock(request, "pptx.api.Package")

    @pytest.fixture
    def prs_(self, request):
        return instance_mock(request, Presentation)

    @pytest.fixture
    def prs_part_(self, request):
        return instance_mock(request, PresentationPart)


class Describe_is_pptx_package(object):
    @pytest.mark.parametrize(
        "content_type",
        [
            CT.PML_PRESENTATION_MAIN,
            CT.PML_PRES_MACRO_MAIN,
            CT.PML_TEMPLATE_MAIN,
            CT.PML_SLIDESHOW_MAIN,
        ],
    )
    def it_is_True_for_each_presentationml_main_content_type(
        self, content_type, prs_part_
    ):
        prs_part_.content_type = content_type
        assert _is_pptx_package(prs_part_) is True

    def it_is_False_for_any_other_content_type(self, prs_part_):
        prs_part_.content_type = "application/vnd.ms-excel"
        assert _is_pptx_package(prs_part_) is False

    # fixture components ---------------------------------------------

    @pytest.fixture
    def prs_part_(self, request):
        return instance_mock(request, PresentationPart)
