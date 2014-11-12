# encoding: utf-8

"""
Test suite for pptx.package module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.package import Package
from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import ImagePart
from pptx.parts.presentation import PresentationPart


from .unitutil.file import absjoin, test_file_dir
from .unitutil.mock import instance_mock, property_mock


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')


class DescribePackage(object):

    def it_loads_default_template_when_opened_with_no_path(self):
        prs = Package.open().presentation
        assert prs is not None
        slide_masters = prs.slide_masters
        assert slide_masters is not None
        assert len(slide_masters) == 1
        slide_layouts = slide_masters[0].slide_layouts
        assert slide_layouts is not None
        assert len(slide_layouts) == 11

    def it_gathers_package_image_parts_on_open(self):
        pkg = Package.open(images_pptx_path)
        assert len(pkg._images) == 7

    def it_provides_ref_to_package_presentation_part(self):
        pkg = Package.open()
        assert isinstance(pkg.presentation, PresentationPart)

    def it_provides_access_to_its_core_properties_part(self):
        pkg = Package.open()
        assert isinstance(pkg.core_properties, CoreProperties)

    def it_can_get_or_add_an_image_part(self, image_part_fixture):
        package, image_file, image_part_ = image_part_fixture

        image_part = package.get_or_add_image_part(image_file)

        package._images.add_image.assert_called_once_with(image_file)
        assert image_part is image_part_

    def it_can_save_itself_to_a_pptx_file(self, temp_pptx_path):
        """
        Package.save produces a .pptx with plausible contents
        """
        # setup ------------------------
        pkg = Package.open()
        # exercise ---------------------
        pkg.save(temp_pptx_path)
        # verify -----------------------
        pkg = Package.open(temp_pptx_path)
        prs = pkg.presentation
        assert prs is not None
        slide_masters = prs.slide_masters
        assert slide_masters is not None
        assert len(slide_masters) == 1
        slide_layouts = slide_masters[0].slide_layouts
        assert slide_layouts is not None
        assert len(slide_layouts) == 11

    # fixtures ---------------------------------------------

    @pytest.fixture
    def image_part_fixture(self, _images_, image_part_):
        package = Package()
        image_file = 'foobar.png'
        package._images.add_image.return_value = image_part_
        return package, image_file, image_part_

    @pytest.fixture
    def temp_pptx_path(self, tmpdir):
        return absjoin(str(tmpdir), 'test-pptx.pptx')

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def _images_(self, request):
        return property_mock(request, Package, '_images')
