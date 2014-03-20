# encoding: utf-8

"""
Test suite for pptx.package module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.parts.coreprops import CoreProperties
from pptx.presentation import Package
from pptx.parts.presentation import PresentationPart


from .unitutil import absjoin, test_file_dir


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

    def it_provides_ref_to_package_core_properties_part(self):
        pkg = Package.open()
        assert isinstance(pkg.core_properties, CoreProperties)

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
    def temp_pptx_path(self, tmpdir):
        return absjoin(str(tmpdir), 'test-pptx.pptx')
