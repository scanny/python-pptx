# encoding: utf-8

"""Test suite for pptx.presentation module."""

from __future__ import absolute_import, print_function

import gc
import pytest

from mock import Mock

from pptx.exceptions import InvalidPackageError
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml_bytes
from pptx.oxml.ns import namespaces, qn
from pptx.parts.coreprops import CoreProperties
from pptx.parts.part import PartCollection
from pptx.parts.slides import Slide, SlideCollection, SlideLayout, SlideMaster
from pptx.presentation import Package, Part, Presentation

from .opc.unitdata.rels import a_rels
from .unitutil import absjoin, parse_xml_file, test_file_dir


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')

nsmap = namespaces('a', 'r', 'p')


class DescribePackage(object):

    def it_loads_default_template_when_constructed_with_no_path(self):
        prs = Package().presentation
        assert prs is not None
        slidemasters = prs.slidemasters
        assert slidemasters is not None
        assert len(slidemasters) == 1
        slidelayouts = slidemasters[0].slidelayouts
        assert slidelayouts is not None
        assert len(slidelayouts) == 11

    def it_tracks_instances_of_itself(self):
        pkg = Package()
        assert pkg in Package.instances()

    def it_garbage_collects_refs_to_old_instances_of_itself(self):
        pkg = Package()
        pkg1_repr = "%r" % pkg
        pkg = Package()
        # pkg2_repr = "%r" % pkg
        gc.collect()
        reprs = [repr(pkg_inst) for pkg_inst in Package.instances()]
        assert pkg1_repr not in reprs

    def it_knows_which_instance_contains_a_specified_part(self):
        # setup ------------------------
        pkg1 = Package(test_pptx_path)  # noqa
        pkg2 = Package(test_pptx_path)
        slide = pkg2.presentation.slides[0]
        # exercise ---------------------
        found_pkg = Package.containing(slide)
        # verify -----------------------
        assert found_pkg == pkg2

    def it_raises_when_no_package_contains_specified_part(self):
        # setup ------------------------
        pkg = Package(test_pptx_path)
        pkg.presentation  # does nothing, just needed to fake out pep8 warning
        part = Mock(name='part')
        # verify -----------------------
        with pytest.raises(KeyError):
            Package.containing(part)

    def it_gathers_packages_image_parts_on_open(self):
        """Package open gathers image parts into image collection"""
        pkg = Package(images_pptx_path)
        assert len(pkg._images) == 7

    def it_returns_an_instance_of_presentation_from_open(self):
        pkg = Package()
        assert isinstance(pkg.presentation, Presentation)

    def it_provides_access_to_the_package_core_properties(self):
        pkg = Package()
        assert isinstance(pkg.core_properties, CoreProperties)

    def it_can_save_itself_to_a_pptx_file(self, temp_pptx_path):
        """Package.save produces a .pptx with plausible contents"""
        # setup ------------------------
        pkg = Package()
        # exercise ---------------------
        pkg.save(temp_pptx_path)
        # verify -----------------------
        pkg = Package(temp_pptx_path)
        prs = pkg.presentation
        assert prs is not None
        slidemasters = prs.slidemasters
        assert slidemasters is not None
        assert len(slidemasters) == 1
        slidelayouts = slidemasters[0].slidelayouts
        assert slidelayouts is not None
        assert len(slidelayouts) == 11

    # fixtures ---------------------------------------------

    @pytest.fixture
    def temp_pptx_path(self, tmpdir):
        return absjoin(str(tmpdir), 'test-pptx.pptx')


class DescribePart(object):

    def it_constructs_presentation_for_rt_officedocument(self):
        obj = Part(RT.OFFICE_DOCUMENT, CT.PML_PRESENTATION_MAIN)
        assert isinstance(obj, Presentation)

    def it_constructs_slide_for_rt_slide(self):
        obj = Part(RT.SLIDE, CT.PML_SLIDE)
        assert isinstance(obj, Slide)

    def it_constructs_slidelayout_for_rt_slidelayout(self):
        obj = Part(RT.SLIDE_LAYOUT, CT.PML_SLIDE_LAYOUT)
        assert isinstance(obj, SlideLayout)

    def it_constructs_slidemaster_for_rt_slidemaster(self):
        obj = Part(RT.SLIDE_MASTER, CT.PML_SLIDE_MASTER)
        assert isinstance(obj, SlideMaster)

    def it_raises_on_construct_attempt_with_invalid_prs_content_type(self):
        with pytest.raises(InvalidPackageError):
            Part(RT.OFFICE_DOCUMENT, CT.PML_SLIDE_MASTER)


class DescribePresentation(object):

    def it_provides_access_to_the_slide_masters(self, prs):
        assert isinstance(prs.slidemasters, PartCollection)

    def it_provides_access_to_the_slides(self, prs):
        assert isinstance(prs.slides, SlideCollection)

    def test__blob_rewrites_sldIdLst(self):
        """Presentation._blob rewrites sldIdLst"""
        # setup ------------------------
        rels = a_rels()
        rels = rels.with_tuple_targets(2, RT.SLIDE_MASTER)
        rels = rels.with_tuple_targets(3, RT.SLIDE)
        rels = rels.build()
        prs = Presentation()
        prs._relationships = rels
        prs.partname = '/ppt/presentation.xml'
        path = absjoin(test_file_dir, 'presentation.xml')
        prs._element = parse_xml_file(path).getroot()
        # exercise ---------------------
        blob = prs._blob
        # verify -----------------------
        presentation = parse_xml_bytes(blob)
        sldIds = presentation.xpath('./p:sldIdLst/p:sldId', namespaces=nsmap)
        expected = ['rId3', 'rId4', 'rId5']
        actual = [sldId.get(qn('r:id')) for sldId in sldIds]
        msg = "expected ordering %s, got %s" % (expected, actual)
        assert actual == expected, msg

    # fixtures ---------------------------------------------

    @pytest.fixture
    def prs(self):
        prs = Presentation()
        prs._element = Mock(name='_element')
        return prs
