# encoding: utf-8

"""Test suite for pptx.slides module."""

from __future__ import absolute_import

import pytest

from lxml import objectify
from mock import Mock, patch, PropertyMock

from pptx.opc import packaging
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import namespaces
from pptx.parts.slides import (
    _BaseSlide, Slide, SlideCollection, SlideLayout, SlideMaster
)
from pptx.presentation import Package, Presentation
from pptx.shapes.shapetree import ShapeCollection

from ..unitutil import (
    absjoin, class_mock, instance_mock, method_mock, parse_xml_file,
    serialize_xml, test_file_dir
)


test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')

nsmap = namespaces('a', 'r', 'p')


def actual_xml(elm):
    objectify.deannotate(elm, cleanup_namespaces=True)
    return serialize_xml(elm, pretty_print=True)


def _sldLayout1():
    path = absjoin(test_file_dir, 'slideLayout1.xml')
    sldLayout = parse_xml_file(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = ShapeCollection(spTree)
    return shapes


class Describe_BaseSlide(object):

    def it_knows_the_name_of_the_slide(self, base_slide):
        """_BaseSlide.name value is correct"""
        # setup ------------------------
        base_slide._element = _sldLayout1()
        # exercise ---------------------
        name = base_slide.name
        # verify -----------------------
        expected = 'Title Slide'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    def it_provides_access_to_the_shapes_on_the_slide(self, base_slide):
        """_BaseSlide.shapes is expected size after _load()"""
        # setup ------------------------
        path = absjoin(test_file_dir, 'slide1.xml')
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/slides/slide1.xml'
        with open(path, 'r') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = []
        part_dict = {}
        base_slide._load(pkgpart, part_dict)
        # exercise ---------------------
        shapes = base_slide.shapes
        # verify -----------------------
        assert len(shapes) == 9

    @patch('pptx.parts.slides._BaseSlide._package', new_callable=PropertyMock)
    def it_can_add_an_image_part_to_the_slide(self, _package, base_slide):
        """_BaseSlide._add_image() returns (image, rel) tuple"""
        # setup ------------------------
        base_slide = base_slide
        image = Mock(name='image')
        rel = Mock(name='rel')
        base_slide._package._images.add_image.return_value = image
        base_slide._add_relationship = Mock('_add_relationship')
        base_slide._add_relationship.return_value = rel
        file = test_image_path
        # exercise ---------------------
        retval_image, retval_rel = base_slide._add_image(file)
        # verify -----------------------
        base_slide._package._images.add_image.assert_called_once_with(file)
        base_slide._add_relationship.assert_called_once_with(RT.IMAGE, image)
        assert retval_image is image
        assert retval_rel is rel

    def it_knows_what_to_do_after_the_slide_is_unmarshaled(self):
        pass

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def base_slide(self):
        return _BaseSlide()


class DescribeSlide(object):

    def it_passes_its_content_type_to_BasePart_on_construction(self, slide):
        """Slide constructor sets correct content type"""
        # exercise ---------------------
        content_type = slide._content_type
        # verify -----------------------
        expected = CT.PML_SLIDE
        actual = content_type
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    def it_establishes_a_relationship_to_its_slide_layout_on_construction(
            self):
        """Slide(slidelayout) adds relationship slide->slidelayout"""
        # setup ------------------------
        slidelayout = SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        # exercise ---------------------
        slide = Slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(slide._relationships)
        msg = ("expected len(slide._relationships) of %d, got %d"
               % (expected, actual))
        assert actual == expected, msg
        # verify values ---------------
        rel = slide._relationships[0]
        expected = ('rId1', RT.SLIDE_LAYOUT, slidelayout)
        actual = (rel.rId, rel.reltype, rel.target)
        msg = "expected relationship\n%s\ngot\n%s" % (expected, actual)
        assert actual == expected, msg

    def it_creates_a_minimal_sld_element_on_construction(self, slide):
        """Slide._element is minimal sld on construction"""
        # setup ------------------------
        path = absjoin(test_file_dir, 'minimal_slide.xml')
        # exercise ---------------------
        elm = slide._element
        # verify -----------------------
        with open(path, 'r') as f:
            expected_xml = f.read()
        assert actual_xml(elm) == expected_xml

    def it_has_slidelayout_property_of_none_on_construction(self, slide):
        """Slide.slidelayout property None on construction"""
        assert slide.slidelayout is None

    def it_sets_slidelayout_on_load(self, slide):
        """Slide._load() sets slidelayout"""
        # setup ------------------------
        path = absjoin(test_file_dir, 'slide1.xml')
        slidelayout = Mock(name='slideLayout')
        slidelayout.partname = '/ppt/slideLayouts/slideLayout1.xml'
        rel = Mock(name='pptx.packaging.Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT.SLIDE_LAYOUT
        rel.target = slidelayout
        pkgpart = Mock(name='pptx.packaging.Part')
        with open(path, 'rb') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = [rel]
        part_dict = {slidelayout.partname: slidelayout}
        slide_ = slide._load(pkgpart, part_dict)
        # exercise ---------------------
        retval = slide_.slidelayout
        # verify -----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        assert actual == expected, msg

    def it_knows_the_minimal_element_xml_for_a_slide(self, slide):
        """Slide._minimal_element generates correct XML"""
        # setup ------------------------
        path = absjoin(test_file_dir, 'minimal_slide.xml')
        # exercise ---------------------
        sld = slide._minimal_element
        # verify -----------------------
        with open(path, 'r') as f:
            expected_xml = f.read()
        assert actual_xml(sld) == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def slide(self):
        return Slide()


class DescribeSlideCollection(object):

    def it_can_add_a_new_slide(
            self, slides, Slide_, slidelayout_, _rename_slides_,
            add_relationship_, slide_):
        slide = slides.add_slide(slidelayout_)
        # verify -----------------------
        Slide_.assert_called_once_with(slidelayout_)
        _rename_slides_.assert_called_once_with()
        add_relationship_.assert_called_once_with(RT.SLIDE, slide_)
        assert slide is slide_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_relationship_(self, request):
        return method_mock(request, Presentation, '_add_relationship')

    @pytest.fixture
    def _rename_slides_(self, request):
        return method_mock(request, SlideCollection, '_rename_slides')

    @pytest.fixture
    def Slide_(self, request, slide_):
        return class_mock(
            request, 'pptx.parts.slides.Slide', return_value=slide_
        )

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slidelayout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slides(self):
        prs = Presentation()
        return SlideCollection(prs)


class DescribeSlideLayout(object):

    def _loaded_slidelayout(self, prs_slidemaster=None):
        """
        Return SlideLayout instance loaded using mocks. *prs_slidemaster* is
        an already-loaded model-side SlideMaster instance (or mock, as
        appropriate to calling test).
        """
        # partname for related slideMaster
        sldmaster_partname = '/ppt/slideMasters/slideMaster1.xml'
        # path to test slideLayout XML
        slidelayout_path = absjoin(test_file_dir, 'slideLayout1.xml')
        # model-side slideMaster part
        if prs_slidemaster is None:
            prs_slidemaster = Mock(spec=SlideMaster)
        # a part dict containing the already-loaded model-side slideMaster
        loaded_part_dict = {sldmaster_partname: prs_slidemaster}
        # a slideMaster package part for rel target
        pkg_slidemaster_part = Mock(spec=packaging.Part)
        pkg_slidemaster_part.partname = sldmaster_partname
        # a package-side relationship from slideLayout to its slideMaster
        rel = Mock(name='pptx.packaging.Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT.SLIDE_MASTER
        rel.target = pkg_slidemaster_part
        # the slideLayout package part to send to _load()
        pkg_slidelayout_part = Mock(spec=packaging.Part)
        pkg_slidelayout_part.relationships = [rel]
        with open(slidelayout_path, 'rb') as f:
            pkg_slidelayout_part.blob = f.read()
        # _load and return
        slidelayout = SlideLayout()
        return slidelayout._load(pkg_slidelayout_part, loaded_part_dict)

    def test__load_sets_slidemaster(self):
        """SlideLayout._load() sets slidemaster"""
        # setup ------------------------
        prs_slidemaster = Mock(spec=SlideMaster)
        # exercise ---------------------
        loaded_slidelayout = self._loaded_slidelayout(prs_slidemaster)
        # verify -----------------------
        expected = prs_slidemaster
        actual = loaded_slidelayout.slidemaster
        msg = "expected: %s, got %s" % (expected, actual)
        assert actual == expected, msg

    def test_slidemaster_raises_on_ref_before_assigned(self, slidelayout):
        """SlideLayout.slidemaster raises on referenced before assigned"""
        with pytest.raises(AssertionError):
            slidelayout.slidemaster

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def slidelayout(self):
        return SlideLayout()


class DescribeSlideMaster(object):

    def test_slidelayouts_property_empty_on_construction(self, slidemaster):
        assert len(slidemaster.slidelayouts) == 0

    def test_slidelayouts_correct_length_after_open(self):
        """SlideMaster.slidelayouts correct length after open"""
        # setup ------------------------
        pkg = Package(test_pptx_path)
        slidemaster = pkg.presentation.slidemasters[0]
        # exercise ---------------------
        slidelayouts = slidemaster.slidelayouts
        # verify -----------------------
        assert len(slidelayouts) == 11

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def slidemaster(self):
        return SlideMaster()
