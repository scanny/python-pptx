# encoding: utf-8

"""Test suite for pptx.slides module."""

from __future__ import absolute_import

from hamcrest import assert_that, is_
from mock import Mock, patch, PropertyMock

from pptx.opc import packaging
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import namespaces
from pptx.parts.slides import (
    _BaseSlide, Slide, SlideCollection, SlideLayout, SlideMaster
)
from pptx.presentation import Package, Presentation
from pptx.shapes.shapetree import ShapeCollection

from ..unitutil import absjoin, parse_xml_file, TestCase, test_file_dir


test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')

nsmap = namespaces('a', 'r', 'p')


def _sldLayout1():
    path = absjoin(test_file_dir, 'slideLayout1.xml')
    sldLayout = parse_xml_file(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = ShapeCollection(spTree)
    return shapes


class Test_BaseSlide(TestCase):
    """Test _BaseSlide"""
    def setUp(self):
        self.base_slide = _BaseSlide()

    def test_name_value(self):
        """_BaseSlide.name value is correct"""
        # setup ------------------------
        self.base_slide._element = _sldLayout1()
        # exercise ---------------------
        name = self.base_slide.name
        # verify -----------------------
        expected = 'Title Slide'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_shapes_size_after__load(self):
        """_BaseSlide.shapes is expected size after _load()"""
        # setup ------------------------
        path = absjoin(test_file_dir, 'slide1.xml')
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/slides/slide1.xml'
        with open(path, 'r') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = []
        part_dict = {}
        self.base_slide._load(pkgpart, part_dict)
        # exercise ---------------------
        shapes = self.base_slide.shapes
        # verify -----------------------
        self.assertLength(shapes, 9)

    @patch('pptx.parts.slides._BaseSlide._package', new_callable=PropertyMock)
    def test__add_image_collaboration(self, _package):
        """_BaseSlide._add_image() returns (image, rel) tuple"""
        # setup ------------------------
        base_slide = self.base_slide
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
        assert_that(retval_image, is_(image))
        assert_that(retval_rel, is_(rel))


class TestSlide(TestCase):
    """Test Slide"""
    def setUp(self):
        self.sld = Slide()

    def test_constructor_sets_correct_content_type(self):
        """Slide constructor sets correct content type"""
        # exercise ---------------------
        content_type = self.sld._content_type
        # verify -----------------------
        expected = CT.PML_SLIDE
        actual = content_type
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_construction_adds_slide_layout_relationship(self):
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
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = slide._relationships[0]
        expected = ('rId1', RT.SLIDE_LAYOUT, slidelayout)
        actual = (rel._rId, rel._reltype, rel.target)
        msg = "expected relationship\n%s\ngot\n%s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__element_minimal_sld_on_construction(self):
        """Slide._element is minimal sld on construction"""
        # setup ------------------------
        path = absjoin(test_file_dir, 'minimal_slide.xml')
        # exercise ---------------------
        elm = self.sld._element
        # verify -----------------------
        with open(path, 'r') as f:
            expected_xml = f.read()
        self.assertEqualLineByLine(expected_xml, elm)

    def test_slidelayout_property_none_on_construction(self):
        """Slide.slidelayout property None on construction"""
        # verify -----------------------
        self.assertIsProperty(self.sld, 'slidelayout', None)

    def test__load_sets_slidelayout(self):
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
        slide = self.sld._load(pkgpart, part_dict)
        # exercise ---------------------
        retval = slide.slidelayout
        # verify -----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__minimal_element_xml(self):
        """Slide._minimal_element generates correct XML"""
        # setup ------------------------
        path = absjoin(test_file_dir, 'minimal_slide.xml')
        # exercise ---------------------
        sld = self.sld._minimal_element
        # verify -----------------------
        with open(path, 'r') as f:
            expected_xml = f.read()
        self.assertEqualLineByLine(expected_xml, sld)


class TestSlideCollection(TestCase):
    """Test SlideCollection"""
    def setUp(self):
        prs = Presentation()
        self.slides = SlideCollection(prs)

    def test_add_slide_returns_slide(self):
        """SlideCollection.add_slide() returns instance of Slide"""
        # exercise ---------------------
        retval = self.slides.add_slide(None)
        # verify -----------------------
        self.assertIsInstance(retval, Slide)

    def test_add_slide_sets_slidelayout(self):
        """
        SlideCollection.add_slide() sets Slide.slidelayout

        Kind of a throw-away test, but was helpful for initial debugging.
        """
        # setup ------------------------
        slidelayout = Mock(name='slideLayout')
        slidelayout.shapes = []
        slide = self.slides.add_slide(slidelayout)
        # exercise ---------------------
        retval = slide.slidelayout
        # verify -----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_slide_adds_slide_layout_relationship(self):
        """SlideCollection.add_slide() adds relationship prs->slide"""
        # setup ------------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = SlideLayout()
        slidelayout._shapes = []
        # exercise ---------------------
        slide = slides.add_slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(prs._relationships)
        msg = ("expected len(prs._relationships) of %d, got %d"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = prs._relationships[0]
        expected = ('rId1', RT.SLIDE, slide)
        actual = (rel._rId, rel._reltype, rel.target)
        msg = ("expected relationship 1:, got 2:\n1: %s\n2: %s"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test_add_slide_sets_partname(self):
        """SlideCollection.add_slide() sets partname of new slide"""
        # setup ------------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = SlideLayout()
        slidelayout._shapes = []
        # exercise ---------------------
        slide = slides.add_slide(slidelayout)
        # verify -----------------------
        expected = '/ppt/slides/slide1.xml'
        actual = slide.partname
        msg = "expected partname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class TestSlideLayout(TestCase):
    """Test SlideLayout"""
    def setUp(self):
        self.slidelayout = SlideLayout()

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
        self.assertEqual(expected, actual, msg)

    def test_slidemaster_is_readonly(self):
        """SlideLayout.slidemaster is read-only"""
        # verify -----------------------
        self.assertIsReadOnly(self.slidelayout, 'slidemaster')

    def test_slidemaster_raises_on_ref_before_assigned(self):
        """SlideLayout.slidemaster raises on referenced before assigned"""
        with self.assertRaises(AssertionError):
            self.slidelayout.slidemaster


class TestSlideMaster(TestCase):
    """Test SlideMaster"""
    def setUp(self):
        self.sldmaster = SlideMaster()

    def test_slidelayouts_property_empty_on_construction(self):
        """SlideMaster.slidelayouts property empty on construction"""
        # verify -----------------------
        self.assertIsSizedProperty(self.sldmaster, 'slidelayouts', 0)

    def test_slidelayouts_correct_length_after_open(self):
        """SlideMaster.slidelayouts correct length after open"""
        # setup ------------------------
        pkg = Package(test_pptx_path)
        slidemaster = pkg.presentation.slidemasters[0]
        # exercise ---------------------
        slidelayouts = slidemaster.slidelayouts
        # verify -----------------------
        self.assertLength(slidelayouts, 11)
