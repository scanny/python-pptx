# encoding: utf-8

"""Test suite for pptx.presentation module."""

from __future__ import absolute_import

import gc
import os

from hamcrest import assert_that, instance_of, is_, is_in, is_not
from mock import Mock

from pptx.exceptions import InvalidPackageError
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.rels import _Relationship, _RelationshipCollection
from pptx.oxml import oxml_fromstring, oxml_parse
from pptx.parts.coreprops import CoreProperties
from pptx.parts.part import BasePart
from pptx.parts.slides import _Slide, _SlideLayout, _SlideMaster
from pptx.presentation import _Package, _Part, Presentation
from pptx.spec import namespaces, qtag

from .unitutil import absjoin, TestCase, test_file_dir


images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')

nsmap = namespaces('a', 'r', 'p')


class PartBuilder(object):
    """Builder class for test Parts"""
    def __init__(self):
        self.partname = '/ppt/slides/slide1.xml'

    def with_partname(self, partname):
        self.partname = partname
        return self

    def build(self):
        p = BasePart()
        p.partname = self.partname
        return p


class RelationshipCollectionBuilder(object):
    """Builder class for test RelationshipCollections"""
    partname_tmpls = {RT.SLIDE_MASTER: '/ppt/slideMasters/slideMaster%d.xml',
                      RT.SLIDE: '/ppt/slides/slide%d.xml'}

    def __init__(self):
        self.relationships = []
        self.next_rel_num = 1
        self.next_partnums = {}
        self.reltype_ordering = None

    def with_ordering(self, *reltypes):
        self.reltype_ordering = tuple(reltypes)
        return self

    def with_tuple_targets(self, count, reltype):
        for i in range(count):
            rId = self._next_rId
            partname = self._next_tuple_partname(reltype)
            target = PartBuilder().with_partname(partname).build()
            rel = _Relationship(rId, reltype, target)
            self.relationships.append(rel)
        return self

    def _next_partnum(self, reltype):
        if reltype not in self.next_partnums:
            self.next_partnums[reltype] = 1
        partnum = self.next_partnums[reltype]
        self.next_partnums[reltype] = partnum + 1
        return partnum

    @property
    def _next_rId(self):
        rId = 'rId%d' % self.next_rel_num
        self.next_rel_num += 1
        return rId

    def _next_tuple_partname(self, reltype):
        partname_tmpl = self.partname_tmpls[reltype]
        partnum = self._next_partnum(reltype)
        return partname_tmpl % partnum

    def build(self):
        rels = _RelationshipCollection()
        for rel in self.relationships:
            rels._additem(rel)
        if self.reltype_ordering:
            rels._reltype_ordering = self.reltype_ordering
        return rels


class Test_Package(TestCase):
    """Test _Package"""
    def setUp(self):
        self.test_pptx_path = absjoin(test_file_dir, 'test_python-pptx.pptx')
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)

    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)

    def test_construction_with_no_path_loads_default_template(self):
        """_Package() call with no path loads default template"""
        prs = _Package().presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))

    def test_instances_are_tracked(self):
        """_Package instances are tracked"""
        pkg = _Package()
        self.assertIn(pkg, _Package.instances())

    def test_instance_refs_are_garbage_collected(self):
        """_Package instance refs are garbage collected with old instances"""
        pkg = _Package()
        pkg1_repr = "%r" % pkg
        pkg = _Package()
        # pkg2_repr = "%r" % pkg
        gc.collect()
        reprs = [repr(pkg_inst) for pkg_inst in _Package.instances()]
        assert_that(pkg1_repr, is_not(is_in(reprs)))

    def test_containing_returns_correct_pkg(self):
        """_Package.containing() returns right package instance"""
        # setup ------------------------
        pkg1 = _Package(test_pptx_path)
        pkg1.presentation  # does nothing, just needed to fake out pep8 warning
        pkg2 = _Package(test_pptx_path)
        slide = pkg2.presentation.slides[0]
        # exercise ---------------------
        found_pkg = _Package.containing(slide)
        # verify -----------------------
        expected = pkg2
        actual = found_pkg
        msg = "expected %r, got %r" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_containing_raises_on_no_pkg_contains_part(self):
        """_Package.containing(part) raises on no package contains part"""
        # setup ------------------------
        pkg = _Package(test_pptx_path)
        pkg.presentation  # does nothing, just needed to fake out pep8 warning
        part = Mock(name='part')
        # verify -----------------------
        with self.assertRaises(KeyError):
            _Package.containing(part)

    def test_open_gathers_image_parts(self):
        """_Package open gathers image parts into image collection"""
        # exercise ---------------------
        pkg = _Package(images_pptx_path)
        # verify -----------------------
        expected = 7
        actual = len(pkg._images)
        msg = "expected image count of %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_presentation_presentation_after_open(self):
        """_Package.presentation is instance of Presentation after open()"""
        # setup ------------------------
        cls = Presentation
        pkg = _Package()
        # exercise ---------------------
        obj = pkg.presentation
        # verify -----------------------
        actual = isinstance(obj, cls)
        msg = ("expected instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)

    def test_it_should_have_core_props(self):
        """_Package should provide access to core document properties"""
        # setup ------------------------
        pkg = _Package()
        # verify -----------------------
        assert_that(pkg.core_properties, is_(instance_of(CoreProperties)))

    def test_saved_file_has_plausible_contents(self):
        """_Package.save produces a .pptx with plausible contents"""
        # setup ------------------------
        pkg = _Package()
        # exercise ---------------------
        pkg.save(self.test_pptx_path)
        # verify -----------------------
        pkg = _Package(self.test_pptx_path)
        prs = pkg.presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))


class Test_Part(TestCase):
    """Test _Part"""
    def test_constructs_presentation_for_rt_officedocument(self):
        """_Part() returns Presentation for RT.OFFICE_DOCUMENT"""
        # setup ------------------------
        cls = Presentation
        # exercise ---------------------
        obj = _Part(RT.OFFICE_DOCUMENT, CT.PML_PRESENTATION_MAIN)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slide_for_rt_slide(self):
        """_Part() returns _Slide for RT.SLIDE"""
        # setup ------------------------
        cls = _Slide
        # exercise ---------------------
        obj = _Part(RT.SLIDE, CT.PML_SLIDE)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slidelayout_for_rt_slidelayout(self):
        """_Part() returns _SlideLayout for RT.SLIDE_LAYOUT"""
        # setup ------------------------
        cls = _SlideLayout
        # exercise ---------------------
        obj = _Part(RT.SLIDE_LAYOUT, CT.PML_SLIDE_LAYOUT)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slidemaster_for_rt_slidemaster(self):
        """_Part() returns _SlideMaster for RT.SLIDE_MASTER"""
        # setup ------------------------
        cls = _SlideMaster
        # exercise ---------------------
        obj = _Part(RT.SLIDE_MASTER, CT.PML_SLIDE_MASTER)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_contructor_raises_on_invalid_prs_content_type(self):
        """_Part() raises on invalid presentation content type"""
        with self.assertRaises(InvalidPackageError):
            _Part(RT.OFFICE_DOCUMENT, CT.PML_SLIDE_MASTER)


class Test_Presentation(TestCase):
    """Test Presentation"""
    def setUp(self):
        self.prs = Presentation()

    def test__blob_rewrites_sldIdLst(self):
        """Presentation._blob rewrites sldIdLst"""
        # setup ------------------------
        rels = RelationshipCollectionBuilder()
        rels = rels.with_tuple_targets(2, RT.SLIDE_MASTER)
        rels = rels.with_tuple_targets(3, RT.SLIDE)
        rels = rels.with_ordering(RT.SLIDE_MASTER, RT.SLIDE)
        rels = rels.build()
        prs = Presentation()
        prs._relationships = rels
        prs.partname = '/ppt/presentation.xml'
        path = absjoin(test_file_dir, 'presentation.xml')
        prs._element = oxml_parse(path).getroot()
        # exercise ---------------------
        blob = prs._blob
        # verify -----------------------
        presentation = oxml_fromstring(blob)
        sldIds = presentation.xpath('./p:sldIdLst/p:sldId', namespaces=nsmap)
        expected = ['rId3', 'rId4', 'rId5']
        actual = [sldId.get(qtag('r:id')) for sldId in sldIds]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_slidemasters_property_empty_on_construction(self):
        """Presentation.slidemasters property empty on construction"""
        # verify -----------------------
        self.assertIsSizedProperty(self.prs, 'slidemasters', 0)

    def test_slidemasters_correct_length_after_pkg_open(self):
        """Presentation.slidemasters correct length after load"""
        # setup ------------------------
        pkg = _Package(test_pptx_path)
        prs = pkg.presentation
        # exercise ---------------------
        slidemasters = prs.slidemasters
        # verify -----------------------
        self.assertLength(slidemasters, 1)

    def test_slides_property_empty_on_construction(self):
        """Presentation.slides property empty on construction"""
        # verify -----------------------
        self.assertIsSizedProperty(self.prs, 'slides', 0)

    def test_slides_correct_length_after_pkg_open(self):
        """Presentation.slides correct length after load"""
        # setup ------------------------
        pkg = _Package(test_pptx_path)
        prs = pkg.presentation
        # exercise ---------------------
        slides = prs.slides
        # verify -----------------------
        self.assertLength(slides, 1)
