# -*- coding: utf-8 -*-
#
# test_presentation.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.presentation module."""

import gc
import inspect
import os

from hamcrest import (assert_that, has_item, is_, is_in, is_not, equal_to,
                      greater_than)
from lxml import etree

try:
    from PIL import Image as PILImage
except ImportError:
    import Image as PILImage

from mock import Mock, patch, PropertyMock

import pptx.presentation

from pptx.presentation import (Package, Collection, _RelationshipCollection,
    _Relationship, Presentation, PartCollection, BasePart, Part,
    SlideCollection, BaseSlide, Slide, SlideLayout, SlideMaster, Image,
    ShapeCollection, BaseShape, Shape, Placeholder, TextFrame, Paragraph, Run)

from pptx.exceptions import InvalidPackageError
from pptx.spec import namespaces, qname
from pptx.spec import (CT_PRESENTATION, CT_SLIDE, CT_SLIDELAYOUT,
    CT_SLIDEMASTER)
from pptx.spec import (RT_HANDOUTMASTER, RT_IMAGE, RT_NOTESMASTER,
    RT_OFFICEDOCUMENT, RT_PRESPROPS, RT_SLIDE, RT_SLIDELAYOUT, RT_SLIDEMASTER,
    RT_TABLESTYLES, RT_THEME, RT_VIEWPROPS)
from pptx.spec import (PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ,
    PH_TYPE_SLDNUM, PH_TYPE_SUBTITLE, PH_TYPE_TBL, PH_TYPE_TITLE,
    PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER)
from pptx.util import Px
from testing import TestCase

import logging
log = logging.getLogger('pptx.test.presentation')
log.setLevel(logging.DEBUG)
# log.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - '
                              '%(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)



# module globals -------------------------------------------------------------
def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir = os.path.split(__file__)[0]
test_file_dir = absjoin(thisdir, 'test_files')

test_image_path  = absjoin(test_file_dir, 'python-icon.jpeg')
new_image_path   = absjoin(test_file_dir, 'monty-truth.png')
test_pptx_path   = absjoin(test_file_dir, 'test.pptx')
images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

nsmap = namespaces('a', 'p', 'r')

def _empty_spTree():
    xml = ('<p:spTree xmlns:p="http://schemas.openxmlformats.org/'
           'presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.'
           'org/drawingml/2006/main"><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
           '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>')
    return etree.fromstring(xml)

def _sldLayout1():
    path = os.path.join(thisdir, 'test_files/slideLayout1.xml')
    sldLayout = etree.parse(path).getroot()
    return sldLayout

def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = ShapeCollection(spTree)
    return shapes


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
    partname_tmpls =\
        { RT_SLIDEMASTER : '/ppt/slideMasters/slideMaster%d.xml'
        , RT_SLIDE       : '/ppt/slides/slide%d.xml'
        }
    
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
            rId = self.__next_rId
            partname = self.__next_tuple_partname(reltype)
            target = PartBuilder().with_partname(partname).build()
            rel = _Relationship(rId, reltype, target)
            self.relationships.append(rel)
        return self
    
    # def with_singleton_target(self, reltype):
    #     rId = self.__next_rId
    #     partname = self.__singleton_partname(reltype)
    #     target = PartBuilder().with_partname(partname).build()
    #     rel = _Relationship(rId, reltype, target)
    #     self.relationships.append(rel)
    #     return self
    # 
    def __next_partnum(self, reltype):
        if reltype not in self.next_partnums:
            self.next_partnums[reltype] = 1
        partnum = self.next_partnums[reltype]
        self.next_partnums[reltype] = partnum + 1
        return partnum
    
    @property
    def __next_rId(self):
        rId = 'rId%d' % self.next_rel_num
        self.next_rel_num += 1
        return rId
    
    def __next_tuple_partname(self, reltype):
        partname_tmpl = self.partname_tmpls[reltype]
        partnum = self.__next_partnum(reltype)
        return partname_tmpl % partnum
    
    def build(self):
        rels = _RelationshipCollection()
        for rel in self.relationships:
            rels._additem(rel)
        if self.reltype_ordering:
            rels._reltype_ordering = self.reltype_ordering
        return rels
    

class TestBasePart(TestCase):
    """Test BasePart"""
    def setUp(self):
        self.basepart = BasePart()
        self.cls = BasePart
    
    def test__add_relationship_adds_specified_relationship(self):
        """BasePart._add_relationship adds specified relationship"""
        # setup -----------------------
        reltype = RT_IMAGE
        target = Mock(name='image')
        # exercise --------------------
        rel = self.basepart._add_relationship(reltype, target)
        # verify ----------------------
        expected = ('rId1', reltype, target)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__blob_value_for_binary_part(self):
        """BasePart._blob value is correct for binary part"""
        # setup -----------------------
        blob = '0123456789'
        self.basepart._load_blob = blob
        self.basepart.partname = '/docProps/thumbnail.jpeg'
        # exercise --------------------
        retval = self.basepart._blob
        # verify ----------------------
        expected = blob
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__blob_value_for_xml_part(self):
        """BasePart._blob value is correct for XML part"""
        # setup -----------------------
        elm = etree.fromstring('<root><elm1 attr="one"/></root>')
        self.basepart._element = elm
        self.basepart.partname = '/ppt/presentation.xml'
        # exercise --------------------
        retval = self.basepart._blob
        # verify ----------------------
        expected = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"\
                   '\n<root>\n  <elm1 attr="one"/>\n</root>\n'
        actual = retval
        msg = "expected: \n'%s'\n, got \n'%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__load_sets__element_for_xml_part(self):
        """BasePart._load() sets _element for xml part"""
        # setup -----------------------
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/presentation.xml'
        pkgpart.blob = '<root><elm1   attr="spam"/></root>'
        pkgpart.relationships = []
        part_dict = {}
        part = self.basepart._load(pkgpart, part_dict)
        # exercise --------------------
        elm = part._element
        # verify ----------------------
        expected = '<root><elm1 attr="spam"/></root>'
        actual = etree.tostring(elm)
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_observable_on_partname(self):
        """BasePart observable on partname value change"""
        # setup -----------------------
        old_partname = '/ppt/slides/slide1.xml'
        new_partname = '/ppt/slides/slide2.xml'
        observer = Mock()
        self.basepart.partname = old_partname
        self.basepart.add_observer(observer)
        # exercise --------------------
        self.basepart.partname = new_partname
        # verify ----------------------
        observer.notify.assert_called_with(self.basepart, 'partname',
                                           new_partname)
    
    def test_partname_setter(self):
        """BasePart.partname setter stores passed value"""
        # setup -----------------------
        partname = '/ppt/presentation.xml'
        # exercise ----------------
        self.basepart.partname = partname
        # verify ------------------
        expected = partname
        actual = self.basepart.partname
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestBaseShape(TestCase):
    """Test BaseShape"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = etree.parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:pic'
        pic = self.sld.xpath(xpath, namespaces=nsmap)[0]
        self.base_shape = BaseShape(pic)
    
    def test_class_present(self):
        """BaseShape class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'BaseShape')
    
    def test_has_textframe_value(self):
        """BaseShape.has_textframe value correct"""
        # setup -----------------------
        spTree = self.sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        shapes = ShapeCollection(spTree)
        indexes = []
        # exercise --------------------
        for idx, shape in enumerate(shapes):
            if shape.has_textframe:
                indexes.append(idx)
        # verify ----------------------
        expected = [0, 1, 3, 5, 6]
        actual = indexes
        msg = "expected txBody element in shapes %s, got %s"\
              % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_id_value(self):
        """BaseShape.id value is correct"""
        # exercise --------------------
        id = self.base_shape.id
        # verify ----------------------
        expected = 6
        actual = id
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_is_placeholder_true_for_placeholder(self):
        """BaseShape.is_placeholder True for placeholder shape"""
        # setup -----------------------
        xpath = './p:cSld/p:spTree/p:sp'
        sp = self.sld.xpath(xpath, namespaces=nsmap)[0]
        base_shape = BaseShape(sp)
        # verify ----------------------
        actual = base_shape.is_placeholder
        msg = "expected True, got %s" % (actual)
        self.assertTrue(actual, msg)
    
    def test_is_placeholder_false_for_non_placeholder(self):
        """BaseShape.is_placeholder False for non-placeholder shape"""
        # verify ----------------------
        actual = self.base_shape.is_placeholder
        msg = "expected False, got %s" % (actual)
        self.assertFalse(actual, msg)
    
    def test__is_title_true_for_title_placeholder(self):
        """BaseShape._is_title True for title placeholder shape"""
        # setup -----------------------
        xpath = './p:cSld/p:spTree/p:sp'
        title_placeholder_sp = self.sld.xpath(xpath, namespaces=nsmap)[0]
        base_shape = BaseShape(title_placeholder_sp)
        # verify ----------------------
        actual = base_shape._is_title
        msg = "expected True, got %s" % (actual)
        self.assertTrue(actual, msg)
    
    def test_name_value(self):
        """BaseShape.name value is correct"""
        # exercise --------------------
        name = self.base_shape.name
        # verify ----------------------
        expected = 'Picture 5'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_textframe_raises_on_no_textframe(self):
        """BaseShape.textframe raises on shape with no text frame"""
        with self.assertRaises(ValueError):
            self.base_shape.textframe
    
    def test_text_setter_structure_and_value(self):
        """assign to BaseShape.text yields single run para set to value"""
        # setup -----------------------
        test_text = 'python-pptx was here!!'
        xpath = './p:cSld/p:spTree/p:sp'
        textbox_sp = self.sld.xpath(xpath, namespaces=nsmap)[2]
        base_shape = BaseShape(textbox_sp)
        # exercise --------------------
        base_shape.text = test_text
        # verify paragraph count ------
        expected = 1
        actual = len(base_shape.textframe.paragraphs)
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value ----------------
        expected = test_text
        actual = base_shape.textframe.paragraphs[0].runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_text_setter_raises_on_no_textframe(self):
        """assignment to BaseShape.text raises for shape with no text frame"""
        with self.assertRaises(TypeError):
            self.base_shape.text = 'test text'
    

class TestBaseSlide(TestCase):
    """Test BaseSlide"""
    def setUp(self):
        self.base_slide = BaseSlide()
    
    def test_name_value(self):
        """BaseSlide.name value is correct"""
        # setup -----------------------
        self.base_slide._element = _sldLayout1()
        # exercise --------------------
        name = self.base_slide.name
        # verify ----------------------
        expected = 'Title Slide'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_shapes_size_after__load(self):
        """BaseSlide.shapes is expected size after _load()"""
        # setup -----------------------
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/slides/slide1.xml'
        with open(path, 'r') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = []
        part_dict = {}
        self.base_slide._load(pkgpart, part_dict)
        # exercise --------------------
        shapes = self.base_slide.shapes
        # verify ----------------------
        self.assertLength(shapes, 9)
    

class TestCollection(TestCase):
    """Test Collection"""
    def setUp(self):
        self.collection = Collection()
    
    def test_class_present(self):
        """Collection class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Collection')
    
    def test_indexable(self):
        """Collection is indexable (e.g. no TypeError on 'collection[0]')"""
        # verify ----------------------
        try:
            self.collection[0]
        except TypeError:
            msg = "'Collection' object does not support indexing"
            self.fail(msg)
        except IndexError:
            pass
    
    def test_is_container(self):
        """Collection is container (e.g. 'x in collection' works)"""
        # verify ----------------------
        try:
            1 in self.collection
        except TypeError:
            msg = "'Collection' object is not container"
            self.fail(msg)
    
    def test_iterable(self):
        """Collection is iterable"""
        # verify ----------------------
        try:
            for x in self.collection:
                pass
        except TypeError:
            msg = "'Collection' object is not iterable"
            self.fail(msg)
    
    def test_sized(self):
        """Collection is sized (e.g. 'len(collection)' works)"""
        # verify ----------------------
        try:
            len(self.collection)
        except TypeError:
            msg = "object of type 'Collection' has no len()"
            self.fail(msg)
    
    def test__values_property_empty_on_construction(self):
        """Collection._values property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.collection, '_values', 0)
    

class TestImage(TestCase):
    """Test Image"""
    def test_construction_from_file(self):
        """Image(path) constructor produces correct attribute values"""
        # exercise --------------------
        image = Image(test_image_path)
        # verify ----------------------
        expected = '.jpeg'
        actual = image.ext
        msg = "expected extension '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        
        expected = 'image/jpeg'
        actual = image._content_type
        msg = "expected content_type '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        
        expected = 3277
        actual = len(image._blob)
        msg = "expected blob size %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_construction_from_file_raises_on_bad_path(self):
        """Image(path) constructor raises on bad path"""
        # verify ----------------------
        with self.assertRaises(IOError):
            image = Image('foobar27.png')
    
    def test___image_file_content_type_known_type(self):
        """Image.__image_file_content_type() correct for known content type"""
        # exercise --------------------
        content_type = Image._Image__image_file_content_type(test_image_path)
        # verify ----------------------
        expected = 'image/jpeg'
        actual = content_type
        msg = ("expected content type '%s', got '%s'" % (expected, actual))
        self.assertEqual(expected, actual, msg)
    
    def test___image_file_content_type_raises_on_bad_ext(self):
        """Image.__image_file_content_type() raises on bad extension"""
        # setup -----------------------
        path = 'image.xj7'
        # verify ----------------------
        with self.assertRaises(TypeError):
            Image._Image__image_file_content_type(path)
    
    def test___image_file_content_type_raises_on_non_img_ext(self):
        """Image.__image_file_content_type() raises on non-image extension"""
        # setup -----------------------
        path = 'image.xml'
        # verify ----------------------
        with self.assertRaises(TypeError):
            Image._Image__image_file_content_type(path)
    

class TestImageCollection(TestCase):
    """Test ImageCollection"""
    def test_add_image_returns_matching_image(self):
        """ImageCollection.add_image() returns existing image on match"""
        # setup -----------------------
        pkg = Package(images_pptx_path)
        matching_idx = 4
        matching_image = pkg._images[matching_idx]
        # exercise --------------------
        image = pkg._images.add_image(test_image_path)
        # verify ----------------------
        expected = matching_image
        actual = image
        msg = ("expected images[%d], got images[%d]"
               % (matching_idx, pkg._images.index(image)))
        self.assertEqual(expected, actual, msg)
    
    def test_add_image_adds_new_image(self):
        """ImageCollection.add_image() adds new image on no match"""
        # setup -----------------------
        pkg = Package(images_pptx_path)
        expected_partname = '/ppt/media/image8.png'
        expected_len = len(pkg._images) + 1
        expected_sha1 = '79769f1e202add2e963158b532e36c2c0f76a70c'
        # exercise --------------------
        image = pkg._images.add_image(new_image_path)
        # verify ----------------------
        expected = (expected_partname, expected_len, expected_sha1)
        actual = (image.partname, len(pkg._images), image._sha1)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestPackage(TestCase):
    """Test Package"""
    def test_construction_with_no_path_loads_default_template(self):
        """Package() call with no path loads default template"""
        prs = Package().presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))
    
    def test_instances_are_tracked(self):
        """Package instances are tracked"""
        pkg = Package()
        self.assertIn(pkg, Package.instances())
    
    def test_instance_refs_are_garbage_collected(self):
        """Package instances are tracked"""
        pkg = Package()
        pkg1_repr = "%r" % pkg
        pkg = Package()
        # pkg2_repr = "%r" % pkg
        gc.collect()
        reprs = [repr(pkg_inst) for pkg_inst in Package.instances()]
        # log.debug("pkg1, pkg2, reprs: %s, %s, %s"
        #           % (pkg1_repr, pkg2_repr, reprs))
        assert_that(pkg1_repr, is_not(is_in(reprs)))
    
    def test_containing_returns_correct_pkg(self):
        """Package.containing() returns right package instance"""
        # setup -----------------------
        pkg1 = Package(test_pptx_path)
        pkg2 = Package(test_pptx_path)
        slide = pkg2.presentation.slides[0]
        # exercise --------------------
        found_pkg = Package.containing(slide)
        # verify ----------------------
        expected = pkg2
        actual = found_pkg
        msg = "expected %r, got %r" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_open_gathers_image_parts(self):
        """Package open gathers image parts into image collection"""
        # exercise --------------------
        pkg = Package(images_pptx_path)
        # verify ----------------------
        expected = 7
        actual = len(pkg._Package__images)
        msg = "expected image count of %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_presentation_presentation_after_open(self):
        """Package.presentation is instance of Presentation after open()"""
        # setup -----------------------
        cls = Presentation
        pkg = Package()
        # exercise --------------------
        obj = pkg.presentation
        # verify ----------------------
        actual = isinstance(obj, cls)
        msg = ("expected instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)
    
    def test_saved_file_has_plausible_contents(self):
        """Package.save produces a .pptx with plausible contents"""
        # setup -----------------------
        test_pptx_path = absjoin(test_file_dir, 'test_python-pptx.pptx')
        if os.path.isfile(test_pptx_path):
            os.remove(test_pptx_path)
        pkg = Package()
        # exercise --------------------
        pkg.save(test_pptx_path)
        # verify ----------------------
        pkg = Package(test_pptx_path)
        prs = pkg.presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))
    

class TestParagraph(TestCase):
    """Test Paragraph"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = etree.parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody/a:p'
        self.pList = self.sld.xpath(xpath, namespaces=nsmap)
    
    def test_runs_size(self):
        """Paragraph.runs is expected size"""
        # setup -----------------------
        actual_lengths = []
        for p in self.pList:
            paragraph = Paragraph(p)
            # exercise ----------------
            actual_lengths.append(len(paragraph.runs))
        # verify ------------------
        expected = [0, 0, 2, 1, 1, 1]
        actual = actual_lengths
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_add_run_increments_run_count(self):
        """Paragraph.add_run() increments run count"""
        # setup -----------------------
        p_elm = self.pList[0]
        paragraph = Paragraph(p_elm)
        # exercise --------------------
        run = paragraph.add_run()
        # verify ----------------------
        expected = 1
        actual = len(paragraph.runs)
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_clear_removes_all_runs(self):
        """Paragraph.clear() removes all runs from paragraph"""
        # setup -----------------------
        p_elm = self.pList[2]
        paragraph = Paragraph(p_elm)
        expected = 2
        actual = len(paragraph.runs)
        msg = "expected pre-test run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # exercise --------------------
        paragraph.clear()
        # verify ----------------------
        expected = 0
        actual = len(paragraph.runs)
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_text_setter_sets_single_run_text(self):
        """assignment to Paragraph.text creates single run containing value"""
        # setup -----------------------
        test_text = 'python-pptx was here!!'
        p_elm = self.pList[2]
        paragraph = Paragraph(p_elm)
        # exercise --------------------
        paragraph.text = test_text
        # verify run count ------------
        expected = 1
        actual = len(paragraph.runs)
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value ----------------
        expected = test_text
        actual = paragraph.runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestPart(TestCase):
    """Test Part"""
    def test_class_present(self):
        """Part class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Part')
    
    def test_constructs_presentation_for_rt_officedocument(self):
        """Part() returns Presentation for RT_OFFICEDOCUMENT"""
        # setup -----------------------
        cls = Presentation
        # exercise --------------------
        obj = Part(RT_OFFICEDOCUMENT, CT_PRESENTATION)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_constructs_slide_for_rt_slide(self):
        """Part() returns Slide for RT_SLIDE"""
        # setup -----------------------
        cls = Slide
        # exercise --------------------
        obj = Part(RT_SLIDE, CT_SLIDE)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_constructs_slidelayout_for_rt_slidelayout(self):
        """Part() returns SlideLayout for RT_SLIDELAYOUT"""
        # setup -----------------------
        cls = SlideLayout
        # exercise --------------------
        obj = Part(RT_SLIDELAYOUT, CT_SLIDELAYOUT)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_constructs_slidemaster_for_rt_slidemaster(self):
        """Part() returns SlideMaster for RT_SLIDEMASTER"""
        # setup -----------------------
        cls = SlideMaster
        # exercise --------------------
        obj = Part(RT_SLIDEMASTER, CT_SLIDEMASTER)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_contructor_raises_on_invalid_prs_content_type(self):
        """Part() raises on invalid presentation content type"""
        with self.assertRaises(InvalidPackageError):
            Part(RT_OFFICEDOCUMENT, CT_SLIDEMASTER)
    

class TestPartCollection(TestCase):
    """Test PartCollection"""
    def test__loadpart_sorts_loaded_parts(self):
        """PartCollection._loadpart sorts loaded parts"""
        # setup -----------------------
        partname1 = '/ppt/slides/slide1.xml'
        partname2 = '/ppt/slides/slide2.xml'
        partname3 = '/ppt/slides/slide3.xml'
        part1 = Mock(name='part1'); part1.partname = partname1
        part2 = Mock(name='part2'); part2.partname = partname2
        part3 = Mock(name='part3'); part3.partname = partname3
        parts = PartCollection()
        # exercise --------------------
        parts._loadpart(part2)
        parts._loadpart(part3)
        parts._loadpart(part1)
        # verify ----------------------
        expected = [partname1, partname2, partname3]
        actual = [part.partname for part in parts]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestPlaceholder(TestCase):
    """Test Placeholder"""
    def test_property_values(self):
        """Placeholder property values are correct"""
        # setup -----------------------
        expected_values =\
            ( (PH_TYPE_CTRTITLE, PH_ORIENT_HORZ, PH_SZ_FULL,     0)
            , (PH_TYPE_DT,       PH_ORIENT_HORZ, PH_SZ_HALF,    10)
            , (PH_TYPE_SUBTITLE, PH_ORIENT_VERT, PH_SZ_FULL,     1)
            , (PH_TYPE_TBL,      PH_ORIENT_HORZ, PH_SZ_QUARTER, 14)
            , (PH_TYPE_SLDNUM,   PH_ORIENT_HORZ, PH_SZ_QUARTER, 12)
            , (PH_TYPE_FTR,      PH_ORIENT_HORZ, PH_SZ_QUARTER, 11)
            )
        shapes = _sldLayout1_shapes()
        # exercise --------------------
        for idx, sp in enumerate(shapes):
            ph = Placeholder(sp)
            values = (ph.type, ph.orient, ph.sz, ph.idx)
            # verify ----------------------
            expected = expected_values[idx]
            actual = values
            msg = "expected shapes[%d] values %s, got %s"\
                   % (idx, expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestPresentation(TestCase):
    """Test Presentation"""
    def setUp(self):
        self.prs = Presentation()
    
    def test__blob_rewrites_sldIdLst(self):
        """Presentation._blob rewrites sldIdLst"""
        # setup -----------------------
        relationships = RelationshipCollectionBuilder()\
                       .with_tuple_targets(2, RT_SLIDEMASTER)\
                       .with_tuple_targets(3, RT_SLIDE)\
                       .with_ordering(RT_SLIDEMASTER, RT_SLIDE)\
                       .build()
        prs = Presentation()
        prs._relationships = relationships
        prs.partname = '/ppt/presentation.xml'
        path = os.path.join(thisdir, 'test_files/presentation.xml')
        prs._element = etree.parse(path).getroot()
        # exercise --------------------
        blob = prs._blob
        # verify ----------------------
        presentation = etree.fromstring(blob)
        sldIds = presentation.xpath('./p:sldIdLst/p:sldId', namespaces=nsmap)
        expected = ['rId3', 'rId4', 'rId5']
        actual = [sldId.get(qname('r', 'id')) for sldId in sldIds]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_slidemasters_property_empty_on_construction(self):
        """Presentation.slidemasters property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.prs, 'slidemasters', 0)
    
    def test_slidemasters_correct_length_after_pkg_open(self):
        """Presentation.slidemasters correct length after load"""
        # setup -----------------------
        pkg = Package(test_pptx_path)
        prs = pkg.presentation
        # exercise --------------------
        slidemasters = prs.slidemasters
        # verify ----------------------
        self.assertLength(slidemasters, 1)
    
    def test_slides_property_empty_on_construction(self):
        """Presentation.slides property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.prs, 'slides', 0)
    
    def test_slides_correct_length_after_pkg_open(self):
        """Presentation.slides correct length after load"""
        # setup -----------------------
        pkg = Package(test_pptx_path)
        prs = pkg.presentation
        # exercise --------------------
        slides = prs.slides
        # verify ----------------------
        self.assertLength(slides, 1)
    

class Test_Relationship(TestCase):
    """Test _Relationship"""
    def setUp(self):
        rId = 'rId1'
        reltype = RT_SLIDE
        target_part = None
        self.rel = _Relationship(rId, reltype, target_part)
    
    def test_constructor_raises_on_bad_rId(self):
        """_Relationship constructor raises on non-standard rId"""
        with self.assertRaises(AssertionError):
            _Relationship('Non-std14', None, None)
    
    def test__num_value(self):
        """_Relationship._num value is correct"""
        # setup -----------------------
        num = 91
        rId = 'rId%d' % num
        rel = _Relationship(rId, None, None)
        # verify ----------------------
        expected = num
        actual = rel._num
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__rId_setter(self):
        """Relationship._rId setter stores passed value"""
        # setup -----------------------
        rId = 'rId9'
        # exercise ----------------
        self.rel._rId = rId
        # verify ------------------
        expected = rId
        actual = self.rel._rId
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class Test_RelationshipCollection(TestCase):
    """Test _RelationshipCollection"""
    def setUp(self):
        self.relationships = _RelationshipCollection()
    
    def test__additem_raises_on_dup_rId(self):
        """_RelationshipCollection._additem raises on duplicate rId"""
        # setup -----------------------
        part1 = BasePart()
        part2 = BasePart()
        rel1 = _Relationship('rId9', None, part1)
        rel2 = _Relationship('rId9', None, part2)
        self.relationships._additem(rel1)
        # verify ----------------------
        with self.assertRaises(ValueError):
            self.relationships._additem(rel2)
    
    def test__additem_maintains_rId_ordering(self):
        """_RelationshipCollection maintains rId ordering on additem()"""
        # setup -----------------------
        part1 = BasePart()
        part2 = BasePart()
        part3 = BasePart()
        rel1 = _Relationship('rId1', None, part1)
        rel2 = _Relationship('rId2', None, part2)
        rel3 = _Relationship('rId3', None, part3)
        # exercise --------------------
        self.relationships._additem(rel2)
        self.relationships._additem(rel1)
        self.relationships._additem(rel3)
        # verify ----------------------
        expected = ['rId1', 'rId2', 'rId3']
        actual = [rel._rId for rel in self.relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def __reltype_ordering_mock(self):
        """
        Return RelationshipCollection instance with mocked-up contents
        suitable for testing _reltype_ordering.
        """
        # setup -----------------------
        partnames =\
            [ '/ppt/slides/slide4.xml'
            , '/ppt/slideLayouts/slideLayout1.xml'
            , '/ppt/slideMasters/slideMaster1.xml'
            , '/ppt/slides/slide1.xml'
            ]
        part1 = Mock(name='part1'); part1.partname = partnames[0]
        part2 = Mock(name='part2'); part2.partname = partnames[1]
        part3 = Mock(name='part3'); part3.partname = partnames[2]
        part4 = Mock(name='part4'); part4.partname = partnames[3]
        rel1 = _Relationship('rId1', RT_SLIDE,       part1)
        rel2 = _Relationship('rId2', RT_SLIDELAYOUT, part2)
        rel3 = _Relationship('rId3', RT_SLIDEMASTER, part3)
        rel4 = _Relationship('rId4', RT_SLIDE,       part4)
        relationships = _RelationshipCollection()
        relationships._additem(rel1)
        relationships._additem(rel2)
        relationships._additem(rel3)
        relationships._additem(rel4)
        return (relationships, partnames)
    
    def test__additem_maintains_reltype_ordering(self):
        """_RelationshipCollection maintains reltype ordering on additem()"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        ordering = (RT_SLIDEMASTER, RT_SLIDELAYOUT, RT_SLIDE)
        relationships._reltype_ordering = ordering
        partname = '/ppt/slides/slide2.xml'
        part = Mock(name='new_part'); part.partname = partname
        rId = relationships._next_rId
        rel = _Relationship(rId, RT_SLIDE, part)
        # exercise --------------------
        relationships._additem(rel)
        # verify ordering -------------
        expected = [partnames[2], partnames[1], partnames[3],
                    partname, partnames[0]]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_rels_of_reltype_return_value(self):
        """RelationshipCollection._rels_of_reltype returns correct rels"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        # exercise --------------------
        retval = relationships.rels_of_reltype(RT_SLIDE)
        # verify ordering -------------
        expected = ['rId1', 'rId4']
        actual = [rel._rId for rel in retval]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__reltype_ordering_sorts_rels(self):
        """RelationshipCollection._reltype_ordering sorts rels"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        ordering = (RT_SLIDEMASTER, RT_SLIDELAYOUT, RT_SLIDE)
        # exercise --------------------
        relationships._reltype_ordering = ordering
        # verify ordering -------------
        expected = [partnames[2], partnames[1], partnames[3], partnames[0]]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__reltype_ordering_renumbers_rels(self):
        """RelationshipCollection._reltype_ordering renumbers rels"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        ordering = (RT_SLIDEMASTER, RT_SLIDELAYOUT, RT_SLIDE)
        # exercise --------------------
        relationships._reltype_ordering = ordering
        # verify renumbering ----------
        expected = ['rId1', 'rId2', 'rId3', 'rId4']
        actual = [rel._rId for rel in relationships]
        msg = "expected numbering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__next_rId_fills_gap(self):
        """_RelationshipCollection._next_rId fills gap in rId sequence"""
        # setup -----------------------
        part1 = BasePart()
        part2 = BasePart()
        part3 = BasePart()
        part4 = BasePart()
        rel1 = _Relationship('rId1', None, part1)
        rel2 = _Relationship('rId2', None, part2)
        rel3 = _Relationship('rId3', None, part3)
        rel4 = _Relationship('rId4', None, part4)
        cases =\
            ( ('rId1', (rel2, rel3, rel4))
            , ('rId2', (rel1, rel3, rel4))
            , ('rId3', (rel1, rel2, rel4))
            , ('rId4', (rel1, rel2, rel3))
            )
        # exercise --------------------
        expected_rIds = []
        actual_rIds = []
        for expected_rId, rels in cases:
            expected_rIds.append(expected_rId)
            relationships = _RelationshipCollection()
            for rel in rels:
                relationships._additem(rel)
            actual_rIds.append(relationships._next_rId)
        # verify ----------------------
        expected = expected_rIds
        actual = actual_rIds
        msg = "expected rIds %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_reorders_on_partname_change(self):
        """RelationshipCollection reorders on partname change"""
        # setup -----------------------
        partname1 = '/ppt/slides/slide1.xml'
        partname2 = '/ppt/slides/slide2.xml'
        partname3 = '/ppt/slides/slide3.xml'
        part1 = PartBuilder().with_partname(partname1).build()
        part2 = PartBuilder().with_partname(partname2).build()
        rel1 = _Relationship('rId1', RT_SLIDE, part1)
        rel2 = _Relationship('rId2', RT_SLIDE, part2)
        relationships = _RelationshipCollection()
        relationships._reltype_ordering = (RT_SLIDE)
        relationships._additem(rel1)
        relationships._additem(rel2)
        # exercise --------------------
        part1.partname = partname3
        # verify ----------------------
        expected = [partname2, partname3]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestRun(TestCase):
    """Test Run"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = etree.parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r'
        self.rList = self.sld.xpath(xpath, namespaces=nsmap)
    
    def test_class_present(self):
        """Run class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Run')
    
    def test_text_value(self):
        """Run.text value is correct"""
        # setup -----------------------
        run = Run(self.rList[1])
        # exercise ----------------
        text = run.text
        # verify ------------------
        expected = ' 2nd run'
        actual = text
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_text_setter(self):
        """Run.text setter stores passed value"""
        # setup -----------------------
        new_value = 'new string'
        run = Run(self.rList[1])
        # exercise ----------------
        run.text = new_value
        # verify ------------------
        expected = new_value
        actual = run.text
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestShape(TestCase):
    """Test Shape"""
    def __loaded_shape(self):
        """
        Return Shape instance loaded from test file.
        """
        sldLayout = _slideLayout1()
        sp = sldLayout.xpath('p:cSld/p:spTree/p:sp', namespaces=nsmap)[0]
        return Shape(sp)
    
    def test_class_present(self):
        """Shape class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Shape')
    

class TestShapeCollection(TestCase):
    """Test ShapeCollection"""
    def setUp(self):
        path = absjoin(test_file_dir, 'slide1.xml')
        sld = etree.parse(path).getroot()
        spTree = sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        self.shapes = ShapeCollection(spTree)
    
    def test_construction_size(self):
        """ShapeCollection is expected size after construction"""
        # verify ----------------------
        self.assertLength(self.shapes, 9)
    
    @patch('pptx.presentation.Collection._values', new_callable=PropertyMock)
    @patch('pptx.presentation.Package')
    def test_add_picture_collaboration(self, MockPackage, mock_values):
        """ShapeCollection.add_picture() calls the right collaborators"""
        # constant values -------------
        rId = 'rId1'
        top = 1
        left = 2
        # setup mockery ---------------
        pkg      = Mock(name='pkg')
        image    = Mock(name='image')
        rel      = Mock(name='rel')
        pic      = Mock(name='pic')
        slide    = Mock(name='slide')
        __pic    = Mock(name='__pic')
        __spTree = Mock(name='__spTree')
        Picture  = Mock(name='Picture')
        MockPackage.containing.return_value = pkg
        pkg._images.add_image.return_value = image
        slide._add_relationship.return_value = rel
        rel._rId = rId
        __pic.return_value = pic
        pptx.presentation.Picture = Picture
        # setup -----------------------
        shapes = ShapeCollection(_empty_spTree(), slide)
        shapes._ShapeCollection__pic = __pic
        shapes._ShapeCollection__spTree = __spTree
        # exercise --------------------
        picture = shapes.add_picture(test_image_path, top, left)
        # verify ----------------------
        MockPackage.containing.assert_called_once_with(slide)
        pkg._images.add_image.assert_called_once_with(test_image_path)
        slide._add_relationship.assert_called_once_with(RT_IMAGE, image)
        __pic.assert_called_once_with(rId, test_image_path, top, left)
        __spTree.append.assert_called_once_with(pic)
        Picture.assert_called_once_with(pic)
        shapes._values.append.assert_called_once_with(picture)
    
    def test_title_value(self):
        """ShapeCollection.title value is ref to correct shape"""
        # exercise --------------------
        title_shape = self.shapes.title
        # verify ----------------------
        expected = 0
        actual = self.shapes.index(title_shape)
        msg = "expected shapes[%d], got shapes[%d]" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_placeholders_values(self):
        """ShapeCollection.placeholders values are correct and sorted"""
        # setup -----------------------
        expected_values =\
            ( ('Title 1',                    PH_TYPE_CTRTITLE,  0)
            , ('Vertical Subtitle 2',        PH_TYPE_SUBTITLE,  1)
            , ('Date Placeholder 7',         PH_TYPE_DT,       10)
            , ('Footer Placeholder 4',       PH_TYPE_FTR,      11)
            , ('Slide Number Placeholder 5', PH_TYPE_SLDNUM,   12)
            , ('Table Placeholder 3',        PH_TYPE_TBL,      14)
            )
        shapes = _sldLayout1_shapes()
        # exercise --------------------
        placeholders = shapes.placeholders
        # verify ----------------------
        for idx, ph in enumerate(placeholders):
            values = (ph.name, ph.type, ph.idx)
            expected = expected_values[idx]
            actual = values
            msg = "expected placeholders[%d] values %s, got %s"\
                   % (idx, expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test__clone_layout_placeholders_shapes(self):
        """ShapeCollection._clone_layout_placeholders clones shapes"""
        # setup -----------------------
        expected_values =\
            ( [2, 'Title 1',                    PH_TYPE_CTRTITLE,  0]
            , [3, 'Vertical Subtitle 2',        PH_TYPE_SUBTITLE,  1]
            , [4, 'Table Placeholder 3',        PH_TYPE_TBL,      14]
            )
        slidelayout = SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        shapes = ShapeCollection(_empty_spTree())
        # exercise --------------------
        shapes._clone_layout_placeholders(slidelayout)
        # verify ----------------------
        for idx, sp in enumerate(shapes):
            # verify is placeholder ---
            is_placeholder = sp.is_placeholder
            msg = ("expected shapes[%d].is_placeholder == True %r"
                   % (idx, sp))
            self.assertTrue(is_placeholder, msg)
            # verify values -----------
            ph = Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = ("expected placeholder[%d] values %s, got %s"
                   % (idx, expected, actual))
            self.assertEqual(expected, actual, msg)
    
    def test___clone_layout_placeholder_values(self):
        """ShapeCollection.__clone_layout_placeholder() values correct"""
        # setup -----------------------
        layout_shapes = _sldLayout1_shapes()
        layout_ph_shapes = [sp for sp in layout_shapes if sp.is_placeholder]
        shapes = ShapeCollection(_empty_spTree())
        expected_values =\
            ( [2, 'Title 1',                    PH_TYPE_CTRTITLE,  0]
            , [3, 'Date Placeholder 2',         PH_TYPE_DT,       10]
            , [4, 'Vertical Subtitle 3',        PH_TYPE_SUBTITLE,  1]
            , [5, 'Table Placeholder 4',        PH_TYPE_TBL,      14]
            , [6, 'Slide Number Placeholder 5', PH_TYPE_SLDNUM,   12]
            , [7, 'Footer Placeholder 6',       PH_TYPE_FTR,      11]
            )
        # exercise --------------------
        for idx, layout_ph_sp in enumerate(layout_ph_shapes):
            layout_ph = Placeholder(layout_ph_sp)
            sp = shapes._ShapeCollection__clone_layout_placeholder(layout_ph)
            # verify ----------------------
            ph = Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = "expected placeholder values %s, got %s" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test___next_ph_name_return_value(self):
        """
        ShapeCollection.__next_ph_name() returns correct value
        
        * basename + 'Placeholder' + num, e.g. 'Table Placeholder 8'
        * numpart of name defaults to id-1, but increments until unique
        * prefix 'Vertical' if orient="vert"
        
        """
        cases =\
            ( (PH_TYPE_OBJ,   3, PH_ORIENT_HORZ, 'Content Placeholder 2')
            , (PH_TYPE_TBL,   4, PH_ORIENT_HORZ, 'Table Placeholder 4')
            , (PH_TYPE_TBL,   7, PH_ORIENT_VERT, 'Vertical Table Placeholder 6')
            , (PH_TYPE_TITLE, 2, PH_ORIENT_HORZ, 'Title 2')
            )
        # setup -----------------------
        shapes = _sldLayout1_shapes()
        for ph_type, id, orient, expected_name in cases:
            # exercise --------------------
            name = shapes._ShapeCollection__next_ph_name(ph_type, id, orient)
            # verify ----------------------
            expected = expected_name
            actual = name
            msg = "expected placeholder name '%s', got '%s'"\
                  % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test___next_shape_id_value(self):
        """ShapeCollection.__next_shape_id value is correct"""
        # setup -----------------------
        shapes = _sldLayout1_shapes()
        # exercise --------------------
        next_id = shapes._ShapeCollection__next_shape_id
        # verify ----------------------
        expected = 4
        actual = next_id
        msg = "expected %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test___pic_return_value(self):
        """ShapeCollection.__pic returns correct value"""
        # setup -----------------------
        test_image = PILImage.open(test_image_path)
        pic_size = tuple(Px(x) for x in test_image.size)
        xml = ('<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2'
            '006/main" xmlns:p="http://schemas.openxmlformats.org/presentatio'
            'nml/2006/main" xmlns:r="http://schemas.openxmlformats.org/office'
            'Document/2006/relationships"><p:nvPicPr><p:cNvPr id="4" name="Pi'
            'cture 3" descr="python-icon.jpeg"/><p:cNvPicPr/><p:nvPr/></p:nvP'
            'icPr><p:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/'
            '></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a'
            ':ext cx="%s" cy="%s"/></a:xfrm><a:prstGeom prst="rect"'
            '><a:avLst/></a:prstGeom></p:spPr></p:pic>' % pic_size)
        # exercise --------------------
        pic = self.shapes._ShapeCollection__pic('rId9', test_image_path, 0, 0)
        # verify ----------------------
        expected = xml
        actual = etree.tostring(pic)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestSlide(TestCase):
    """Test Slide"""
    def setUp(self):
        self.sld = Slide()
    
    def test_class_present(self):
        """Slide class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Slide')
    
    def test_constructor_sets_correct_content_type(self):
        """Slide constructor sets correct content type"""
        # exercise --------------------
        content_type = self.sld._content_type
        # verify ----------------------
        expected = CT_SLIDE
        actual = content_type
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_construction_adds_slide_layout_relationship(self):
        """Slide(slidelayout) adds relationship slide->slidelayout"""
        # setup -----------------------
        slidelayout = SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        # exercise --------------------
        slide = Slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(slide._relationships)
        msg = ("expected len(slide._relationships) of %d, got %d"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = slide._relationships[0]
        expected = ('rId1', RT_SLIDELAYOUT, slidelayout)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = "expected relationship\n%s\ngot\n%s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__element_minimal_sld_on_construction(self):
        """Slide._element is minimal sld on construction"""
        # setup -----------------------
        path = os.path.join(thisdir, 'test_files/minimal_slide.xml')
        # exercise --------------------
        elm = self.sld._element
        # verify ----------------------
        with open(path, 'r') as f:
            expected = f.read()
        actual = etree.tostring(elm, encoding='UTF-8', pretty_print=True,
                                                         standalone=True)
        msg = "expected:\n%s\n, got\n%s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_slidelayout_property_none_on_construction(self):
        """Slide.slidelayout property None on construction"""
        # verify ----------------------
        self.assertIsProperty(self.sld, 'slidelayout', None)
    
    def test__load_sets_slidelayout(self):
        """Slide._load() sets slidelayout"""
        # setup -----------------------
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        slidelayout = Mock(name='slideLayout')
        slidelayout.partname = '/ppt/slideLayouts/slideLayout1.xml'
        rel = Mock(name='pptx.packaging._Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT_SLIDELAYOUT
        rel.target = slidelayout
        pkgpart = Mock(name='pptx.packaging.Part')
        with open(path, 'rb') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = [rel]
        part_dict = {slidelayout.partname: slidelayout}
        slide = self.sld._load(pkgpart, part_dict)
        # exercise --------------------
        retval = slide.slidelayout
        # verify ----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestSlideCollection(TestCase):
    """Test SlideCollection"""
    def setUp(self):
        prs = Presentation()
        self.slides = SlideCollection(prs)
    
    def test_add_slide_returns_slide(self):
        """SlideCollection.add_slide() returns instance of Slide"""
        # exercise --------------------
        retval = self.slides.add_slide(None)
        # verify ----------------------
        self.assertIsInstance(retval, Slide)
    
    def test_add_slide_sets_slidelayout(self):
        """
        SlideCollection.add_slide() sets Slide.slidelayout
        
        Kind of a throw-away test, but was helpful for initial debugging.
        """
        # setup -----------------------
        slidelayout = Mock(name='slideLayout')
        slidelayout.shapes = []
        slide = self.slides.add_slide(slidelayout)
        # exercise --------------------
        retval = slide.slidelayout
        # verify ----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_add_slide_adds_slide_layout_relationship(self):
        """SlideCollection.add_slide() adds relationship prs->slide"""
        # setup -----------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = SlideLayout()
        slidelayout._shapes = []
        # exercise --------------------
        slide = slides.add_slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(prs._relationships)
        msg = ("expected len(prs._relationships) of %d, got %d"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = prs._relationships[0]
        expected = ('rId1', RT_SLIDE, slide)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = ("expected relationship 1:, got 2:\n1: %s\n2: %s"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
    
    def test_add_slide_sets_partname(self):
        """SlideCollection.add_slide() sets partname of new slide"""
        # setup -----------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = SlideLayout()
        slidelayout._shapes = []
        # exercise --------------------
        slide = slides.add_slide(slidelayout)
        # verify ----------------------
        expected = '/ppt/slides/slide1.xml'
        actual = slide.partname
        msg = "expected partname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestSlideLayout(TestCase):
    """Test SlideLayout"""
    def setUp(self):
        self.slidelayout = SlideLayout()
    
    def __loaded_slidelayout(self, prs_slidemaster=None):
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
        pkg_slidemaster_part = Mock(spec=pptx.packaging.Part)
        pkg_slidemaster_part.partname = sldmaster_partname
        # a package-side relationship from slideLayout to its slideMaster
        rel = Mock(name='pptx.packaging._Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT_SLIDEMASTER
        rel.target = pkg_slidemaster_part
        # the slideLayout package part to send to _load()
        pkg_slidelayout_part = Mock(spec=pptx.packaging.Part)
        pkg_slidelayout_part.relationships = [rel]
        with open(slidelayout_path, 'rb') as f:
            pkg_slidelayout_part.blob = f.read()
        # _load and return
        slidelayout = SlideLayout()
        return slidelayout._load(pkg_slidelayout_part, loaded_part_dict)
    
    def test_class_present(self):
        """SlideLayout class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'SlideLayout')
    
    def test__load_sets_slidemaster(self):
        """SlideLayout._load() sets slidemaster"""
        # setup -----------------------
        prs_slidemaster = Mock(spec=SlideMaster)
        # exercise --------------------
        loaded_slidelayout = self.__loaded_slidelayout(prs_slidemaster)
        # verify ----------------------
        expected = prs_slidemaster
        actual = loaded_slidelayout.slidemaster
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_slidemaster_is_readonly(self):
        """SlideLayout.slidemaster is read-only"""
        # verify ----------------------
        self.assertIsReadOnly(self.slidelayout, 'slidemaster')
    
    def test_slidemaster_raises_on_ref_before_assigned(self):
        """SlideLayout.slidemaster raises on referenced before assigned"""
        with self.assertRaises(AssertionError):
            self.slidelayout.slidemaster
    

class TestSlideMaster(TestCase):
    """Test SlideMaster"""
    def setUp(self):
        self.sldmaster = SlideMaster()
    
    def test_class_present(self):
        """SlideMaster class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'SlideMaster')
    
    def test_slidelayouts_property_empty_on_construction(self):
        """SlideMaster.slidelayouts property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.sldmaster, 'slidelayouts', 0)
    
    def test_slidelayouts_correct_length_after_open(self):
        """SlideMaster.slidelayouts correct length after open"""
        # setup -----------------------
        pkg = Package(test_pptx_path)
        slidemaster = pkg.presentation.slidemasters[0]
        # exercise --------------------
        slidelayouts = slidemaster.slidelayouts
        # verify ----------------------
        self.assertLength(slidelayouts, 11)
    

class TestTextFrame(TestCase):
    """Test TextFrame"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = etree.parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody'
        self.txBodyList = self.sld.xpath(xpath, namespaces=nsmap)
    
    def test_class_present(self):
        """TextFrame class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'TextFrame')
    
    def test_paragraphs_size(self):
        """TextFrame.paragraphs is expected size"""
        # setup -----------------------
        actual_lengths = []
        for txBody in self.txBodyList:
            textframe = TextFrame(txBody)
            # exercise ----------------
            actual_lengths.append(len(textframe.paragraphs))
        # verify ------------------
        expected = [1, 1, 2, 1, 1]
        actual = actual_lengths
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_text_setter_structure_and_value(self):
        """assign to TextFrame.text yields single run para set to value"""
        # setup -----------------------
        test_text = 'python-pptx was here!!'
        txBody = self.txBodyList[2]
        textframe = TextFrame(txBody)
        # exercise --------------------
        textframe.text = test_text
        # verify paragraph count ------
        expected = 1
        actual = len(textframe.paragraphs)
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value ----------------
        expected = test_text
        actual = textframe.paragraphs[0].runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

