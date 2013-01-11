# -*- coding: utf-8 -*-
#
# test_presentation.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.presentation module."""

import inspect
import os

from lxml import etree
from mock import Mock

import pptx.presentation

from pptx.exceptions import InvalidPackageError
from pptx.spec import namespaces, qname
from pptx.spec import (CT_PRESENTATION, CT_SLIDE, CT_SLIDELAYOUT,
    CT_SLIDEMASTER)
from pptx.spec import (RT_HANDOUTMASTER, RT_NOTESMASTER, RT_OFFICEDOCUMENT,
    RT_PRESPROPS, RT_SLIDE, RT_SLIDELAYOUT, RT_SLIDEMASTER, RT_TABLESTYLES,
    RT_THEME, RT_VIEWPROPS)
from pptx.spec import (PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ,
    PH_TYPE_SLDNUM, PH_TYPE_SUBTITLE, PH_TYPE_TBL, PH_TYPE_TITLE,
    PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER)
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
thisdir = os.path.abspath(os.path.split(__file__)[0])
test_pptx_path = os.path.join(thisdir, 'test_files/test.pptx')
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
    shapes = pptx.presentation.ShapeCollection(spTree)
    return shapes


class PartBuilder(object):
    """Builder class for test Parts"""
    def __init__(self):
        self.partname = '/ppt/slides/slide1.xml'
    
    def with_partname(self, partname):
        self.partname = partname
        return self
    
    def build(self):
        p = pptx.presentation.BasePart()
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
            rel = pptx.presentation._Relationship(rId, reltype, target)
            self.relationships.append(rel)
        return self
    
    # def with_singleton_target(self, reltype):
    #     rId = self.__next_rId
    #     partname = self.__singleton_partname(reltype)
    #     target = PartBuilder().with_partname(partname).build()
    #     rel = pptx.presentation._Relationship(rId, reltype, target)
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
        rels = pptx.presentation._RelationshipCollection()
        for rel in self.relationships:
            rels._additem(rel)
        if self.reltype_ordering:
            rels._reltype_ordering = self.reltype_ordering
        return rels
    

class TestBasePart(TestCase):
    """Test pptx.presentation.BasePart"""
    def setUp(self):
        self.basepart = pptx.presentation.BasePart()
        self.cls = pptx.presentation.BasePart
    
    def test_class_present(self):
        """BasePart class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'BasePart')
    
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
    
    def test__content_type_property_raises_on_none(self):
        """BasePart._content_type raises ValueError when None"""
        self.assertPropertyRaisesOnNone(self.basepart, '_content_type')
    
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
    """Test pptx.presentation.BaseShape"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = etree.parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:pic'
        pic = self.sld.xpath(xpath, namespaces=nsmap)[0]
        self.base_shape = pptx.presentation.BaseShape(pic)
    
    def test_class_present(self):
        """BaseShape class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'BaseShape')
    
    def test_has_textframe_value(self):
        """BaseShape.has_textframe value correct"""
        # setup -----------------------
        spTree = self.sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        shapes = pptx.presentation.ShapeCollection(spTree)
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
        base_shape = pptx.presentation.BaseShape(sp)
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
        base_shape = pptx.presentation.BaseShape(title_placeholder_sp)
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
    

class TestBaseSlide(TestCase):
    """Test pptx.presentation.BaseSlide"""
    def setUp(self):
        self.base_slide = pptx.presentation.BaseSlide()
    
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
    """Test pptx.presentation.Collection"""
    def setUp(self):
        self.collection = pptx.presentation.Collection()
    
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
    

class TestPackage(TestCase):
    """Test pptx.presentation.Package"""
    def setUp(self):
        self.pkg = pptx.presentation.Package()
    
    def test_class_present(self):
        """Package class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Package')
    
    def test_open_method_present(self):
        """Package class has method 'open'"""
        self.assertClassHasMethod(pptx.presentation.Package, 'open')
    
    def test_open_returns_self(self):
        """Package.open() returns self-reference"""
        # exercise --------------------
        retval = self.pkg.open(test_pptx_path)
        # verify ----------------------
        expected = self.pkg
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_presentation_property_none_on_construction(self):
        """Package.presentation property None on construction"""
        # verify ----------------------
        self.assertIsProperty(self.pkg, 'presentation', None)
    
    def test_presentation_presentation_after_open(self):
        """Package.presentation Presentation after open()"""
        # setup -----------------------
        cls = pptx.presentation.Presentation
        self.pkg.open(test_pptx_path)
        # exercise --------------------
        obj = self.pkg.presentation
        # verify ----------------------
        actual = isinstance(obj, cls)
        msg = ("expected instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)
    
    def test__relationships_property_empty_on_construction(self):
        """Package._relationships property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.pkg, '_relationships', 0)
    

class TestParagraph(TestCase):
    """Test pptx.presentation.Paragraph"""
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
            paragraph = pptx.presentation.Paragraph(p)
            # exercise ----------------
            actual_lengths.append(len(paragraph.runs))
        # verify ------------------
        expected = [0, 0, 2, 1, 1, 1]
        actual = actual_lengths
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestPart(TestCase):
    """Test pptx.presentation.Part"""
    def test_class_present(self):
        """Part class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Part')
    
    def test_constructs_presentation_for_rt_officedocument(self):
        """Part() returns Presentation for RT_OFFICEDOCUMENT"""
        # setup -----------------------
        cls = pptx.presentation.Presentation
        # exercise --------------------
        obj = pptx.presentation.Part(RT_OFFICEDOCUMENT, CT_PRESENTATION)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_constructs_slide_for_rt_slide(self):
        """Part() returns Slide for RT_SLIDE"""
        # setup -----------------------
        cls = pptx.presentation.Slide
        # exercise --------------------
        obj = pptx.presentation.Part(RT_SLIDE, CT_SLIDE)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_constructs_slidelayout_for_rt_slidelayout(self):
        """Part() returns SlideLayout for RT_SLIDELAYOUT"""
        # setup -----------------------
        cls = pptx.presentation.SlideLayout
        # exercise --------------------
        obj = pptx.presentation.Part(RT_SLIDELAYOUT, CT_SLIDELAYOUT)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_constructs_slidemaster_for_rt_slidemaster(self):
        """Part() returns SlideMaster for RT_SLIDEMASTER"""
        # setup -----------------------
        cls = pptx.presentation.SlideMaster
        # exercise --------------------
        obj = pptx.presentation.Part(RT_SLIDEMASTER, CT_SLIDEMASTER)
        # verify ----------------------
        self.assertIsInstance(obj, cls)
    
    def test_contructor_raises_on_invalid_prs_content_type(self):
        """Part() raises on invalid presentation content type"""
        with self.assertRaises(InvalidPackageError):
            pptx.presentation.Part(RT_OFFICEDOCUMENT, CT_SLIDEMASTER)
    

class TestPartCollection(TestCase):
    """Test pptx.presentation.PartCollection"""
    def test__loadpart_sorts_loaded_parts(self):
        """PartCollection._loadpart sorts loaded parts"""
        # setup -----------------------
        partname1 = '/ppt/slides/slide1.xml'
        partname2 = '/ppt/slides/slide2.xml'
        partname3 = '/ppt/slides/slide3.xml'
        part1 = Mock(name='part1'); part1.partname = partname1
        part2 = Mock(name='part2'); part2.partname = partname2
        part3 = Mock(name='part3'); part3.partname = partname3
        parts = pptx.presentation.PartCollection()
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
    """Test pptx.presentation.Placeholder"""
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
            ph = pptx.presentation.Placeholder(sp)
            values = (ph.type, ph.orient, ph.sz, ph.idx)
            # verify ----------------------
            expected = expected_values[idx]
            actual = values
            msg = "expected shapes[%d] values %s, got %s"\
                   % (idx, expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestPresentation(TestCase):
    """Test pptx.presentation.Presentation"""
    def setUp(self):
        self.prs = pptx.presentation.Presentation()
    
    def test__blob_rewrites_sldIdLst(self):
        """Presentation._blob rewrites sldIdLst"""
        # setup -----------------------
        relationships = RelationshipCollectionBuilder()\
                       .with_tuple_targets(2, RT_SLIDEMASTER)\
                       .with_tuple_targets(3, RT_SLIDE)\
                       .with_ordering(RT_SLIDEMASTER, RT_SLIDE)\
                       .build()
        prs = pptx.presentation.Presentation()
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
        pkg = pptx.presentation.Package().open(test_pptx_path)
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
        pkg = pptx.presentation.Package().open(test_pptx_path)
        prs = pkg.presentation
        # exercise --------------------
        slides = prs.slides
        # verify ----------------------
        self.assertLength(slides, 1)
    

class Test_Relationship(TestCase):
    """Test pptx.presentation._Relationship"""
    def setUp(self):
        rId = 'rId1'
        reltype = RT_SLIDE
        target_part = None
        self.rel = pptx.presentation._Relationship(rId, reltype, target_part)
    
    def test_constructor_raises_on_bad_rId(self):
        """_Relationship constructor raises on non-standard rId"""
        with self.assertRaises(AssertionError):
            pptx.presentation._Relationship('Non-std14', None, None)
    
    def test__num_value(self):
        """_Relationship._num value is correct"""
        # setup -----------------------
        num = 91
        rId = 'rId%d' % num
        rel = pptx.presentation._Relationship(rId, None, None)
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
    """Test pptx.presentation._RelationshipCollection"""
    def setUp(self):
        self.relationships = pptx.presentation._RelationshipCollection()
    
    def test__additem_raises_on_dup_rId(self):
        """_RelationshipCollection._additem raises on duplicate rId"""
        # setup -----------------------
        part1 = pptx.presentation.BasePart()
        part2 = pptx.presentation.BasePart()
        rel1 = pptx.presentation._Relationship('rId9', None, part1)
        rel2 = pptx.presentation._Relationship('rId9', None, part2)
        self.relationships._additem(rel1)
        # verify ----------------------
        with self.assertRaises(ValueError):
            self.relationships._additem(rel2)
    
    def test__additem_maintains_rId_ordering(self):
        """_RelationshipCollection maintains rId ordering on additem()"""
        # setup -----------------------
        part1 = pptx.presentation.BasePart()
        part2 = pptx.presentation.BasePart()
        part3 = pptx.presentation.BasePart()
        rel1 = pptx.presentation._Relationship('rId1', None, part1)
        rel2 = pptx.presentation._Relationship('rId2', None, part2)
        rel3 = pptx.presentation._Relationship('rId3', None, part3)
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
        rel1 = pptx.presentation._Relationship('rId1', RT_SLIDE,       part1)
        rel2 = pptx.presentation._Relationship('rId2', RT_SLIDELAYOUT, part2)
        rel3 = pptx.presentation._Relationship('rId3', RT_SLIDEMASTER, part3)
        rel4 = pptx.presentation._Relationship('rId4', RT_SLIDE,       part4)
        relationships = pptx.presentation._RelationshipCollection()
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
        rel = pptx.presentation._Relationship(rId, RT_SLIDE, part)
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
        part1 = pptx.presentation.BasePart()
        part2 = pptx.presentation.BasePart()
        part3 = pptx.presentation.BasePart()
        part4 = pptx.presentation.BasePart()
        rel1 = pptx.presentation._Relationship('rId1', None, part1)
        rel2 = pptx.presentation._Relationship('rId2', None, part2)
        rel3 = pptx.presentation._Relationship('rId3', None, part3)
        rel4 = pptx.presentation._Relationship('rId4', None, part4)
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
            relationships = pptx.presentation._RelationshipCollection()
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
        rel1 = pptx.presentation._Relationship('rId1', RT_SLIDE, part1)
        rel2 = pptx.presentation._Relationship('rId2', RT_SLIDE, part2)
        relationships = pptx.presentation._RelationshipCollection()
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
    """Test pptx.presentation.Run"""
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
        run = pptx.presentation.Run(self.rList[1])
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
        run = pptx.presentation.Run(self.rList[1])
        # exercise ----------------
        run.text = new_value
        # verify ------------------
        expected = new_value
        actual = run.text
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestShape(TestCase):
    """Test pptx.presentation.Shape"""
    def __loaded_shape(self):
        """
        Return Shape instance loaded from test file.
        """
        sldLayout = _slideLayout1()
        sp = sldLayout.xpath('p:cSld/p:spTree/p:sp', namespaces=nsmap)[0]
        return pptx.presentation.Shape(sp)
    
    def test_class_present(self):
        """Shape class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'Shape')
    

class TestShapeCollection(TestCase):
    """Test pptx.presentation.ShapeCollection"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        sld = etree.parse(path).getroot()
        spTree = sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        self.shapes = pptx.presentation.ShapeCollection(spTree)
    
    def test_size_after_construction(self):
        """ShapeCollection is expected size after construction"""
        # verify ----------------------
        self.assertLength(self.shapes, 9)
    
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
        slidelayout = pptx.presentation.SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        shapes = pptx.presentation.ShapeCollection(_empty_spTree())
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
            ph = pptx.presentation.Placeholder(sp)
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
        shapes = pptx.presentation.ShapeCollection(_empty_spTree())
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
            layout_ph = pptx.presentation.Placeholder(layout_ph_sp)
            sp = shapes._ShapeCollection__clone_layout_placeholder(layout_ph)
            # verify ----------------------
            ph = pptx.presentation.Placeholder(sp)
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
    

class TestSlide(TestCase):
    """Test pptx.presentation.Slide"""
    def setUp(self):
        self.sld = pptx.presentation.Slide()
    
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
        with open(path, 'r') as f:
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
    """Test pptx.presentation.SlideCollection"""
    def setUp(self):
        self.slides = pptx.presentation.SlideCollection()
    
    def test_add_slide_returns_slide(self):
        """SlideCollection.add_slide() returns instance of Slide"""
        # exercise --------------------
        retval = self.slides.add_slide(None)
        # verify ----------------------
        self.assertIsInstance(retval, pptx.presentation.Slide)
    
    def test_add_slide_sets_slidelayout(self):
        """SlideCollection.add_slide() sets Slide.slidelayout"""
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
    

class TestSlideLayout(TestCase):
    """Test pptx.presentation.SlideLayout"""
    def setUp(self):
        self.slidelayout = pptx.presentation.SlideLayout()
    
    def __loaded_slidelayout(self, prs_slidemaster=None):
        """
        Return SlideLayout instance loaded using mocks. *prs_slidemaster* is
        an already-loaded model-side SlideMaster instance (or mock, as
        appropriate to calling test).
        """
        # partname for related slideMaster
        sldmaster_partname = '/ppt/slideMasters/slideMaster1.xml'
        # path to test slideLayout XML
        relpath = ('../pptx/pptx_template/default/ppt/slideLayouts/'
                   'slideLayout1.xml')
        slidelayout_path = os.path.abspath(os.path.join(thisdir, relpath))
        # model-side slideMaster part
        if prs_slidemaster is None:
            prs_slidemaster = Mock(spec=pptx.presentation.SlideMaster)
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
        with open(slidelayout_path, 'r') as f:
            pkg_slidelayout_part.blob = f.read()
        # _load and return
        slidelayout = pptx.presentation.SlideLayout()
        return slidelayout._load(pkg_slidelayout_part, loaded_part_dict)
    
    def test_class_present(self):
        """SlideLayout class present in presentation module"""
        self.assertClassInModule(pptx.presentation, 'SlideLayout')
    
    def test__load_sets_slidemaster(self):
        """SlideLayout._load() sets slidemaster"""
        # setup -----------------------
        prs_slidemaster = Mock(spec=pptx.presentation.SlideMaster)
        # # exercise --------------------
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
    """Test pptx.presentation.SlideMaster"""
    def setUp(self):
        self.sldmaster = pptx.presentation.SlideMaster()
    
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
        pkg = pptx.presentation.Package().open(test_pptx_path)
        slidemaster = pkg.presentation.slidemasters[0]
        # exercise --------------------
        slidelayouts = slidemaster.slidelayouts
        # verify ----------------------
        self.assertLength(slidelayouts, 11)
    

class TestTextFrame(TestCase):
    """Test pptx.presentation.TextFrame"""
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
            textframe = pptx.presentation.TextFrame(txBody)
            # exercise ----------------
            actual_lengths.append(len(textframe.paragraphs))
        # verify ------------------
        expected = [1, 1, 2, 1, 1]
        actual = actual_lengths
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

