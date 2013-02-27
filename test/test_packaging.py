# -*- coding: utf-8 -*-
#
# test_packaging.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.packaging module."""

import inspect
import os
import sys
import unittest2
import zipfile

from collections import namedtuple
from lxml import etree
from mock import Mock
from StringIO import StringIO
from unittest2 import skipIf

from .context import pptx

import pptx.packaging
import pptx.presentation

from pptx.exceptions import CorruptedPackageError, DuplicateKeyError,\
                            NotXMLError, PackageNotFoundError
from pptx.spec import PTS_CARDINALITY_TUPLE, PTS_HASRELS_ALWAYS

from testing import TestCase

# module globals -------------------------------------------------------------
def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir = os.path.split(__file__)[0]
test_file_dir = absjoin(thisdir, 'test_files')

test_pptx_path = absjoin(test_file_dir, 'test.pptx')
dir_pkg_path   = absjoin(test_file_dir, 'expanded_pptx')
zip_pkg_path   = test_pptx_path

class MockParent(object):
    """Stub out parent attributes."""
    def __init__(self, baseURI=None, itemURI=None):
        self.baseURI = baseURI
        self.itemURI = itemURI
    


# ============================================================================
# Test Classes
# ============================================================================

class TestBaseFileSystem(TestCase):
    """Test pptx.packaging.FileSystem"""
    def test_getblob_raises_on_bad_itemuri(self):
        """BaseFileSystem.getblob(itemURI) raises on bad itemURI"""
        # setup -----------------------
        bad_itemURI = '/spam/eggs/egg1.xml'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        # verify ----------------------
        with self.assertRaises(LookupError):
            fs.getblob(bad_itemURI)
    
    def test_getelement_raises_on_bad_itemuri(self):
        """BaseFileSystem.getelement(itemURI) raises on bad itemURI"""
        # setup -----------------------
        bad_itemURI = '/spam/eggs/egg1.xml'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        # verify ----------------------
        with self.assertRaises(LookupError):
            fs.getelement(bad_itemURI)
    

class Test_ContentTypesItem(TestCase):
    """Test pptx.packaging._ContentTypesItem"""
    def setUp(self):
        self.cti = pptx.packaging._ContentTypesItem()
    
    def test_class_present(self):
        """_ContentTypesItem class present in packaging module"""
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, '_ContentTypesItem')
    
    def test_compose_method_present(self):
        """
        _ContentTypesItem class has method 'compose'.
        
        NOTE: This test will fail if method raises an exception.
        """
        self.assertClassHasMethod(pptx.packaging._ContentTypesItem, 'compose')
    
    def test_compose_returns_self(self):
        """_ContentTypesItem.compose() returns self-reference"""
        # setup -----------------------
        pkg = pptx.packaging.Package().open(zip_pkg_path)
        # exercise --------------------
        retval = self.cti.compose(pkg.parts)
        # verify ----------------------
        expected = self.cti
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_compose_correct_count(self):
        """_ContentTypesItem.compose() produces expected element count"""
        # setup -----------------------
        pkg = pptx.packaging.Package().open(zip_pkg_path)
        # exercise --------------------
        self.cti.compose(pkg.parts)
        # verify ----------------------
        self.assertLength(self.cti, 24)
    
    def test_compose_raises_on_bad_partname_ext(self):
        """_ContentTypesItem.compose() raises on bad partname extension"""
        # setup -----------------------
        MockPart = namedtuple('MockPart', 'partname')
        parts = [MockPart('/ppt/!blat/rhumba.1x&')]
        # verify ----------------------
        with self.assertRaises(LookupError):
            self.cti.compose(parts)
    
    def test_element_attribute_present(self):
        """_ContentTypesItem instance has attribute 'element'"""
        self.assertInstHasAttr(self.cti, 'element')
    
    def test_element_correct_length(self):
        """_ContentTypesItem.element() has expected element count"""
        # setup -----------------------
        pkg = pptx.packaging.Package().open(zip_pkg_path)
        # exercise --------------------
        self.cti.compose(pkg.parts)
        # verify ----------------------
        self.assertLength(self.cti.element, 24)
    
    def test_getitem_method_present(self):
        """_ContentTypesItem class has method '__getitem__'"""
        self.assertClassHasMethod(pptx.packaging._ContentTypesItem,
                                  '__getitem__')
    
    def test_getitem_raises_before_load(self):
        """_ContentTypesItem[partname] raises before load"""
        # verify ----------------------
        with self.assertRaises(ValueError):
            self.cti['/ppt/presentation.xml']
    
    def test_getitem_raises_on_bad_partname(self):
        """_ContentTypesItem[partname] raises on bad partname"""
        # setup ------------------------
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        self.cti.load(fs)
        # verify ----------------------
        with self.assertRaises(LookupError):
            self.cti['!blat/rhumba.1x&']
    
    def test_load_spotcheck(self):
        """_ContentTypesItem can load itself from a filesystem instance"""
        # setup ------------------------
        dir_fs = pptx.packaging.FileSystem(dir_pkg_path)
        zip_fs = pptx.packaging.FileSystem(zip_pkg_path)
        file_fs = pptx.packaging.FileSystem(open(zip_pkg_path))
        for fs in (dir_fs, zip_fs, file_fs):
            # exercise ---------------------
            cti = pptx.packaging._ContentTypesItem().load(fs)
            # test -------------------------
            expected = 'application/vnd.openxmlformats-officedocument'\
                       '.presentationml.slideLayout+xml'
            actual = cti['/ppt/slideLayouts/slideLayout1.xml']
            msg = "expected content type '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestDirectoryFileSystem(TestCase):
    """Test pptx.packaging.DirectoryFileSystem"""
    def test_class_present(self):
        """DirectoryFileSystem class present in packaging module"""
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'DirectoryFileSystem')
    

class TestFileSystem(TestCase):
    """Test pptx.packaging.FileSystem"""
    def setUp(self):
        self.dir_fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.zip_fs = pptx.packaging.FileSystem(zip_pkg_path)
        self.file_fs = pptx.packaging.FileSystem(open(zip_pkg_path))
    
    def test_class_present(self):
        """FileSystem class present in packaging module"""
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'FileSystem')
    
    def test_construction(self):
        """FileSystem factory returns package-appropriate class"""
        test_datasets =\
            ( (self.dir_fs, 'DirectoryFileSystem')
            , (self.zip_fs, 'ZipFileSystem')
            , (self.file_fs, 'ZipFileSystem')
            )
        for fs, expected in test_datasets:
            actual = fs.__class__.__name__
            msg = "expected class name '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_constructor_raises_on_bad_path(self):
        """FileSystem(path) constructor raises on bad path"""
        # setup -----------------------
        bad_path = 'blat/rhumba.1x&'
        # verify ----------------------
        with self.assertRaises(PackageNotFoundError):
            pptx.packaging.FileSystem(bad_path)
    
    def test_constructor_raises_on_bad_file(self):
        """FileSystem(path) constructor raises on bad file"""
        # setup -----------------------
        bad_file = StringIO('not a zip file')
        # verify ----------------------
        with self.assertRaises(PackageNotFoundError):
            pptx.packaging.FileSystem(bad_file)
    
    def test_contains(self):
        """'in' operator returns True if URI is in filesystem"""
        expected_URIs =\
            ( '/[Content_Types].xml'
            , '/docProps/app.xml'
            , '/ppt/presentation.xml'
            , '/ppt/slideMasters/slideMaster1.xml'
            , '/ppt/slideLayouts/_rels/slideLayout1.xml.rels'
            )
        for fs in (self.dir_fs, self.zip_fs, self.file_fs):
            for uri in expected_URIs:
                self.assertIn(uri, fs)
    
    def test_getelement_return_count(self):
        """ElementTree element for specified package item is returned"""
        for fs in (self.dir_fs, self.zip_fs, self.file_fs):
            elm = fs.getelement('/[Content_Types].xml')
            self.assertLength(elm, 24)
    
    def test_getelement_raises_on_binary(self):
        """Calling getelement() for binary item raises exception"""
        # call getelement for thumbnail
        for fs in (self.dir_fs, self.zip_fs, self.file_fs):
            with self.assertRaises(NotXMLError):
                fs.getelement('/docProps/thumbnail.jpeg')
    
    def test_getstream_correct_length(self):
        """StringIO instance for specified package item is returned"""
        for fs in (self.dir_fs, self.zip_fs, self.file_fs):
            stream = fs.getstream('/[Content_Types].xml')
            elm = etree.parse(stream).getroot()
            self.assertLength(elm, 24)
    
    def test_getstream_raises_on_bad_URI(self):
        """FileSystem.getstream() raises on bad URI"""
        for fs in (self.dir_fs, self.zip_fs, self.file_fs):
            with self.assertRaises(LookupError):
                fs.getstream('!blat/rhumba.xml')
    
    def test_itemURIs_count(self):
        """FileSystem.itemURIs has expected count"""
        # verify ----------------------
        for fs, fsname in ((self.dir_fs, 'dir_fs'),
                           (self.zip_fs, 'zip_fs'),
                           (self.file_fs, 'file_fs')):
            self.assertLength(fs.itemURIs, 38)
    
    def test_itemURIs_plausible(self):
        """All URIs in FileSystem.itemURIs are plausible"""
        # verify ----------------------
        for fs in (self.dir_fs, self.zip_fs, self.file_fs):
            for itemURI in fs.itemURIs:
                # plausible segment count
                expected_min = 1
                expected_max = 4
                segment_count = len(itemURI.split('/'))-1  # leading slash produces empty string in split list
                msg = "item URI has implausible number of segments:\nitemURI ==> '%s'" % (itemURI)
                self.assertTrue(segment_count >= expected_min, msg)
                self.assertTrue(segment_count <= expected_max, msg)
                # check for forward slash
                msg = "item URI '%s' does not start with forward slash ('/')" % (itemURI)
                self.assertTrue(itemURI.startswith('/'), msg)
    

class TestPackage(TestCase):
    """Test pptx.packaging.Package"""
    def setUp(self):
        self.pkg = pptx.packaging.Package()
        self.test_pptx_path = absjoin(thisdir, 'test_python-pptx.pptx')
    
    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)
    
    def test_marshal_method_present(self):
        """Package class has method 'marshal'"""
        self.assertClassHasMethod(pptx.packaging.Package, 'marshal')
    
    def test_marshal_returns_self(self):
        """Package.marshal() returns self-reference"""
        # setup -----------------------
        model_pkg = pptx.presentation.Package(test_pptx_path)
        # exercise --------------------
        retval = self.pkg.marshal(model_pkg)
        # verify ----------------------
        expected = self.pkg
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_open_returns_self(self):
        """Package.open() returns self-reference"""
        for path in (dir_pkg_path, zip_pkg_path, open(zip_pkg_path)):
            # exercise ----------------
            retval = self.pkg.open(path)
            # verify ------------------
            expected = self.pkg
            actual = retval
            msg = "expected '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_open_part_count(self):
        """Package.open() produces expected part count"""
        # exercise --------------------
        self.pkg.open(zip_pkg_path)
        # verify ----------------------
        self.assertLength(self.pkg.parts, 22)
    
    def test_open_populates_target_part(self):
        """Part.open() populates Relationship.target"""
        # setup -----------------------
        cls = pptx.packaging.Part
        # exercise --------------------
        self.pkg.open(zip_pkg_path)
        # verify ----------------------
        for rel in self.pkg.relationships:
            obj = rel.target
            self.assertIsInstance(obj, cls)
        for part in self.pkg.parts:
            for rel in part.relationships:
                obj = rel.target
                self.assertIsInstance(obj, cls)
    
    # def test_blob_correct_length_after_load_binary_part(self):
    #     """Part.blob correct length after load binary part"""
    #     # setup -----------------------
    #     partname = '/docProps/thumbnail.jpeg'
    #     content_type = 'image/jpeg'
    #     self.part.load(self.fs, partname, content_type)
    #     # exercise --------------------
    #     blob = self.part.blob
    #     # verify ----------------------
    #     self.assertLength(blob, 8147)
    # 
    # def test__rels_element_correct_xml(self):
    #     BOTH PKG AND PARTS, MAYBE USE TEST FILES
    #     """RelationshipsItem.element produces expected XML"""
    #     # setup -----------------------
    #     ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    #     exp_xml = '<Relationships xmlns="%s"/>' % ns
    #     # exercise --------------------
    #     elm = self.ri.element
    #     # verify ----------------------
    #     xml_out = etree.tostring(elm)
    #     expected = exp_xml
    #     actual = xml_out
    #     msg = "expected '%s'\n%sgot '%s'" % (expected, ' '*21, actual)
    #     self.assertEqual(expected, actual, msg)
    # 
    # def test_typespec_correct_after_load(self):
    #     """Part.typespec is correct after load"""
    #     # setup -----------------------
    #     self.part.load(self.fs, self.partname, self.content_type)
    #     # exercise --------------------
    #     typespec = self.part.typespec
    #     test_cases =\
    #         ( (typespec.basename    , 'slideMaster')
    #         , (typespec.cardinality , 'multiple'   )
    #         , (typespec.has_rels    , 'always'     )
    #         , (typespec.format      , 'xml'        )
    #         )
    #     # verify ----------------------
    #     for actual, expected in test_cases:
    #         msg = "expected '%s', got '%s'" % (expected, actual)
    #         self.assertEqual(expected, actual, msg)
    # 
    def test_parts_property_empty_on_construction(self):
        """Package.parts property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.pkg, 'parts', 0)
    
    def test_relationships_property_empty_on_construction(self):
        """Package.relationships property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.pkg, 'relationships', 0)
    
    def test_relationships_correct_length_after_open(self):
        """Package.relationships correct length after open()"""
        # exercise --------------------
        self.pkg.open(zip_pkg_path)
        # verify ----------------------
        self.assertLength(self.pkg.relationships, 4)
    
    def test_relationships_discarded_before_open(self):
        """Package.relationships discards rels before second open()"""
        # setup -----------------------
        self.pkg.open(zip_pkg_path)
        # exercise --------------------
        self.pkg.open(dir_pkg_path)
        # verify ----------------------
        self.assertLength(self.pkg.relationships, 4)
    
    def test_save_writes_pptx_zipfile(self):
        """Package.save() writes .pptx file"""
        # setup -----------------------
        self.pkg.open(dir_pkg_path)
        save_path = self.test_pptx_path
        # exercise --------------------
        self.pkg.save(save_path)
        # verify ----------------------
        actual = zipfile.is_zipfile(save_path)
        msg = "no zipfile at %s" % (save_path)
        self.assertTrue(actual, msg)
    
    def test_save_adds_pptx_ext(self):
        """Package.save() adds .pptx extension if not supplied"""
        # setup -----------------------
        self.pkg.open(dir_pkg_path)
        save_path = self.test_pptx_path[:-5]  # trim off .pptx extension
        exp_path = self.test_pptx_path
        # exercise --------------------
        self.pkg.save(save_path)
        # verify ----------------------
        actual = zipfile.is_zipfile(exp_path)
        msg = "no zipfile at %s" % (exp_path)
        self.assertTrue(actual, msg)
    
    def test_save_writes_file_object(self):
        """Package.save() can write to a file-like object"""
        # setup -----------------------
        self.pkg.open(dir_pkg_path)
        save_file = StringIO()
        # exercise --------------------
        self.pkg.save(save_file)
        # verify ----------------------
        actual = zipfile.is_zipfile(save_file)
        msg = "no zipfile created"
        self.assertTrue(actual, msg)
    
    def test_save_member_count(self):
        """Package.save() produces expected zip member count"""
        # setup -----------------------
        self.pkg.open(dir_pkg_path)
        save_path = self.test_pptx_path
        # exercise --------------------
        self.pkg.save(save_path)
        # verify ----------------------
        zip = zipfile.ZipFile(save_path)
        names = zip.namelist()
        zip.close()
        partnames = [name for name in names if not name.endswith('/')]
        self.assertLength(partnames, 38)
    

class TestPart(TestCase):
    """Test pptx.packaging.Part"""
    def setUp(self):
        # create Part instance and ancestors
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.partname = '/ppt/slideMasters/slideMaster1.xml'
        self.content_type = 'application/vnd.openxmlformats-officedocument'\
                            '.presentationml.slideMaster+xml'
        self.part = pptx.packaging.Part()
        self.test_pptx_path = absjoin(thisdir, 'test_python-pptx.pptx')
    
    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)
    
    def test_blob_none_on_construction(self):
        """Part.blob is None on construction"""
        expected = None
        actual = self.part.blob
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_content_type_correct(self):
        """Part.content_type returns correct value."""
        # setup -----------------------
        self.part.typespec = Mock()
        self.part.typespec.content_type = self.content_type
        # exercise --------------------
        retval = self.part.content_type
        # verify ----------------------
        expected = self.content_type
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test__load_raises_on_missing_rels_item(self):
        """Part._load() raises on missing rels item"""
        # setup -----------------------
        path = absjoin(test_file_dir, 'missing_rels_item.pptx')
        fs = pptx.packaging.FileSystem(path)
        pkg = pptx.packaging.Package()
        # verify ----------------------
        with self.assertRaises(CorruptedPackageError):
            pkg.open(path)
    
    def test__marshal_method_present(self):
        """Part class has method '_marshal'"""
        self.assertClassHasMethod(pptx.packaging.Part, '_marshal')
    
    def test_partname_property_none_on_construction(self):
        """Part.partname property None on construction"""
        # verify ----------------------
        self.assertIsProperty(self.part, 'partname', None)
    
    def test_relationships_property_empty_on_construction(self):
        """Part.relationships property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.part, 'relationships', 0)
    
    def test_typespec_attribute_present(self):
        """Part instance has attribute 'typespec'"""
        self.assertInstHasAttr(self.part, 'typespec')
    
    def test_typespec_none_on_construction(self):
        """Part.typespec is None on construction"""
        expected = None
        actual = self.part.typespec
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestPartTypeSpec(TestCase):
    """Test pptx.packaging.PartTypeSpec"""
    def test_class_present(self):
        """PartTypeSpec class present in packaging module"""
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'PartTypeSpec')
    
    def test_construction_returns_correct_parttypespec(self):
        """PartTypeSpec(content_type) returns correct PartTypeSpec"""
        # setup -----------------------
        content_type = 'application/vnd.openxmlformats-officedocument.'\
                       'presentationml.notesMaster+xml'
        # exercise --------------------
        pts = pptx.packaging.PartTypeSpec(content_type)
        # verify ----------------------
        actual = (pts.basename    == 'notesMaster'
              and pts.ext         == '.xml'
              and pts.cardinality == PTS_CARDINALITY_TUPLE
              and pts.required    == False
              and pts.baseURI     == '/ppt/notesMasters'
              and pts.has_rels    == PTS_HASRELS_ALWAYS
              and pts.reltype     == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster')
        msg = "PartTypeSpec('%s') returns unexpected values" % (content_type)
        self.assertTrue(actual, msg)
    
    def test_construction_raises_on_bad_contenttype(self):
        """PartTypeSpec(ct) raises exception on unrecognized content type"""
        # setup -----------------------
        content_type = 'application/vnd.baloneyMaster+xml'
        # verify ----------------------
        with self.assertRaises(KeyError):
            pptx.packaging.PartTypeSpec(content_type)
    
    def test_format_correct(self):
        """PartTypeSpec.format returns correct value"""
        # setup -----------------------
        ct_slideMaster = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.slideLayout+xml'
        ct_jpeg = 'image/jpeg'
        ct_printerSettings = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.printerSettings'
        ct_slide = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.slide+xml'
        test_cases =\
            ( ( ct_slideMaster     , 'xml'    )
            , ( ct_jpeg            , 'binary' )
            , ( ct_printerSettings , 'binary' )
            , ( ct_slide           , 'xml'    )
            )
        # exercise --------------------
        for ct, exp_format in test_cases:
            pts = pptx.packaging.PartTypeSpec(ct)
            # verify ----------------------
            expected = exp_format
            actual = pts.format
            msg = "expected '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestRelationship(TestCase):
    """Test pptx.packaging.Relationship"""
    def setUp(self):
        """Create a new relationship from a string"""
        self.baseURI = '/'
        self.rId = 'rId1'
        self.reltype = 'http://schemas.openxmlformats.org/officeDocument/'\
                       '2006/relationships/officeDocument'
        self.target = 'ppt/presentation.xml'
        tmpl = '<Relationship Id="%s" Type="%s" Target="%s"/>'
        self.rel_xml = tmpl % (self.rId, self.reltype, self.target)
        self.rel_elm = etree.fromstring(self.rel_xml)
        # self.rel = pptx.packaging.Relationship()._load(self.baseURI,
        #                                                self.rel_elm)
    
    def test_class_present(self):
        """Relationship class present in packaging module"""
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'Relationship')
    
    # def test_construction_correct_attr_values(self):
    #     """Relationship attributes loaded from ElementTree.Element"""
    #     # preconditions ---------------
    #     self.assertInstHasAttr(self.rel, 'rId')
    #     self.assertInstHasAttr(self.rel, 'reltype')
    #     self.assertInstHasAttr(self.rel, 'target')
    #     # exercise --------------------
    #     rel = self.rel._load(self.baseURI, self.rel_elm)
    #     # verify ----------------------
    #     expected = (self.rId, self.reltype, self.target)
    #     actual = (rel.rId, rel.reltype)
    #     msg = "expected '%s', got '%s'" % (expected, actual)
    #     self.assertEqual(expected, actual, msg)
    # 
    # @skipIf(not hasattr(self.rel, 'element')
    # def test_element_correct_xml(self):
    #     """Relationship.element produces expected XML"""
    #     # precondition ----------------
    #     self.assertInstHasAttr(self.rel, 'element')
    #     # exercise --------------------
    #     elm_out = self.rel.element
    #     # verify ----------------------
    #     xml_out = etree.tostring(elm_out)
    #     expected = self.rel_xml
    #     actual = xml_out
    #     msg = "expected '%s'\n%sgot '%s'" % (expected, ' '*21, actual)
    #     self.assertEqual(expected, actual, msg)
    # 

class TestZipFileSystem(TestCase):
    """Test pptx.packaging.ZipFileSystem (writing aspect)"""
    def setUp(self):
        self.test_pptx_path = absjoin(test_file_dir, 'test_python-pptx.pptx')
        self.xml_in =\
            """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"""\
            """<p:presentationPr xmlns:a="http://main" """\
            """xmlns:r="http://relationships" """\
            """xmlns:p="http://presentationml">"""\
            """<p:extLst>"""\
            """<p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}">"""\
            """<r:discardImageEditData val="0"/>"""\
            """</p:ext>"""\
            """</p:extLst>"""\
            """</p:presentationPr>"""
        self.xml_out =\
            """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"""\
            """<p:presentationPr xmlns:a="http://main"\n"""\
            """                  xmlns:r="http://relationships"\n"""\
            """                  xmlns:p="http://presentationml">\n"""\
            """  <p:extLst>\n"""\
            """    <p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}">\n"""\
            """      <r:discardImageEditData val="0"/>\n"""\
            """    </p:ext>\n"""\
            """  </p:extLst>\n"""\
            """</p:presentationPr>"""
    
    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)
    
    def test_getblob_correct_length(self):
        """ZipFileSystem.getblob() returns object of expected length"""
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        # exercise --------------------
        blob = fs.getblob(partname)
        # verify ----------------------
        self.assertLength(blob, 8147)
    
    def test_getblob_raises_on_bad_URI(self):
        """ZipFileSystem.getblob() raises on bad URI"""
        # setup -----------------------
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        # verify ----------------------
        with self.assertRaises(LookupError):
            fs.getstream('!blat/rhumba.xml')
    
    def test_write_blob_round_trips(self):
        """ZipFileSystem.write_blob() round-trips intact"""
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        in_blob = fs.getblob(partname)
        test_fs = pptx.packaging.ZipFileSystem(self.test_pptx_path, 'w')
        # exercise --------------------
        test_fs.write_blob(in_blob, partname)
        # verify ----------------------
        out_blob = test_fs.getblob(partname)
        expected = in_blob
        actual = out_blob
        msg = "retrived blob (len %d) differs from original (len %d)"\
               % (len(actual), len(expected))
        self.assertEqual(expected, actual, msg)
    
    def test_write_blob_raises_on_dup_itemuri(self):
        """ZipFileSystem.write_blob() raises on duplicate itemURI"""
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        blob = fs.getblob(partname)
        test_fs = pptx.packaging.ZipFileSystem(self.test_pptx_path, 'w')
        test_fs.write_blob(blob, partname)
        # verify ----------------------
        with self.assertRaises(DuplicateKeyError):
            test_fs.write_blob(blob, partname)
    
    def test_write_element_round_trips(self):
        """ZipFileSystem.write_element() round-trips intact"""
        # setup -----------------------
        elm = etree.fromstring(self.xml_in)
        itemURI = '/ppt/test.xml'
        zipfs = pptx.packaging.ZipFileSystem(self.test_pptx_path, 'w')
        # exercise --------------------
        zipfs.write_element(elm, itemURI)
        # verify ----------------------
        stream = zipfs.getstream(itemURI)
        xml_out = stream.read()
        stream.close()
        expected = self.xml_out
        actual = xml_out
        msg = "expected \n%s\n, got\n%s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_write_element_raises_on_dup_itemuri(self):
        """ZipFileSystem.write_element() raises on duplicate itemURI"""
        # setup -----------------------
        elm = etree.fromstring(self.xml_in)
        itemURI = '/ppt/test.xml'
        zipfs = pptx.packaging.ZipFileSystem(self.test_pptx_path, 'w')
        # exercise --------------------
        zipfs.write_element(elm, itemURI)
        # verify ----------------------
        with self.assertRaises(DuplicateKeyError):
            zipfs.write_element(elm, itemURI)
    


